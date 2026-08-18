from __future__ import annotations

import os
import re
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
from flask import Flask, jsonify, render_template, request
from sqlalchemy.exc import SQLAlchemyError

from db import (
    AppMetaDB,
    BlockClosureDB,
    CycleDefinitionDB,
    MassAdjustmentDB,
    RowAdjustmentDB,
    SessionLocal,
    TpsrRecord,
    WeekAdjustmentDB,
    bump_data_version,
    format_short_week,
    get_data_version,
    normalize_text,
    parse_week_code,
    process_tpsr_excel_upload,
    seed_data_from_excel_if_empty,
)

BASE_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = BASE_DIR / "siembras_podas" / "tpsr.xlsx"

app = Flask(__name__)

# Automatically create tables and seed Postgres if empty on startup
seed_data_from_excel_if_empty(WORKBOOK_PATH)


def shift_week(week_code: int, delta_weeks: int) -> int:
    year = week_code // 100
    week = week_code % 100
    monday = date.fromisocalendar(year, week, 1)
    shifted = monday + timedelta(weeks=delta_weeks)
    iso_year, iso_week, _ = shifted.isocalendar()
    return iso_year * 100 + iso_week


def build_week_window_with_lookback(center_week: int, weeks_back: int, weeks_forward: int) -> list[int]:
    return [shift_week(center_week, offset) for offset in range(-weeks_back, weeks_forward + 1)]


def current_week_code() -> int:
    today = date.today()
    iso_year, iso_week, _ = today.isocalendar()
    return iso_year * 100 + iso_week


def float_or_zero(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    try:
        return float(value)
    except ValueError:
        return 0.0


def to_int_or_none(value: Any) -> int | None:
    if value in (None, ""):
        return None
    try:
        return int(round(float(value)))
    except (ValueError, TypeError):
        return None


def to_float_or_none(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def normalize_optional_text(value: Any, fallback: str = "*") -> str:
    text = normalize_text(value)
    return text if text else fallback


# In-memory snapshot cache synchronized across Gunicorn workers via PostgreSQL data_version
_WORKER_CACHE: dict[str, Any] = {"version": -1, "snapshot": None}


def invalidate_snapshot_cache() -> None:
    bump_data_version()
    _WORKER_CACHE["snapshot"] = None
    _WORKER_CACHE["version"] = -1


def build_projection_snapshot_from_db() -> dict[str, Any]:
    db = SessionLocal()
    try:
        tpsr_recs = db.query(TpsrRecord).filter(TpsrRecord.plants > 0).all()
        cycle_defs = db.query(CycleDefinitionDB).all()
        row_adjs = db.query(RowAdjustmentDB).all()
        mass_adjs = db.query(MassAdjustmentDB).all()
        week_adjs = db.query(WeekAdjustmentDB).all()
        block_closes = db.query(BlockClosureDB).all()

        cycle_map = {}
        for c in cycle_defs:
            curve_values = tuple(
                float(val)
                for val in (c.curve or "1.0").split(",")
                if val.strip() and float_or_zero(val) > 0
            ) or (1.0,)
            cycle_map[(c.product_master_norm, c.variety_norm, c.activity)] = {
                "product_master": c.product_master,
                "variety": c.variety,
                "activity": c.activity,
                "cycle_weeks": c.cycle_weeks,
                "waste_rate": c.waste_rate,
                "stems_per_plant": c.stems_per_plant,
                "curve": curve_values,
            }

        row_adjustments = {
            (r.activity, r.product_master_norm, r.variety_norm, r.block_norm, r.source_week): {
                "cycle_weeks": r.cycle_weeks,
                "waste_rate": r.waste_rate,
                "stems_per_plant": r.stems_per_plant,
            }
            for r in row_adjs
        }

        mass_adjustments = [
            {
                "product_master_norm": m.product_master_norm,
                "variety_norm": m.variety_norm,
                "activity": m.activity,
                "cycle_weeks": m.cycle_weeks,
                "waste_rate": m.waste_rate,
                "stems_per_plant": m.stems_per_plant,
            }
            for m in mass_adjs
        ]

        week_adjustments = {
            (
                w.activity,
                w.product_master_norm,
                w.variety_norm,
                w.block_norm,
                w.source_week,
                w.harvest_week,
            ): {
                "agronomo_estimate": w.agronomo_estimate,
                "real_closed": w.real_closed,
                "is_dump": bool(w.is_dump),
            }
            for w in week_adjs
        }

        block_closures = {
            (b.activity, b.product_master_norm, b.variety_norm, b.block_norm, b.source_week): b.is_closed
            for b in block_closes
        }

        # Consolidar registros TPSR a nivel de bloque único (activity, product_master, variety, block, source_week)
        consolidated_dict: dict[tuple[str, str, str, str, int], dict[str, Any]] = {}
        for rec in tpsr_recs:
            activity = rec.activity
            source_week = rec.source_week
            plants = rec.plants
            product_master = rec.product_master
            product_master_norm = rec.product_master_norm
            variety = rec.variety
            variety_norm = rec.variety_norm
            block = rec.block
            block_norm = rec.block_norm

            if "SPLASH" in variety_norm or (product_master and product_master.upper() == "VERONICA SPRAY"):
                product_master = "VERONICA SPRAY"
                product_master_norm = "VERONICA SPRAY"

            key = (activity, product_master_norm, variety_norm, block_norm, source_week)
            if key not in consolidated_dict:
                consolidated_dict[key] = {
                    "activity": activity,
                    "source_week": source_week,
                    "plants": plants,
                    "product_master": product_master,
                    "product_master_norm": product_master_norm,
                    "variety": variety,
                    "variety_norm": variety_norm,
                    "block": block,
                    "block_norm": block_norm,
                }
            else:
                consolidated_dict[key]["plants"] += plants

        source_rows = []
        projection_rows = []
        missing_cycles: dict[tuple[str, str, str], int] = {}

        for item in consolidated_dict.values():
            activity = item["activity"]
            source_week = item["source_week"]
            plants = item["plants"]
            product_master = item["product_master"]
            product_master_norm = item["product_master_norm"]
            variety = item["variety"]
            variety_norm = item["variety_norm"]
            block = item["block"]
            block_norm = item["block_norm"]

            cycle_key = (product_master_norm, variety_norm, activity)
            cycle = cycle_map.get(cycle_key)
            if cycle is None:
                for (p_norm, v_norm, act), cdef in cycle_map.items():
                    if v_norm == variety_norm and act == activity:
                        cycle = cdef
                        break

            source_rows.append(
                {
                    "activity": activity,
                    "product_master": product_master,
                    "product_master_norm": product_master_norm,
                    "variety": variety,
                    "source_week": source_week,
                    "source_week_short": format_short_week(source_week),
                    "block": block,
                    "plants": plants,
                    "bed_location": "",
                    "pruning_number": "",
                    "cycle_found": cycle is not None,
                    "cycle_weeks": None,
                    "waste_rate": None,
                    "stems_per_plant": None,
                    "ideal_cycle_weeks": cycle["cycle_weeks"] if cycle else None,
                    "ideal_waste_rate": cycle["waste_rate"] if cycle else None,
                    "ideal_stems_per_plant": cycle["stems_per_plant"] if cycle else None,
                    "curve_csv": ",".join(str(v) for v in cycle["curve"]) if cycle else "",
                    "program_total": 0,
                    "block_closed": False,
                }
            )

            if cycle is None:
                missing_cycles[cycle_key] = missing_cycles.get(cycle_key, 0) + 1
                continue

            matched_rules = []
            for rule in mass_adjustments:
                if rule["product_master_norm"] not in {"*", product_master_norm}:
                    continue
                if rule["variety_norm"] not in {"*", variety_norm}:
                    continue
                if rule["activity"] not in {"*", activity}:
                    continue
                score = (
                    (1 if rule["product_master_norm"] != "*" else 0)
                    + (1 if rule["variety_norm"] != "*" else 0)
                    + (1 if rule["activity"] != "*" else 0)
                )
                matched_rules.append((score, rule))

            mass_override = {}
            if matched_rules:
                matched_rules.sort(key=lambda item: item[0], reverse=True)
                mass_override = matched_rules[0][1]

            row_override_key = (activity, product_master_norm, variety_norm, block_norm, source_week)
            row_override = row_adjustments.get(row_override_key, {})

            effective_cycle_weeks = (
                row_override.get("cycle_weeks")
                or mass_override.get("cycle_weeks")
                or cycle["cycle_weeks"]
            )
            effective_waste_rate = (
                row_override.get("waste_rate")
                if row_override.get("waste_rate") is not None
                else (
                    mass_override.get("waste_rate")
                    if mass_override.get("waste_rate") is not None
                    else cycle["waste_rate"]
                )
            )
            effective_stems_per_plant = (
                row_override.get("stems_per_plant")
                or mass_override.get("stems_per_plant")
                or cycle["stems_per_plant"]
            )

            source_rows[-1]["cycle_weeks"] = int(effective_cycle_weeks)
            source_rows[-1]["waste_rate"] = float(effective_waste_rate)
            source_rows[-1]["stems_per_plant"] = float(effective_stems_per_plant)
            source_rows[-1]["program_total"] = round(
                plants * effective_stems_per_plant * (1 - effective_waste_rate)
            )

            base_exportable = plants * effective_stems_per_plant * (1 - effective_waste_rate)

            model_projection_by_week = {}
            curve_meta_by_week = {}
            for curve_index, curve_pct in enumerate(cycle["curve"]):
                harvest_week = shift_week(source_week, int(effective_cycle_weeks) + curve_index)
                model_exportable = round(base_exportable * curve_pct)
                model_projection_by_week[harvest_week] = model_projection_by_week.get(
                    harvest_week, 0
                ) + int(model_exportable)
                curve_meta_by_week.setdefault(harvest_week, (curve_index + 1, float(curve_pct)))

            all_harvest_weeks = sorted(set(model_projection_by_week.keys()))
            override_candidates = [
                k
                for k in week_adjustments.keys()
                if k[:5] == (activity, product_master_norm, variety_norm, block_norm, source_week)
            ]
            for k in override_candidates:
                all_harvest_weeks.append(int(k[5]))
            all_harvest_weeks = sorted(set(all_harvest_weeks))

            closure_key = (activity, product_master_norm, variety_norm, block_norm, source_week)
            block_closed = bool(block_closures.get(closure_key, False))
            source_rows[-1]["block_closed"] = block_closed

            for harvest_week in all_harvest_weeks:
                model_exportable = int(model_projection_by_week.get(harvest_week, 0))
                week_override_key = (
                    activity,
                    product_master_norm,
                    variety_norm,
                    block_norm,
                    source_week,
                    harvest_week,
                )
                week_override = week_adjustments.get(week_override_key, {})
                exportable_stems = model_exportable
                weekly_status = "MODELO"
                is_dump = bool(week_override.get("is_dump", False))
                if week_override.get("agronomo_estimate") is not None:
                    exportable_stems = int(week_override["agronomo_estimate"])
                    weekly_status = "AGRONOMO"
                if week_override.get("real_closed") is not None:
                    exportable_stems = int(week_override["real_closed"])
                    weekly_status = "REAL"
                if is_dump:
                    weekly_status = "DUMP"

                curve_week, curve_pct = curve_meta_by_week.get(harvest_week, (0, 0.0))

                projection_rows.append(
                    {
                        "activity": activity,
                        "product_master": product_master,
                        "product_master_norm": product_master_norm,
                        "variety": variety,
                        "block": block,
                        "plants": plants,
                        "source_week": source_week,
                        "source_week_short": format_short_week(source_week),
                        "harvest_week": harvest_week,
                        "harvest_week_short": format_short_week(harvest_week),
                        "cycle_weeks": int(effective_cycle_weeks),
                        "curve_week": curve_week,
                        "curve_pct": curve_pct,
                        "waste_rate": float(effective_waste_rate),
                        "stems_per_plant": float(effective_stems_per_plant),
                        "model_exportable_stems": int(model_exportable),
                        "exportable_stems": int(exportable_stems),
                        "weekly_status": weekly_status,
                        "is_dump": is_dump,
                        "block_closed": block_closed,
                    }
                )

        return {
            "source_rows": source_rows,
            "projection_rows": projection_rows,
            "missing_cycles": [
                {
                    "product_master": product,
                    "variety": variety,
                    "activity": activity,
                    "rows": total_rows,
                }
                for (product, variety, activity), total_rows in sorted(
                    missing_cycles.items(), key=lambda item: item[1], reverse=True
                )
            ],
            "generated_at": pd.Timestamp.now().isoformat(),
        }
    finally:
        db.close()


def get_snapshot(force_refresh: bool = False) -> dict[str, Any]:
    current_version = get_data_version()
    if force_refresh or _WORKER_CACHE["snapshot"] is None or _WORKER_CACHE["version"] != current_version:
        _WORKER_CACHE["snapshot"] = build_projection_snapshot_from_db()
        _WORKER_CACHE["version"] = current_version
    return dict(_WORKER_CACHE["snapshot"])


def available_weeks(snapshot: dict[str, Any]) -> list[int]:
    projection_rows = snapshot.get("projection_rows", [])
    return sorted({int(row["harvest_week"]) for row in projection_rows if row.get("harvest_week")})


def available_products(snapshot: dict[str, Any]) -> list[str]:
    source_rows = snapshot.get("source_rows", [])
    return sorted(
        {
            str(row.get("product_master") or "").strip()
            for row in source_rows
            if row.get("product_master") and row.get("cycle_found")
        }
    )


def resolve_selected_week(raw_week: str, snapshot: dict[str, Any]) -> int:
    requested = parse_week_code(raw_week)
    week_options = available_weeks(snapshot)
    if requested is not None:
        return requested
    if week_options:
        current = current_week_code()
        return min(week_options, key=lambda week_code: abs(week_code - current))
    return current_week_code()


def product_variety_sort_key(product_master: str, variety: str) -> tuple[str, int, str]:
    product = (product_master or "").upper().strip()
    variety_upper = (variety or "").upper().strip()

    if "VERONICA" in product:
        if "SPLASH" in variety_upper:
            if "WHITE" in variety_upper:
                order = 11
            elif "BLUE" in variety_upper:
                order = 12
            elif "PINK" in variety_upper:
                order = 13
            else:
                order = 15
        else:
            if "WHITE" in variety_upper:
                order = 1
            elif "BLUE" in variety_upper:
                order = 2
            elif "PINK" in variety_upper:
                order = 3
            else:
                order = 5
        return (product, order, variety_upper)

    if "HYPERICUM" in product:
        if "ROYAL" in variety_upper:
            order = 1
        elif "TRIUMP" in variety_upper:
            order = 2
        elif "MAJESTIC" in variety_upper or ("PINK" in variety_upper and "MAJESTIC" not in variety_upper):
            order = 3
        elif "CHERRY" in variety_upper:
            order = 4
        elif "SNOW" in variety_upper:
            order = 5
        elif "RENATA" in variety_upper:
            order = 6
        elif "RED" in variety_upper:
            order = 7
        elif "H21-117" in variety_upper:
            order = 8
        elif "H21-68" in variety_upper:
            order = 9
        elif "H21-67" in variety_upper:
            order = 10
        elif re.match(r"^[A-Z0-9]+[-/]", variety_upper) or variety_upper.startswith("H") or variety_upper.startswith("COD"):
            order = 20
        else:
            order = 30
        return (product, order, variety_upper)

    return (product, 100, variety_upper)


def format_variety_display(variety: str) -> str:
    v = (variety or "").strip()
    v_upper = v.upper()
    if v_upper.startswith("SKYLER SPLASH "):
        return "Splash " + v[14:].title()
    elif v_upper == "SKYLER WHITE IMPROVED":
        return "Skyler White Imp."
    elif v_upper.startswith("SKYLER "):
        return "Skyler " + v[7:].title()
    elif v_upper.startswith("MAGICAL "):
        return v[8:].title()
    return v.title() if (v.isupper() and len(v) > 3) else v


def aggregate_for_week(
    snapshot: dict[str, Any],
    selected_week: int,
    horizon_weeks_back: int,
    horizon_weeks_forward: int,
    selected_product_master: str = "",
) -> dict[str, Any]:
    projection_frame = pd.DataFrame(snapshot["projection_rows"])
    source_frame = pd.DataFrame(snapshot["source_rows"])
    selected_product_master = (selected_product_master or "").strip()

    if selected_product_master and not source_frame.empty:
        selected_norm = normalize_text(selected_product_master)
        source_frame = source_frame.loc[source_frame["product_master_norm"] == selected_norm].copy()
        if not projection_frame.empty:
            projection_frame = projection_frame.loc[
                projection_frame["product_master_norm"] == selected_norm
            ].copy()

    week_window = build_week_window_with_lookback(
        selected_week, horizon_weeks_back, horizon_weeks_forward
    )
    week_columns = [
        {"week_code": week_code, "label": format_short_week(week_code)} for week_code in week_window
    ]

    if projection_frame.empty:
        matrix_rows: list[dict[str, Any]] = []
        totals_by_activity = {"SIEMBRA": 0, "PODA": 0}
        totals_by_week = {column["label"]: 0 for column in week_columns}
        totals_by_week_by_activity = {
            "SIEMBRA": {column["label"]: 0 for column in week_columns},
            "PODA": {column["label"]: 0 for column in week_columns},
        }
    else:
        window_projection = projection_frame.loc[
            projection_frame["harvest_week"].isin(week_window)
        ].copy()
        totals_by_activity = {
            activity: int(round(total))
            for activity, total in window_projection.groupby("activity")["exportable_stems"].sum().items()
        }
        totals_by_activity.setdefault("SIEMBRA", 0)
        totals_by_activity.setdefault("PODA", 0)
        totals_by_week = {
            column["label"]: int(
                round(
                    float(
                        window_projection.loc[
                            window_projection["harvest_week"] == column["week_code"],
                            "exportable_stems",
                        ].sum()
                    )
                )
            )
            for column in week_columns
        }
        totals_by_week_by_activity = {
            "SIEMBRA": {
                column["label"]: int(
                    round(
                        float(
                            window_projection.loc[
                                (window_projection["harvest_week"] == column["week_code"])
                                & (window_projection["activity"] == "SIEMBRA"),
                                "exportable_stems",
                            ].sum()
                        )
                    )
                )
                for column in week_columns
            },
            "PODA": {
                column["label"]: int(
                    round(
                        float(
                            window_projection.loc[
                                (window_projection["harvest_week"] == column["week_code"])
                                & (window_projection["activity"] == "PODA"),
                                "exportable_stems",
                            ].sum()
                        )
                    )
                )
                for column in week_columns
            },
        }

        if window_projection.empty:
            matrix_rows = []
        else:
            projection_grouped = (
                window_projection.groupby(
                    [
                        "product_master",
                        "variety",
                        "activity",
                        "source_week",
                        "source_week_short",
                        "block",
                        "cycle_weeks",
                        "stems_per_plant",
                        "waste_rate",
                        "harvest_week",
                    ],
                    dropna=False,
                    as_index=False,
                )
                .agg(
                    {
                        "exportable_stems": "sum",
                        "weekly_status": "first",
                        "is_dump": "first",
                        "model_exportable_stems": "sum",
                    }
                )
                .sort_values(["variety", "activity", "source_week", "block", "harvest_week"])
            )
            source_grouped = (
                source_frame.groupby(
                    [
                        "product_master",
                        "variety",
                        "activity",
                        "source_week",
                        "source_week_short",
                        "block",
                    ],
                    dropna=False,
                    as_index=False,
                )
                .agg(
                    {
                        "plants": "sum",
                        "cycle_weeks": "first",
                        "stems_per_plant": "first",
                        "waste_rate": "first",
                        "ideal_cycle_weeks": "first",
                        "ideal_stems_per_plant": "first",
                        "ideal_waste_rate": "first",
                        "curve_csv": "first",
                        "program_total": "sum",
                        "block_closed": "first",
                    }
                )
                .sort_values(["product_master", "variety", "activity", "source_week", "block"])
            )
            source_keys = {
                (
                    row.product_master,
                    row.variety,
                    row.activity,
                    row.source_week,
                    row.block,
                )
                for row in projection_grouped.itertuples(index=False)
            }
            source_grouped = source_grouped.loc[
                source_grouped.apply(
                    lambda row: (
                        row["product_master"],
                        row["variety"],
                        row["activity"],
                        row["source_week"],
                        row["block"],
                    )
                    in source_keys,
                    axis=1,
                )
            ]

            projection_lookup = {
                (
                    row.product_master,
                    row.variety,
                    row.activity,
                    row.source_week,
                    row.block,
                    row.harvest_week,
                ): int(round(row.exportable_stems))
                for row in projection_grouped.itertuples(index=False)
            }
            projection_status_lookup = {
                (
                    row.product_master,
                    row.variety,
                    row.activity,
                    row.source_week,
                    row.block,
                    row.harvest_week,
                ): str(row.weekly_status or "MODELO")
                for row in projection_grouped.itertuples(index=False)
            }
            projection_dump_lookup = {
                (
                    row.product_master,
                    row.variety,
                    row.activity,
                    row.source_week,
                    row.block,
                    row.harvest_week,
                ): bool(getattr(row, "is_dump", False))
                for row in projection_grouped.itertuples(index=False)
            }
            matrix_rows = []
            for row in source_grouped.itertuples(index=False):
                weekly_projection = {
                    column["label"]: projection_lookup.get(
                        (
                            row.product_master,
                            row.variety,
                            row.activity,
                            row.source_week,
                            row.block,
                            column["week_code"],
                        ),
                        0,
                    )
                    for column in week_columns
                }
                weekly_status = {
                    column["label"]: projection_status_lookup.get(
                        (
                            row.product_master,
                            row.variety,
                            row.activity,
                            row.source_week,
                            row.block,
                            column["week_code"],
                        ),
                        "MODELO",
                    )
                    for column in week_columns
                }
                is_dump_by_week = {
                    column["label"]: projection_dump_lookup.get(
                        (
                            row.product_master,
                            row.variety,
                            row.activity,
                            row.source_week,
                            row.block,
                            column["week_code"],
                        ),
                        False,
                    )
                    for column in week_columns
                }
                week_code_by_label = {
                    column["label"]: column["week_code"] for column in week_columns
                }
                eff_cycle = 0 if pd.isna(row.cycle_weeks) else int(row.cycle_weeks)
                harvest_start_week = shift_week(int(row.source_week), eff_cycle) if eff_cycle else int(row.source_week)
                harvest_start_week_short = format_short_week(harvest_start_week)

                matrix_rows.append(
                    {
                        "variety": row.variety,
                        "variety_display": format_variety_display(row.variety),
                        "product_master": row.product_master,
                        "activity": row.activity,
                        "source_week": row.source_week,
                        "source_week_short": row.source_week_short,
                        "harvest_start_week": harvest_start_week,
                        "harvest_start_week_short": harvest_start_week_short,
                        "block": row.block,
                        "plants": int(row.plants),
                        "cycle_weeks": None if pd.isna(row.cycle_weeks) else int(row.cycle_weeks),
                        "ideal_cycle_weeks": None
                        if pd.isna(row.ideal_cycle_weeks)
                        else int(row.ideal_cycle_weeks),
                        "curve_csv": "" if pd.isna(row.curve_csv) else str(row.curve_csv),
                        "stems_per_plant": 0.0 if pd.isna(row.stems_per_plant) else float(row.stems_per_plant),
                        "ideal_stems_per_plant": 0.0
                        if pd.isna(row.ideal_stems_per_plant)
                        else float(row.ideal_stems_per_plant),
                        "waste_rate": 0.0 if pd.isna(row.waste_rate) else float(row.waste_rate),
                        "ideal_waste_rate": 0.0
                        if pd.isna(row.ideal_waste_rate)
                        else float(row.ideal_waste_rate),
                        "program_total": 0 if pd.isna(row.program_total) else int(round(float(row.program_total))),
                        "weekly_projection": weekly_projection,
                        "weekly_status": weekly_status,
                        "is_dump": is_dump_by_week,
                        "week_code_by_label": week_code_by_label,
                        "window_total": int(sum(weekly_projection.values())),
                        "block_closed": bool(row.block_closed),
                        "real_stems_per_plant": round(
                            (int(sum(weekly_projection.values())) / int(row.plants)) if int(row.plants) > 0 else 0,
                            2,
                        ),
                    }
                )

    grouped_varieties: dict[tuple[str, str], dict[str, Any]] = {}
    for row in matrix_rows:
        variety_bucket = grouped_varieties.setdefault(
            (row["product_master"], row["variety"]),
            {
                "product_master": row["product_master"],
                "variety": row["variety"],
                "variety_display": row["variety_display"],
                "siembras": [],
                "podas": [],
                "total_siembras": 0,
                "total_podas": 0,
            },
        )
        target_list = "siembras" if row["activity"] == "SIEMBRA" else "podas"
        variety_bucket[target_list].append(row)
        if row["activity"] == "SIEMBRA":
            variety_bucket["total_siembras"] += row["window_total"]
        else:
            variety_bucket["total_podas"] += row["window_total"]

    varieties = sorted(
        grouped_varieties.values(),
        key=lambda item: product_variety_sort_key(item.get("product_master", ""), item.get("variety", "")),
    )
    for item in varieties:
        item["total_siembras"] = int(round(item["total_siembras"]))
        item["total_podas"] = int(round(item["total_podas"]))
        item["total_general"] = int(round(item["total_siembras"] + item["total_podas"]))

    return {
        "selected_week": selected_week,
        "selected_week_short": format_short_week(selected_week),
        "selected_product_master": selected_product_master,
        "horizon_weeks_back": horizon_weeks_back,
        "horizon_weeks_forward": horizon_weeks_forward,
        "week_columns": week_columns,
        "varieties": varieties,
        "totals_by_activity": totals_by_activity,
        "totals_by_week": totals_by_week,
        "totals_by_week_by_activity": totals_by_week_by_activity,
        "selected_week_total": int(totals_by_week.get(format_short_week(selected_week), 0)),
        "grand_total": int(round(sum(totals_by_activity.values()))),
    }


def filter_weekly_view_for_real(weekly_view: dict[str, Any]) -> dict[str, Any]:
    week_labels = [column["label"] for column in weekly_view.get("week_columns", [])]
    totals_by_activity = {"SIEMBRA": 0, "PODA": 0}
    totals_by_week_by_activity = {
        "SIEMBRA": {label: 0 for label in week_labels},
        "PODA": {label: 0 for label in week_labels},
    }

    filtered_varieties: list[dict[str, Any]] = []
    for variety in weekly_view.get("varieties", []):
        filtered_siembras = []
        filtered_podas = []

        for row in variety.get("siembras", []):
            statuses = row.get("weekly_status", {}).values()
            if any(status in {"AGRONOMO", "REAL"} for status in statuses):
                filtered_siembras.append(row)
                totals_by_activity["SIEMBRA"] += int(row.get("window_total", 0))
                for label, value in row.get("weekly_projection", {}).items():
                    if label in totals_by_week_by_activity["SIEMBRA"]:
                        totals_by_week_by_activity["SIEMBRA"][label] += int(value)

        for row in variety.get("podas", []):
            statuses = row.get("weekly_status", {}).values()
            if any(status in {"AGRONOMO", "REAL"} for status in statuses):
                filtered_podas.append(row)
                totals_by_activity["PODA"] += int(row.get("window_total", 0))
                for label, value in row.get("weekly_projection", {}).items():
                    if label in totals_by_week_by_activity["PODA"]:
                        totals_by_week_by_activity["PODA"][label] += int(value)

        if not filtered_siembras and not filtered_podas:
            continue

        total_siembras = sum(int(row.get("window_total", 0)) for row in filtered_siembras)
        total_podas = sum(int(row.get("window_total", 0)) for row in filtered_podas)
        filtered_varieties.append(
            {
                "product_master": variety.get("product_master", ""),
                "variety": variety.get("variety", ""),
                "variety_display": variety.get("variety_display", format_variety_display(variety.get("variety", ""))),
                "siembras": filtered_siembras,
                "podas": filtered_podas,
                "total_siembras": total_siembras,
                "total_podas": total_podas,
                "total_general": total_siembras + total_podas,
            }
        )

    totals_by_week = {
        label: int(totals_by_week_by_activity["SIEMBRA"][label] + totals_by_week_by_activity["PODA"][label])
        for label in week_labels
    }
    selected_week_short = weekly_view.get("selected_week_short", "")

    return {
        **weekly_view,
        "varieties": filtered_varieties,
        "totals_by_activity": totals_by_activity,
        "totals_by_week_by_activity": totals_by_week_by_activity,
        "totals_by_week": totals_by_week,
        "selected_week_total": int(totals_by_week.get(selected_week_short, 0)),
        "grand_total": int(totals_by_activity["SIEMBRA"] + totals_by_activity["PODA"]),
    }


# ================= Flask Routes & Middlewares =================

@app.after_request
def add_no_cache_headers(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

@app.route("/")
def index() -> str:
    requested_week = request.args.get("semana", "")
    requested_horizon = request.args.get("horizonte", "")
    requested_horizon_back = request.args.get("horizonte_atras", "")
    requested_horizon_forward = request.args.get("horizonte_adelante", "")
    force_refresh = request.args.get("refresh", "0") == "1"

    snapshot: dict[str, Any] = get_snapshot(force_refresh=force_refresh)
    selected_week = resolve_selected_week(requested_week, snapshot)

    prods = available_products(snapshot)
    if "producto" not in request.args:
        requested_product_master = prods[0] if prods else "HYPERICUM"
    else:
        requested_product_master = request.args.get("producto", "").strip()

    if requested_horizon and not requested_horizon_back and not requested_horizon_forward:
        horizon_weeks_back = max(0, min(20, int(requested_horizon or 4)))
        horizon_weeks_forward = max(1, min(20, int(requested_horizon or 4)))
    else:
        horizon_weeks_back = max(0, min(20, int(requested_horizon_back or 4)))
        horizon_weeks_forward = max(1, min(20, int(requested_horizon_forward or 8)))

    weekly_view = aggregate_for_week(
        snapshot,
        selected_week,
        horizon_weeks_back,
        horizon_weeks_forward,
        selected_product_master=requested_product_master,
    )
    missing_cycles = snapshot.get("missing_cycles", [])

    return render_template(
        "index.html",
        workbook_path="PostgreSQL Database",
        weekly_view=weekly_view,
        page_mode="AGRONOMO",
        page_title="Llenado Agronomo",
        week_options=[
            {"week_code": week_code, "week_short": format_short_week(week_code)}
            for week_code in available_weeks(snapshot)
        ],
        product_options=prods,
        variety_options=[item["variety"] for item in weekly_view["varieties"]],
        missing_cycles=missing_cycles[:20],
        db_status=(True, "Conectado a PostgreSQL"),
    )


@app.route("/reales")
def reales() -> str:
    requested_week = request.args.get("semana", "")
    requested_horizon = request.args.get("horizonte", "")
    requested_horizon_back = request.args.get("horizonte_atras", "")
    requested_horizon_forward = request.args.get("horizonte_adelante", "")
    force_refresh = request.args.get("refresh", "0") == "1"

    snapshot: dict[str, Any] = get_snapshot(force_refresh=force_refresh)
    selected_week = resolve_selected_week(requested_week, snapshot)

    prods = available_products(snapshot)
    if "producto" not in request.args:
        requested_product_master = prods[0] if prods else "HYPERICUM"
    else:
        requested_product_master = request.args.get("producto", "").strip()

    if requested_horizon and not requested_horizon_back and not requested_horizon_forward:
        horizon_weeks_back = max(0, min(20, int(requested_horizon or 4)))
        horizon_weeks_forward = max(1, min(20, int(requested_horizon or 4)))
    else:
        horizon_weeks_back = max(0, min(20, int(requested_horizon_back or 4)))
        horizon_weeks_forward = max(1, min(20, int(requested_horizon_forward or 8)))

    weekly_view = aggregate_for_week(
        snapshot,
        selected_week,
        horizon_weeks_back,
        horizon_weeks_forward,
        selected_product_master=requested_product_master,
    )
    # Load all curve projection rows so user can input real data for any stem/block

    return render_template(
        "index.html",
        workbook_path="PostgreSQL Database",
        weekly_view=weekly_view,
        page_mode="REAL",
        page_title="Datos Reales",
        week_options=[
            {"week_code": week_code, "week_short": format_short_week(week_code)}
            for week_code in available_weeks(snapshot)
        ],
        product_options=prods,
        variety_options=[item["variety"] for item in weekly_view["varieties"]],
        missing_cycles=[],
        db_status=(True, "Conectado a PostgreSQL"),
    )


@app.route("/parametros")
def parametros() -> str:
    db = SessionLocal()
    try:
        snapshot = get_snapshot()
        missing_cycles = snapshot.get("missing_cycles", [])

        cycles = db.query(CycleDefinitionDB).order_by(
            CycleDefinitionDB.product_master, CycleDefinitionDB.variety, CycleDefinitionDB.activity
        ).all()
        cycle_rows = [
            {
                "product_master": c.product_master,
                "variety": c.variety,
                "activity": c.activity,
                "cycle_weeks": c.cycle_weeks,
                "waste_rate_pct": round(c.waste_rate * 100, 2),
                "stems_per_plant": c.stems_per_plant,
                "curve": c.curve or "",
            }
            for c in cycles
        ]
        return render_template(
            "parametros.html",
            workbook_path="PostgreSQL Database",
            cycle_rows=cycle_rows,
            missing_cycles=missing_cycles,
        )
    finally:
        db.close()


@app.route("/tpsr")
def tpsr_view():
    db = SessionLocal()
    try:
        page = request.args.get("page", 1, type=int)
        per_page = request.args.get("per_page", 100, type=int)
        search_query = request.args.get("search", "").strip()
        filter_ac = request.args.get("ac", "").strip()
        filter_week = request.args.get("semana", "").strip()
        filter_pm = request.args.get("pm", "").strip()
        filter_variety = request.args.get("variedad", "").strip()
        filter_block = request.args.get("bloque", "").strip()

        query = db.query(TpsrRecord)

        if filter_ac:
            query = query.filter(TpsrRecord.activity == normalize_text(filter_ac))

        if filter_week:
            parsed_sw = parse_week_code(filter_week)
            if parsed_sw:
                query = query.filter(TpsrRecord.source_week == parsed_sw)
            else:
                query = query.filter(TpsrRecord.source_week_short.like(f"%{filter_week}%"))

        if filter_pm:
            norm_pm = normalize_text(filter_pm)
            query = query.filter(TpsrRecord.product_master_norm == norm_pm)

        if filter_variety:
            norm_var = normalize_text(filter_variety)
            query = query.filter(TpsrRecord.variety_norm.like(f"%{norm_var}%"))

        if filter_block:
            norm_blk = normalize_text(filter_block)
            query = query.filter(TpsrRecord.block_norm == norm_blk)

        if search_query:
            norm_q = normalize_text(search_query)
            query = query.filter(
                (TpsrRecord.variety_norm.like(f"%{norm_q}%")) |
                (TpsrRecord.product_master_norm.like(f"%{norm_q}%")) |
                (TpsrRecord.block_norm.like(f"%{norm_q}%")) |
                (TpsrRecord.activity.like(f"%{norm_q}%"))
            )

        total_records = query.count()
        records = query.order_by(TpsrRecord.source_week.desc(), TpsrRecord.id.desc()).offset((page - 1) * per_page).limit(per_page).all()

        tpsr_rows = [
            {
                "row_index": r.id,  # Database ID
                "activity": r.activity,
                "product_master": r.product_master,
                "variety": r.variety,
                "variety_display": format_variety_display(r.variety),
                "source_week": r.source_week,
                "source_week_short": format_short_week(r.source_week),
                "block": r.block,
                "bed_location": r.bed_location or "",
                "pruning_number": r.pruning_number or "",
                "plants": r.plants,
            }
            for r in records
        ]

        total_pages = (total_records + per_page - 1) // per_page if per_page > 0 else 1

        pm_records = db.query(TpsrRecord.product_master, TpsrRecord.variety).distinct().order_by(TpsrRecord.product_master, TpsrRecord.variety).all()
        pm_varieties_map = {}
        for pm, var in pm_records:
            if not pm or not var:
                continue
            pm_clean = pm.strip()
            var_clean = var.strip()
            if pm_clean not in pm_varieties_map:
                pm_varieties_map[pm_clean] = []
            if var_clean not in pm_varieties_map[pm_clean]:
                pm_varieties_map[pm_clean].append(var_clean)

        available_pms = sorted(list(pm_varieties_map.keys()))
        available_varieties = sorted(list(set(v for vars in pm_varieties_map.values() for v in vars)))
        available_blocks = sorted([b[0] for b in db.query(TpsrRecord.block).distinct().order_by(TpsrRecord.block).all() if b[0]])
        available_weeks = sorted(list(set(format_short_week(w[0]) for w in db.query(TpsrRecord.source_week).distinct().all() if w[0])), reverse=True)

        return render_template(
            "tpsr.html",
            tpsr_rows=tpsr_rows,
            total_records=total_records,
            page=page,
            per_page=per_page,
            total_pages=total_pages,
            search_query=search_query,
            filter_ac=filter_ac,
            filter_week=filter_week,
            filter_pm=filter_pm,
            filter_variety=filter_variety,
            filter_block=filter_block,
            available_pms=available_pms,
            available_varieties=available_varieties,
            available_blocks=available_blocks,
            available_weeks=available_weeks,
            pm_varieties_map=pm_varieties_map,
            workbook_path="PostgreSQL Database",
        )
    finally:
        db.close()


def parse_curve_csv(raw: str | None) -> tuple[float, ...]:
    if not raw:
        return (1.0,)
    try:
        vals = [float(v.strip()) for v in str(raw).split(",") if v.strip()]
        return tuple(vals) if vals else (1.0,)
    except Exception:
        return (1.0,)


def calculate_week_difference(w1: int, w2: int) -> int:
    try:
        w1, w2 = int(w1), int(w2)
        y1, k1 = w1 // 100, w1 % 100
        y2, k2 = w2 // 100, w2 % 100
        return (y2 - y1) * 52 + (k2 - k1)
    except Exception:
        return 0


def find_cycle_definition(
    cycles_map: dict[tuple[str, str, str], dict[str, Any]],
    pm_norm: str,
    var_norm: str,
    ac_norm: str,
) -> dict[str, Any]:
    # 1. Exact match
    c = cycles_map.get((pm_norm, var_norm, ac_norm))
    if c:
        return c
    # 2. Match by variety and activity
    for (p, v, a), cdef in cycles_map.items():
        if v == var_norm and a == ac_norm:
            return cdef
    # 3. Match by variety only
    for (p, v, a), cdef in cycles_map.items():
        if v == var_norm:
            return cdef
    # 4. Match by product_master and activity
    for (p, v, a), cdef in cycles_map.items():
        if p == pm_norm and a == ac_norm:
            return cdef
    # 5. Match by product_master only
    for (p, v, a), cdef in cycles_map.items():
        if p == pm_norm:
            return cdef

    default_stems = 4.0 if "VERONICA" in pm_norm else (9.5 if "GYPSO" in pm_norm else 1.0)
    return {
        "stems_per_plant": default_stems,
        "cycle_weeks": 10,
        "waste_rate": 0.10,
        "curve": (0.10, 0.25, 0.35, 0.20, 0.10),
    }


def get_statistics_data(
    year_filter: str = "",
    pm_filter: str = "",
    variety_filter: str = "",
    activity_filter: str = "",
) -> dict[str, Any]:
    db = SessionLocal()
    try:
        # 1. Base mappings for cascading selectors
        pm_var_rows = db.query(TpsrRecord.product_master, TpsrRecord.variety).distinct().all()
        pm_varieties_map: dict[str, list[str]] = {}
        for pm, var in pm_var_rows:
            if pm and var:
                pm_varieties_map.setdefault(pm, []).append(var)
        for pm in pm_varieties_map:
            pm_varieties_map[pm] = sorted(list(set(pm_varieties_map[pm])))

        sws = db.query(TpsrRecord.source_week).distinct().all()
        extracted_years = sorted(
            list(set(int(str(s[0])[:4]) for s in sws if s[0] and len(str(s[0])) >= 4)),
            reverse=True,
        )
        available_years = [str(y) for y in extracted_years]
        # 2. Cycle Definitions Map & Valid Curve PMs
        cycles_list = db.query(CycleDefinitionDB).all()
        cycles_map: dict[tuple[str, str, str], dict[str, Any]] = {}
        valid_curve_pms: set[str] = set()
        for c in cycles_list:
            key = (c.product_master_norm, c.variety_norm, normalize_text(c.activity))
            c_curve = parse_curve_csv(c.curve)
            cycles_map[key] = {
                "cycle_weeks": c.cycle_weeks,
                "waste_rate": c.waste_rate,
                "stems_per_plant": c.stems_per_plant,
                "curve": c_curve,
            }
            if c.product_master and c_curve and any(x > 0 for x in c_curve):
                valid_curve_pms.add(c.product_master)

        available_pms = sorted([pm for pm in pm_varieties_map.keys() if pm in valid_curve_pms])

        # 3. Query Block Closures
        closures_list = db.query(BlockClosureDB).all()
        closed_blocks_set: set[tuple[str, str, int, str, str]] = set(
            (c.product_master_norm, c.variety_norm, c.source_week, c.activity, c.block_norm)
            for c in closures_list
            if c.is_closed
        )

        # 4. Query TPSR records with filters
        tpsr_query = db.query(TpsrRecord)
        if year_filter:
            try:
                y_int = int(year_filter)
                tpsr_query = tpsr_query.filter(
                    (TpsrRecord.year == y_int)
                    | ((TpsrRecord.source_week >= y_int * 100) & (TpsrRecord.source_week < (y_int + 1) * 100))
                )
            except ValueError:
                pass

        if pm_filter:
            norm_pm = normalize_text(pm_filter)
            tpsr_query = tpsr_query.filter(TpsrRecord.product_master_norm == norm_pm)

        if variety_filter:
            norm_var = normalize_text(variety_filter)
            tpsr_query = tpsr_query.filter(TpsrRecord.variety_norm == norm_var)

        if activity_filter:
            norm_ac = normalize_text(activity_filter)
            tpsr_query = tpsr_query.filter(TpsrRecord.activity == norm_ac)

        tpsr_records = tpsr_query.all()

        # 5. Query Week Adjustments
        week_adj_query = db.query(WeekAdjustmentDB)
        if pm_filter:
            week_adj_query = week_adj_query.filter(WeekAdjustmentDB.product_master_norm == normalize_text(pm_filter))
        if variety_filter:
            week_adj_query = week_adj_query.filter(WeekAdjustmentDB.variety_norm == normalize_text(variety_filter))
        if activity_filter:
            week_adj_query = week_adj_query.filter(WeekAdjustmentDB.activity == normalize_text(activity_filter))

        week_adjustments = week_adj_query.all()
        adj_map_by_block: dict[tuple[str, str, int, str, str], list[WeekAdjustmentDB]] = {}
        for adj in week_adjustments:
            k = (adj.product_master_norm, adj.variety_norm, adj.source_week, adj.activity, adj.block_norm)
            adj_map_by_block.setdefault(k, []).append(adj)

        # 6. Group TPSR records by individual block: (product_master, variety, source_week, activity, block)
        grouped_block_lots: dict[tuple[str, str, int, str, str], dict[str, Any]] = {}
        for r in tpsr_records:
            pm = r.product_master
            var = r.variety
            if "SPLASH" in (r.variety_norm or "") or (pm and pm.upper() == "VERONICA SPRAY"):
                pm = "VERONICA SPRAY"
            if valid_curve_pms and pm not in valid_curve_pms:
                continue

            k = (pm, var, r.source_week, r.activity, r.block)
            if k not in grouped_block_lots:
                grouped_block_lots[k] = {
                    "product_master": pm,
                    "variety": var,
                    "variety_display": format_variety_display(var),
                    "source_week": r.source_week,
                    "source_week_short": format_short_week(r.source_week),
                    "activity": r.activity,
                    "block": r.block,
                    "year": int(str(r.source_week)[:4]) if len(str(r.source_week)) >= 4 else None,
                    "plants": 0,
                }
            grouped_block_lots[k]["plants"] += (r.plants or 0)

        # 7. Process individual block rows and calculate metrics
        flat_block_rows: list[dict[str, Any]] = []

        total_production_sum = 0
        total_real_closed_sum = 0
        total_plants_sum = 0
        total_dumps_sum = 0
        real_stems_per_plant_list = []
        cycle_real_list = []
        waste_list = []
        ideal_stems_per_plant_list = []

        relative_curve_real_sums: dict[int, float] = {i: 0.0 for i in range(12)}
        relative_curve_real_counts: dict[int, int] = {i: 0 for i in range(12)}
        relative_curve_ideal_sums: dict[int, float] = {i: 0.0 for i in range(12)}
        relative_curve_ideal_counts: dict[int, int] = {i: 0 for i in range(12)}

        chronological_real_stems: dict[str, int] = {}
        chronological_agronomo_stems: dict[str, int] = {}

        for k, lot in grouped_block_lots.items():
            pm, var, sw, ac, blk = k
            pm_norm = normalize_text(pm)
            var_norm = normalize_text(var)
            ac_norm = normalize_text(ac)
            blk_norm = normalize_text(blk)
            k_norm = (pm_norm, var_norm, sw, ac_norm, blk_norm)

            c_def = find_cycle_definition(cycles_map, pm_norm, var_norm, ac_norm)
            plants = lot["plants"]
            effective_cycle_weeks = c_def["cycle_weeks"]
            effective_waste_rate = c_def["waste_rate"]
            effective_stems_pp = c_def["stems_per_plant"]

            base_exportable = round(plants * effective_stems_pp * (1 - effective_waste_rate))

            adjs = adj_map_by_block.get(k_norm, [])
            adj_real_map = {a.harvest_week: a.real_closed for a in adjs if a.real_closed is not None}
            adj_agro_map = {a.harvest_week: a.agronomo_estimate for a in adjs if a.agronomo_estimate is not None}

            # Build complete curve weeks
            all_hw_set = set()
            curve_hw_map = {}
            for idx, pct in enumerate(c_def["curve"]):
                hw = shift_week(sw, effective_cycle_weeks + idx)
                all_hw_set.add(hw)
                curve_hw_map[hw] = round(base_exportable * pct)

            for hw in adj_agro_map:
                all_hw_set.add(hw)
            for hw in adj_real_map:
                all_hw_set.add(hw)

            harvest_breakdown = []
            block_total_production = 0
            block_real_stems = 0

            for hw in sorted(all_hw_set):
                hw_short = format_short_week(hw)
                if hw in adj_real_map:
                    val = adj_real_map[hw]
                    src = "REAL"
                    block_real_stems += val
                    chronological_real_stems[hw_short] = chronological_real_stems.get(hw_short, 0) + val
                elif hw in adj_agro_map:
                    val = adj_agro_map[hw]
                    src = "AGRONOMO"
                    chronological_agronomo_stems[hw_short] = chronological_agronomo_stems.get(hw_short, 0) + val
                else:
                    val = curve_hw_map.get(hw, 0)
                    src = "MODELO"

                block_total_production += val
                hw_pp = round(val / plants, 2) if plants > 0 else 0.0
                harvest_breakdown.append({
                    "harvest_week": hw,
                    "harvest_week_short": hw_short,
                    "stems": val,
                    "stems_pp": hw_pp,
                    "source": src,
                })

            for item in harvest_breakdown:
                item["pct_of_total"] = round((item["stems"] / block_total_production * 100.0), 1) if block_total_production > 0 else 0.0

            is_closed = k_norm in closed_blocks_set
            has_real_data = block_real_stems > 0

            block_t_pl = round(block_total_production / plants, 1) if plants > 0 else 0.0
            is_below_target = (block_t_pl < effective_stems_pp)

            # Cycle calculation
            real_cycle_weeks = effective_cycle_weeks
            if adj_real_map:
                min_hw = min(adj_real_map.keys())
                real_cycle_weeks = max(1, calculate_week_difference(sw, min_hw))
                cycle_real_list.append(real_cycle_weeks)

            # Relative curve accumulation
            if block_real_stems > 0 and adj_real_map:
                sorted_hw = sorted(adj_real_map.keys())
                start_hw = sorted_hw[0]
                for hw, stems in adj_real_map.items():
                    offset = max(0, min(11, calculate_week_difference(start_hw, hw)))
                    pct = (stems / block_real_stems) * 100.0
                    relative_curve_real_sums[offset] += pct
                    relative_curve_real_counts[offset] += 1

            ideal_curve = c_def["curve"]
            for idx, pct in enumerate(ideal_curve):
                if idx < 12:
                    relative_curve_ideal_sums[idx] += (pct * 100.0)
                    relative_curve_ideal_counts[idx] += 1

            # Include block in summary
            if plants > 0:
                total_plants_sum += plants
                total_production_sum += block_total_production
                total_real_closed_sum += block_real_stems
                real_stems_per_plant_list.append(block_t_pl)
                ideal_stems_per_plant_list.append(effective_stems_pp)
                waste_list.append(effective_waste_rate * 100.0)

                harvest_weeks_str = ", ".join(item["harvest_week_short"] for item in harvest_breakdown)
                status_label = "CERRADO" if is_closed else ("EN COSECHA" if has_real_data else "PENDIENTE")

                block_item = {
                    "product_master": pm,
                    "variety": var,
                    "variety_display": lot["variety_display"],
                    "block": blk,
                    "source_week": sw,
                    "source_week_short": lot["source_week_short"],
                    "activity": ac,
                    "plants": plants,
                    "exportable_stems": base_exportable,
                    "total_production": block_total_production,
                    "real_stems": block_real_stems,
                    "real_stems_per_plant": block_t_pl,
                    "ideal_stems_per_plant": round(effective_stems_pp, 1),
                    "is_below_target": is_below_target,
                    "real_cycle_weeks": real_cycle_weeks,
                    "ideal_cycle_weeks": effective_cycle_weeks,
                    "waste_pct": round(effective_waste_rate * 100.0, 1),
                    "is_closed": is_closed,
                    "status": status_label,
                    "has_real_data": has_real_data,
                    "harvest_weeks_str": harvest_weeks_str,
                    "harvest_breakdown": harvest_breakdown,
                }
                flat_block_rows.append(block_item)

        flat_block_rows.sort(key=lambda x: (x["source_week"], x["product_master"], x["variety"], x["block"]))

        # 8. Hierarchical Grouping: Producto Maestro -> Semana de Origen -> Bloques
        pm_sw_map: dict[str, dict[tuple[int, str], list[dict[str, Any]]]] = {}
        for b in flat_block_rows:
            pm = b["product_master"]
            sw_k = (b["source_week"], b["activity"])
            pm_sw_map.setdefault(pm, {}).setdefault(sw_k, []).append(b)

        grouped_products: list[dict[str, Any]] = []
        chart3_weeks_map: dict[str, dict[str, Any]] = {}

        for pm in sorted(pm_sw_map.keys()):
            if valid_curve_pms and pm not in valid_curve_pms:
                continue

            sw_dict = pm_sw_map[pm]
            sw_groups = []
            pm_tot_plants = 0
            pm_tot_prod = 0
            pm_tot_real = 0
            pm_closed_cnt = 0
            pm_all_blocks = []

            for (sw, ac) in sorted(sw_dict.keys()):
                blocks = sw_dict[(sw, ac)]
                blocks.sort(key=lambda x: x["block"])
                sw_plants = sum(b["plants"] for b in blocks)
                sw_prod = sum(b["total_production"] for b in blocks)
                sw_real = sum(b["real_stems"] for b in blocks)
                sw_closed = sum(1 for b in blocks if b["is_closed"])
                sw_t_pl = round(sw_prod / sw_plants, 1) if sw_plants > 0 else 0.0
                sw_ideal_t_pl = round(sum(b["ideal_stems_per_plant"] * b["plants"] for b in blocks) / sw_plants, 1) if sw_plants > 0 else 0.0
                sw_cycle = round(sum(b["real_cycle_weeks"] for b in blocks) / len(blocks), 1) if blocks else 0.0
                sw_waste = round(sum(b["waste_pct"] for b in blocks) / len(blocks), 1) if blocks else 10.0
                sw_short = format_short_week(sw)

                # Chart 3 X-axis accumulation (by Source Week)
                chart3_label = f"Sem {sw_short}"
                if chart3_label not in chart3_weeks_map:
                    chart3_weeks_map[chart3_label] = {
                        "sw_short": sw_short,
                        "total_plants": 0,
                        "total_production": 0,
                        "ideal_target": sw_ideal_t_pl,
                        "blocks_count": 0,
                    }
                chart3_weeks_map[chart3_label]["total_plants"] += sw_plants
                chart3_weeks_map[chart3_label]["total_production"] += sw_prod
                chart3_weeks_map[chart3_label]["blocks_count"] += len(blocks)

                sw_groups.append({
                    "source_week": sw,
                    "source_week_short": sw_short,
                    "activity": ac,
                    "total_plants": sw_plants,
                    "total_production": sw_prod,
                    "total_real_stems": sw_real,
                    "avg_real_stems_pp": sw_t_pl,
                    "avg_ideal_stems_pp": sw_ideal_t_pl,
                    "avg_cycle_weeks": sw_cycle,
                    "avg_waste_pct": sw_waste,
                    "blocks_count": len(blocks),
                    "closed_blocks_count": sw_closed,
                    "blocks": blocks,
                })

                pm_tot_plants += sw_plants
                pm_tot_prod += sw_prod
                pm_tot_real += sw_real
                pm_closed_cnt += sw_closed
                pm_all_blocks.extend(blocks)

            pm_t_pl = round(pm_tot_prod / pm_tot_plants, 1) if pm_tot_plants > 0 else 0.0
            pm_ideal_t_pl = round(sum(b["ideal_stems_per_plant"] * b["plants"] for b in pm_all_blocks) / pm_tot_plants, 1) if pm_tot_plants > 0 else 0.0
            pm_cycle = round(sum(b["real_cycle_weeks"] for b in pm_all_blocks) / len(pm_all_blocks), 1) if pm_all_blocks else 0.0
            pm_waste = round(sum(b["waste_pct"] for b in pm_all_blocks) / len(pm_all_blocks), 1) if pm_all_blocks else 10.0

            grouped_products.append({
                "product_master": pm,
                "total_plants": pm_tot_plants,
                "total_production": pm_tot_prod,
                "total_real_stems": pm_tot_real,
                "avg_real_stems_pp": pm_t_pl,
                "avg_ideal_stems_pp": pm_ideal_t_pl,
                "avg_cycle_weeks": pm_cycle,
                "avg_waste_pct": pm_waste,
                "total_blocks": len(pm_all_blocks),
                "closed_blocks_count": pm_closed_cnt,
                "source_week_groups": sw_groups,
            })

        # Overall KPIs
        avg_real_stems_pp = round(total_production_sum / total_plants_sum, 1) if (total_plants_sum > 0 and total_production_sum > 0) else (round(sum(real_stems_per_plant_list) / len(real_stems_per_plant_list), 1) if real_stems_per_plant_list else 0.0)
        avg_ideal_stems_pp = round(sum(ideal_stems_per_plant_list) / len(ideal_stems_per_plant_list), 1) if ideal_stems_per_plant_list else (4.0 if "VERONICA" in pm_filter.upper() else (9.5 if "GYPSO" in pm_filter.upper() else 1.0))
        avg_cycle = round(sum(cycle_real_list) / len(cycle_real_list), 1) if cycle_real_list else 0.0
        avg_waste = round(sum(waste_list) / len(waste_list), 1) if waste_list else 10.0
        total_closed_blocks = sum(1 for r in flat_block_rows if r["is_closed"])

        # Chart 1: Relative Curve
        curve_labels = [f"Sem {i}" if i == 0 else f"Sem +{i}" for i in range(8)]
        chart_real_curve = []
        chart_ideal_curve = []
        for i in range(8):
            r_val = (relative_curve_real_sums[i] / relative_curve_real_counts[i]) if relative_curve_real_counts[i] > 0 else 0.0
            i_val = (relative_curve_ideal_sums[i] / relative_curve_ideal_counts[i]) if relative_curve_ideal_counts[i] > 0 else 0.0
            chart_real_curve.append(round(r_val, 1))
            chart_ideal_curve.append(round(i_val, 1))

        r_sum = sum(chart_real_curve)
        if r_sum > 0:
            chart_real_curve = [round((v / r_sum) * 100.0, 1) for v in chart_real_curve]

        # Chart 2: Chronological Harvest Evolution with Secondary Axis (% Variation / Fulfillment)
        sorted_chrono_weeks = sorted(list(set(list(chronological_real_stems.keys()) + list(chronological_agronomo_stems.keys()))))[-16:]
        chrono_real_data = [chronological_real_stems.get(w, 0) for w in sorted_chrono_weeks]
        chrono_agro_data = [chronological_agronomo_stems.get(w, 0) for w in sorted_chrono_weeks]
        chrono_fulfillment_pct = [
            round((r / a * 100.0), 1) if a > 0 else (100.0 if r > 0 else 0.0)
            for r, a in zip(chrono_real_data, chrono_agro_data)
        ]
        chrono_variation_pct = [
            round(((r - a) / a * 100.0), 1) if a > 0 else 0.0
            for r, a in zip(chrono_real_data, chrono_agro_data)
        ]

        # Chart 2 (Previously 3): Real Stems / Plant by Source Week vs Target Meta Line (Expanded Horizontally)
        chart3_sorted_labels = sorted(list(chart3_weeks_map.keys()))[-20:]
        chart_stems_pp_real = []
        chart_stems_pp_target = []
        for lbl in chart3_sorted_labels:
            info = chart3_weeks_map[lbl]
            pp = round(info["total_production"] / info["total_plants"], 1) if info["total_plants"] > 0 else 0.0
            chart_stems_pp_real.append(pp)
            chart_stems_pp_target.append(avg_ideal_stems_pp)

        # Chart 4: Hybrid Lifecycle Curve (Closed Real Stems + Future Projected Estimates)
        hybrid_weeks = sorted_chrono_weeks
        chart_hybrid_real = [chronological_real_stems.get(w, 0) for w in hybrid_weeks]
        chart_hybrid_projected = [chronological_agronomo_stems.get(w, 0) for w in hybrid_weeks]

        return {
            "kpis": {
                "total_real_stems": total_production_sum,
                "total_plants": total_plants_sum,
                "avg_real_stems_per_plant": avg_real_stems_pp,
                "avg_ideal_stems_per_plant": avg_ideal_stems_pp,
                "avg_cycle_weeks": avg_cycle,
                "avg_waste_pct": avg_waste,
                "total_dumps": total_dumps_sum,
                "total_lots": len(flat_block_rows),
                "total_closed_blocks": total_closed_blocks,
            },
            "table_rows": flat_block_rows,
            "grouped_products": grouped_products,
            "chart_relative": {
                "labels": curve_labels,
                "real_curve": chart_real_curve,
                "ideal_curve": chart_ideal_curve,
            },
            "chart_chrono": {
                "labels": sorted_chrono_weeks,
                "real_stems": chrono_real_data,
                "agronomo_stems": chrono_agro_data,
                "fulfillment_pct": chrono_fulfillment_pct,
                "variation_pct": chrono_variation_pct,
            },
            "chart_stems_pp": {
                "labels": chart3_sorted_labels,
                "real_stems_pp": chart_stems_pp_real,
                "target_stems_pp": chart_stems_pp_target,
            },
            "chart_hybrid": {
                "labels": hybrid_weeks,
                "real_stems": chart_hybrid_real,
                "projected_stems": chart_hybrid_projected,
            },
            "available_years": available_years,
            "available_pms": available_pms,
            "pm_varieties_map": pm_varieties_map,
            "selected_year": year_filter,
            "selected_pm": pm_filter,
            "selected_variety": variety_filter,
            "selected_activity": activity_filter,
        }
    finally:
        db.close()


@app.route("/estadistica")
def estadistica_view() -> str:
    year_filter = request.args.get("year", "").strip()
    pm_filter = request.args.get("pm", "").strip()
    variety_filter = request.args.get("variedad", "").strip()
    activity_filter = request.args.get("ac", "").strip()

    stats_data = get_statistics_data(
        year_filter=year_filter,
        pm_filter=pm_filter,
        variety_filter=variety_filter,
        activity_filter=activity_filter,
    )

    return render_template(
        "estadistica.html",
        stats=stats_data,
        workbook_path="PostgreSQL Database",
    )


@app.route("/api/estadistica")
def estadistica_api():
    year_filter = request.args.get("year", "").strip()
    pm_filter = request.args.get("pm", "").strip()
    variety_filter = request.args.get("variedad", "").strip()
    activity_filter = request.args.get("ac", "").strip()

    stats_data = get_statistics_data(
        year_filter=year_filter,
        pm_filter=pm_filter,
        variety_filter=variety_filter,
        activity_filter=activity_filter,
    )
    return jsonify(stats_data)


@app.post("/api/tpsr-cargar")
def upload_tpsr_file_api():
    if "file" not in request.files:
        return jsonify({"ok": False, "message": "No se seleccionó ningún archivo."}), 400

    uploaded_file = request.files["file"]
    if not uploaded_file.filename or not uploaded_file.filename.lower().endswith(".xlsx"):
        return jsonify({"ok": False, "message": "Por favor sube un archivo Excel (.xlsx) válido."}), 400

    try:
        res = process_tpsr_excel_upload(uploaded_file.stream, commit=True)
        if res.get("ok"):
            invalidate_snapshot_cache()
            return jsonify(res)
        else:
            return jsonify(res), 400
    except Exception as exc:
        return jsonify({"ok": False, "message": f"Error al procesar archivo: {exc}"}), 500


@app.route("/api/proyecciones")
def projections_api():
    requested_week = request.args.get("semana", "")
    requested_horizon = request.args.get("horizonte", "")
    requested_horizon_back = request.args.get("horizonte_atras", "")
    requested_horizon_forward = request.args.get("horizonte_adelante", "")
    requested_product_master = request.args.get("producto", "").strip()
    force_refresh = request.args.get("refresh", "0") == "1"

    snapshot: dict[str, Any] = get_snapshot(force_refresh=force_refresh)
    selected_week = resolve_selected_week(requested_week, snapshot)
    if requested_horizon and not requested_horizon_back and not requested_horizon_forward:
        horizon_weeks_back = max(0, min(20, int(requested_horizon or 4)))
        horizon_weeks_forward = max(1, min(20, int(requested_horizon or 4)))
    else:
        horizon_weeks_back = max(0, min(20, int(requested_horizon_back or 4)))
        horizon_weeks_forward = max(1, min(20, int(requested_horizon_forward or 8)))

    weekly_view = aggregate_for_week(
        snapshot,
        selected_week,
        horizon_weeks_back,
        horizon_weeks_forward,
        selected_product_master=requested_product_master,
    )
    return jsonify(weekly_view)


@app.post("/api/ajustes-fila")
def save_row_adjustment_api():
    payload = request.get_json(silent=True) or {}
    activity = normalize_text(payload.get("activity"))
    product_master = str(payload.get("product_master") or "").strip()
    variety = str(payload.get("variety") or "").strip()
    block = str(payload.get("block") or "").strip()
    source_week = parse_week_code(payload.get("source_week"))

    cycle_weeks = to_int_or_none(payload.get("cycle_weeks"))
    waste_rate_pct = to_float_or_none(payload.get("waste_rate_pct"))
    stems_per_plant = to_float_or_none(payload.get("stems_per_plant"))

    if activity not in {"SIEMBRA", "PODA"}:
        return jsonify({"ok": False, "message": "Actividad invalida."}), 400
    if not product_master or not variety or not block or source_week is None:
        return jsonify({"ok": False, "message": "Faltan llaves para guardar el ajuste."}), 400

    db = SessionLocal()
    try:
        pm_norm = normalize_text(product_master)
        v_norm = normalize_text(variety)
        b_norm = normalize_text(block)

        existing = db.query(RowAdjustmentDB).filter_by(
            activity=activity,
            product_master_norm=pm_norm,
            variety_norm=v_norm,
            block_norm=b_norm,
            source_week=source_week,
        ).first()

        if existing:
            existing.cycle_weeks = cycle_weeks
            existing.waste_rate = waste_rate_pct / 100 if waste_rate_pct is not None else None
            existing.stems_per_plant = stems_per_plant
            existing.updated_at = datetime.utcnow()
        else:
            new_adj = RowAdjustmentDB(
                activity=activity,
                product_master_norm=pm_norm,
                variety_norm=v_norm,
                block=block,
                block_norm=b_norm,
                source_week=source_week,
                cycle_weeks=cycle_weeks,
                waste_rate=waste_rate_pct / 100 if waste_rate_pct is not None else None,
                stems_per_plant=stems_per_plant,
            )
            db.add(new_adj)

        db.commit()
        invalidate_snapshot_cache()
        return jsonify({"ok": True, "message": "Ajuste guardado en PostgreSQL."})
    except SQLAlchemyError as exc:
        db.rollback()
        return jsonify({"ok": False, "message": str(exc)}), 500
    finally:
        db.close()


@app.post("/api/ajustes-semana")
def save_week_adjustment_api():
    payload = request.get_json(silent=True) or {}
    activity = normalize_text(payload.get("activity"))
    product_master = str(payload.get("product_master") or "").strip()
    variety = str(payload.get("variety") or "").strip()
    block = str(payload.get("block") or "").strip()
    source_week = parse_week_code(payload.get("source_week"))
    harvest_week = parse_week_code(payload.get("harvest_week"))
    raw_value = payload.get("value")
    value = to_int_or_none(raw_value)
    mode = normalize_text(payload.get("mode"))
    is_dump = payload.get("is_dump")
    is_reset = bool(payload.get("reset")) or (raw_value in (None, "") and is_dump is None)
    reset_all = bool(payload.get("reset_all"))

    if activity not in {"SIEMBRA", "PODA"}:
        return jsonify({"ok": False, "message": "Actividad invalida."}), 400
    if mode not in {"AGRONOMO", "REAL"}:
        return jsonify({"ok": False, "message": "Modo invalido."}), 400
    if not product_master or not variety or not block or source_week is None or harvest_week is None:
        return jsonify({"ok": False, "message": "Faltan llaves."}), 400

    db = SessionLocal()
    try:
        pm_norm = normalize_text(product_master)
        v_norm = normalize_text(variety)
        b_norm = normalize_text(block)

        existing = db.query(WeekAdjustmentDB).filter_by(
            activity=activity,
            product_master_norm=pm_norm,
            variety_norm=v_norm,
            block_norm=b_norm,
            source_week=source_week,
            harvest_week=harvest_week,
        ).first()

        if existing:
            if reset_all:
                existing.agronomo_estimate = None
                existing.real_closed = None
                existing.is_dump = False
            elif mode == "AGRONOMO":
                existing.agronomo_estimate = value
            else:
                existing.real_closed = value
                if is_dump is not None:
                    existing.is_dump = bool(is_dump)
                elif is_reset:
                    existing.is_dump = False
            
            existing.updated_at = datetime.utcnow()

            if existing.agronomo_estimate is None and existing.real_closed is None and not existing.is_dump:
                db.delete(existing)
        else:
            if not is_reset and not reset_all:
                if value is not None or (is_dump is not None and bool(is_dump)):
                    new_adj = WeekAdjustmentDB(
                        activity=activity,
                        product_master_norm=pm_norm,
                        variety_norm=v_norm,
                        block=block,
                        block_norm=b_norm,
                        source_week=source_week,
                        harvest_week=harvest_week,
                        agronomo_estimate=value if mode == "AGRONOMO" else None,
                        real_closed=value if mode == "REAL" else None,
                        is_dump=bool(is_dump) if is_dump is not None else False,
                    )
                    db.add(new_adj)

        db.commit()
        invalidate_snapshot_cache()

        weekly_status = "MODELO"
        is_dump_val = False
        check = db.query(WeekAdjustmentDB).filter_by(
            activity=activity,
            product_master_norm=pm_norm,
            variety_norm=v_norm,
            block_norm=b_norm,
            source_week=source_week,
            harvest_week=harvest_week,
        ).first()
        if check:
            if check.real_closed is not None:
                weekly_status = "REAL"
            elif check.agronomo_estimate is not None:
                weekly_status = "AGRONOMO"
            if check.is_dump:
                weekly_status = "DUMP"
                is_dump_val = True

        return jsonify({
            "ok": True,
            "message": "Dato semanal guardado en PostgreSQL.",
            "weekly_status": weekly_status,
            "is_dump": is_dump_val,
        })
    except SQLAlchemyError as exc:
        db.rollback()
        return jsonify({"ok": False, "message": str(exc)}), 500
    finally:
        db.close()


@app.post("/api/cierre-bloque")
def save_block_closure_api():
    payload = request.get_json(silent=True) or {}
    activity = normalize_text(payload.get("activity"))
    product_master = str(payload.get("product_master") or "").strip()
    variety = str(payload.get("variety") or "").strip()
    block = str(payload.get("block") or "").strip()
    source_week = parse_week_code(payload.get("source_week"))
    is_closed = bool(payload.get("is_closed"))

    if activity not in {"SIEMBRA", "PODA"}:
        return jsonify({"ok": False, "message": "Actividad invalida."}), 400
    if not product_master or not variety or not block or source_week is None:
        return jsonify({"ok": False, "message": "Faltan llaves."}), 400

    db = SessionLocal()
    try:
        pm_norm = normalize_text(product_master)
        v_norm = normalize_text(variety)
        b_norm = normalize_text(block)

        existing = db.query(BlockClosureDB).filter_by(
            activity=activity,
            product_master_norm=pm_norm,
            variety_norm=v_norm,
            block_norm=b_norm,
            source_week=source_week,
        ).first()

        if existing:
            existing.is_closed = is_closed
            existing.updated_at = datetime.utcnow()
        else:
            new_cls = BlockClosureDB(
                activity=activity,
                product_master_norm=pm_norm,
                variety_norm=v_norm,
                block=block,
                block_norm=b_norm,
                source_week=source_week,
                is_closed=is_closed,
            )
            db.add(new_cls)

        db.commit()
        invalidate_snapshot_cache()
        return jsonify({"ok": True, "message": "Cierre de bloque actualizado en PostgreSQL."})
    except SQLAlchemyError as exc:
        db.rollback()
        return jsonify({"ok": False, "message": str(exc)}), 500
    finally:
        db.close()


@app.post("/api/ciclos-actualizar")
def update_cycle_api():
    payload = request.get_json(silent=True) or {}
    product_master = str(payload.get("product_master") or "").strip()
    variety = str(payload.get("variety") or "").strip()
    activity = normalize_text(payload.get("activity"))
    cycle_weeks = to_int_or_none(payload.get("cycle_weeks"))
    waste_rate_pct = to_float_or_none(payload.get("waste_rate_pct"))
    stems_per_plant = to_float_or_none(payload.get("stems_per_plant"))
    curve = str(payload.get("curve") or "").strip()

    if not product_master or not variety or activity not in {"SIEMBRA", "PODA"}:
        return jsonify({"ok": False, "message": "Llaves invalidas."}), 400

    db = SessionLocal()
    try:
        pm_norm = normalize_text(product_master)
        v_norm = normalize_text(variety)

        cycle = db.query(CycleDefinitionDB).filter_by(
            product_master_norm=pm_norm,
            variety_norm=v_norm,
            activity=activity,
        ).first()

        if not cycle:
            cycle = CycleDefinitionDB(
                product_master=product_master,
                product_master_norm=pm_norm,
                variety=variety,
                variety_norm=v_norm,
                activity=activity,
            )
            db.add(cycle)

        if cycle_weeks is not None: cycle.cycle_weeks = cycle_weeks
        if waste_rate_pct is not None: cycle.waste_rate = waste_rate_pct / 100
        if stems_per_plant is not None: cycle.stems_per_plant = stems_per_plant
        if curve: cycle.curve = curve
        cycle.updated_at = datetime.utcnow()

        db.commit()
        invalidate_snapshot_cache()
        return jsonify({"ok": True, "message": "Parámetro actualizado en PostgreSQL."})
    except SQLAlchemyError as exc:
        db.rollback()
        return jsonify({"ok": False, "message": str(exc)}), 500
    finally:
        db.close()


@app.post("/api/ajustes-masivos")
def save_mass_adjustment_api():
    payload = request.get_json(silent=True) or {}
    product_master = str(payload.get("product_master") or "").strip()
    variety = str(payload.get("variety") or "").strip()
    activity = normalize_optional_text(payload.get("activity"), fallback="*")

    cycle_weeks = to_int_or_none(payload.get("cycle_weeks"))
    waste_rate_pct = to_float_or_none(payload.get("waste_rate_pct"))
    stems_per_plant = to_float_or_none(payload.get("stems_per_plant"))

    if not product_master:
        return jsonify({"ok": False, "message": "Producto maestro obligatorio."}), 400

    db = SessionLocal()
    try:
        pm_norm = normalize_optional_text(product_master)
        v_norm = normalize_optional_text(variety)

        existing = db.query(MassAdjustmentDB).filter_by(
            product_master_norm=pm_norm,
            variety_norm=v_norm,
            activity=activity,
        ).first()

        if existing:
            existing.cycle_weeks = cycle_weeks
            existing.waste_rate = waste_rate_pct / 100 if waste_rate_pct is not None else None
            existing.stems_per_plant = stems_per_plant
            existing.updated_at = datetime.utcnow()
        else:
            new_adj = MassAdjustmentDB(
                product_master_norm=pm_norm,
                variety_norm=v_norm,
                activity=activity,
                cycle_weeks=cycle_weeks,
                waste_rate=waste_rate_pct / 100 if waste_rate_pct is not None else None,
                stems_per_plant=stems_per_plant,
            )
            db.add(new_adj)

        db.commit()
        invalidate_snapshot_cache()
        return jsonify({"ok": True, "message": "Ajuste masivo guardado en PostgreSQL."})
    except SQLAlchemyError as exc:
        db.rollback()
        return jsonify({"ok": False, "message": str(exc)}), 500
    finally:
        db.close()


@app.post("/api/tpsr-actualizar")
def update_tpsr_api():
    payload = request.get_json(silent=True) or {}
    rec_id = to_int_or_none(payload.get("row_index") or payload.get("id"))
    if rec_id is None or rec_id < 1:
        return jsonify({"ok": False, "message": "ID de registro inválido."}), 400

    db = SessionLocal()
    try:
        rec = db.query(TpsrRecord).filter_by(id=rec_id).first()
        if not rec:
            return jsonify({"ok": False, "message": "Registro no encontrado en PostgreSQL."}), 404

        if "activity" in payload and payload["activity"]:
            act = normalize_text(payload["activity"])
            if act in {"SIEMBRA", "PODA"}:
                rec.activity = act

        if "source_week" in payload and payload["source_week"]:
            sw = parse_week_code(payload["source_week"])
            if sw is not None:
                rec.source_week = sw
                rec.source_week_short = format_short_week(sw)

        if "product_master" in payload and payload["product_master"]:
            pm = str(payload["product_master"]).strip()
            if pm:
                rec.product_master = pm
                rec.product_master_norm = normalize_text(pm)

        if "variety" in payload and payload["variety"]:
            var = str(payload["variety"]).strip()
            if var:
                rec.variety = var
                rec.variety_norm = normalize_text(var)

        if "block" in payload and payload["block"]:
            blk = str(payload["block"]).strip()
            if blk:
                rec.block = blk
                rec.block_norm = normalize_text(blk)

        if "plants" in payload and payload["plants"] is not None:
            pl = to_int_or_none(payload["plants"])
            if pl is not None and pl >= 0:
                rec.plants = pl

        if "bed_location" in payload:
            rec.bed_location = str(payload.get("bed_location") or "").strip()

        if "pruning_number" in payload:
            rec.pruning_number = str(payload.get("pruning_number") or "").strip()

        rec.updated_at = datetime.utcnow()
        db.commit()
        invalidate_snapshot_cache()
        return jsonify({"ok": True, "message": "Registro TPSR actualizado correctamente en PostgreSQL."})
    except SQLAlchemyError as exc:
        db.rollback()
        return jsonify({"ok": False, "message": str(exc)}), 500
    finally:
        db.close()


@app.post("/api/tpsr-crear")
def create_tpsr_api():
    payload = request.get_json(silent=True) or {}
    act = normalize_text(payload.get("activity"))
    sw = parse_week_code(payload.get("source_week"))
    pm = str(payload.get("product_master") or "").strip()
    var = str(payload.get("variety") or "").strip()
    blk = str(payload.get("block") or "").strip()
    plants = to_int_or_none(payload.get("plants"))

    if act not in {"SIEMBRA", "PODA"}:
        return jsonify({"ok": False, "message": "Actividad debe ser SIEMBRA o PODA."}), 400
    if sw is None:
        return jsonify({"ok": False, "message": "Semana origen inválida (ej. 2610 o 202610)."}), 400
    if not pm or not var or not blk:
        return jsonify({"ok": False, "message": "Producto maestro, variedad y bloque son obligatorios."}), 400
    if plants is None or plants < 0:
        return jsonify({"ok": False, "message": "Número de plantas inválido."}), 400

    db = SessionLocal()
    try:
        new_rec = TpsrRecord(
            source_week=sw,
            source_week_short=format_short_week(sw),
            activity=act,
            product_master=pm,
            product_master_norm=normalize_text(pm),
            variety=var,
            variety_norm=normalize_text(var),
            block=blk,
            block_norm=normalize_text(blk),
            plants=plants,
            bed_location=str(payload.get("bed_location") or "").strip(),
            pruning_number=str(payload.get("pruning_number") or "").strip(),
        )
        db.add(new_rec)
        db.commit()
        invalidate_snapshot_cache()
        return jsonify({"ok": True, "message": "Nuevo registro TPSR creado con éxito en PostgreSQL.", "id": new_rec.id})
    except SQLAlchemyError as exc:
        db.rollback()
        return jsonify({"ok": False, "message": str(exc)}), 500
    finally:
        db.close()


@app.post("/api/tpsr-eliminar")
def delete_tpsr_api():
    payload = request.get_json(silent=True) or {}
    rec_id = to_int_or_none(payload.get("row_index") or payload.get("id"))

    if rec_id is None or rec_id < 1:
        return jsonify({"ok": False, "message": "ID de registro inválido."}), 400

    db = SessionLocal()
    try:
        rec = db.query(TpsrRecord).filter_by(id=rec_id).first()
        if not rec:
            return jsonify({"ok": False, "message": "Registro no encontrado en PostgreSQL."}), 404

        db.delete(rec)
        db.commit()
        invalidate_snapshot_cache()
        return jsonify({"ok": True, "message": "Registro eliminado correctamente de PostgreSQL."})
    except SQLAlchemyError as exc:
        db.rollback()
        return jsonify({"ok": False, "message": str(exc)}), 500
    finally:
        db.close()


if __name__ == "__main__":
    debug_mode = os.environ.get("FLASK_DEBUG", "0").lower() in ("1", "true")
    app.run(host="0.0.0.0", port=5000, debug=debug_mode)
