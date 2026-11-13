
from __future__ import annotations

import os
import re
import time
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List
from collections import defaultdict

import csv
import gzip
import io
import pandas as pd
from flask import (
    Flask,
    Response,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from sqlalchemy.exc import SQLAlchemyError

from db import (
    AppMetaDB,
    BlockClosureDB,
    CycleDefinitionDB,
    EstimateAuditDB,
    MassAdjustmentDB,
    RowAdjustmentDB,
    SessionLocal,
    TpsrRecord,
    UserDB,
    WeekAdjustmentDB,
    bump_data_version,
    format_short_week,
    get_data_version,
    normalize_text,
    parse_week_code,
    process_tpsr_excel_upload,
    seed_data_from_excel_if_empty,
    verify_user,
    create_user,
    update_user,
    delete_user,
    list_users,
    get_user_by_id,
    parse_allowed_views,
)

BASE_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = BASE_DIR / "siembras_podas" / "tpsr.xlsx"

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "proyecciones_secret_session_key_2026_prod")
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=7)


def merge_bed_locations(bed_locs: list[str]) -> str:
    if not bed_locs:
        return ""
    clean_locs = [b.strip() for b in bed_locs if b and str(b).strip() and str(b).strip() not in ("None", "nan", "-")]
    if not clean_locs:
        return ""
    if len(clean_locs) == 1:
        return clean_locs[0]

    nums = []
    has_unparsed = False
    for loc in clean_locs:
        parts = re.findall(r"\d+", loc)
        if parts:
            nums.extend(int(p) for p in parts)
        else:
            has_unparsed = True

    if nums and not has_unparsed:
        min_n = min(nums)
        max_n = max(nums)
        if min_n == max_n:
            return f"{min_n}"
        return f"{min_n} a {max_n}"

    return ", ".join(sorted(set(clean_locs)))


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
_WORKER_CACHE: dict[str, Any] = {"version": -1, "timestamp": 0.0, "snapshot": None}
_STATS_CACHE: dict[tuple, Any] = {}


def invalidate_snapshot_cache() -> None:
    bump_data_version()
    _WORKER_CACHE["snapshots"] = {}
    _WORKER_CACHE["snapshot"] = None
    _WORKER_CACHE["version"] = -1
    _WORKER_CACHE["timestamp"] = 0.0
    _STATS_CACHE.clear()


def build_projection_snapshot_from_db(year: int | None = 2026) -> dict[str, Any]:
    db = SessionLocal()
    try:
        tpsr_query = db.query(TpsrRecord)
        if year:
            tpsr_query = tpsr_query.filter(
                (TpsrRecord.year == year) |
                ((TpsrRecord.source_week >= year * 100) & (TpsrRecord.source_week < (year + 1) * 100))
            )
        tpsr_recs = tpsr_query.all()
        cycle_defs = db.query(CycleDefinitionDB).all()

        row_adjs_query = db.query(RowAdjustmentDB)
        if year:
            row_adjs_query = row_adjs_query.filter(
                (RowAdjustmentDB.source_week >= year * 100) & (RowAdjustmentDB.source_week < (year + 1) * 100)
            )
        row_adjs = row_adjs_query.all()

        mass_adjs = db.query(MassAdjustmentDB).all()

        week_adjs_query = db.query(WeekAdjustmentDB)
        if year:
            week_adjs_query = week_adjs_query.filter(
                ((WeekAdjustmentDB.source_week >= year * 100) & (WeekAdjustmentDB.source_week < (year + 1) * 100)) |
                ((WeekAdjustmentDB.harvest_week >= year * 100) & (WeekAdjustmentDB.harvest_week < (year + 1) * 100))
            )
        week_adjs = week_adjs_query.all()

        block_closes_query = db.query(BlockClosureDB)
        if year:
            block_closes_query = block_closes_query.filter(
                (BlockClosureDB.source_week >= year * 100) & (BlockClosureDB.source_week < (year + 1) * 100)
            )
        block_closes = block_closes_query.all()

        cycle_map = {}
        cycle_map_by_var_act = {}
        cycle_map_by_clean_v = {}
        for c in cycle_defs:
            curve_values = tuple(
                float(val)
                for val in (c.curve or "1.0").split(",")
                if val.strip() and float_or_zero(val) > 0
            ) or (1.0,)
            c_pm = "VERONICA" if "VERONICA" in (c.product_master or "").upper() else c.product_master
            c_pm_norm = "VERONICA" if "VERONICA" in c.product_master_norm else c.product_master_norm
            c_entry = {
                "product_master": c_pm,
                "variety": c.variety,
                "activity": c.activity,
                "cycle_weeks": c.cycle_weeks,
                "waste_rate": c.waste_rate,
                "stems_per_plant": c.stems_per_plant,
                "curve": curve_values,
            }
            cycle_map[(c_pm_norm, c.variety_norm, c.activity)] = c_entry
            cycle_map[(c.product_master_norm, c.variety_norm, c.activity)] = c_entry
            cycle_map_by_var_act[(c.variety_norm, c.activity)] = c_entry
            clean_v = c.variety_norm.rstrip("S")
            if len(clean_v) > 3:
                cycle_map_by_clean_v[(clean_v, c.activity)] = c_entry

        row_adjustments = {
            (
                r.activity,
                "VERONICA" if "VERONICA" in r.product_master_norm else r.product_master_norm,
                r.variety_norm,
                r.block_norm,
                r.source_week,
            ): {
                "cycle_weeks": r.cycle_weeks,
                "waste_rate": r.waste_rate,
                "stems_per_plant": r.stems_per_plant,
            }
            for r in row_adjs
        }

        mass_adjustments = [
            {
                "product_master_norm": "VERONICA" if "VERONICA" in m.product_master_norm else m.product_master_norm,
                "variety_norm": m.variety_norm,
                "activity": m.activity,
                "cycle_weeks": m.cycle_weeks,
                "waste_rate": m.waste_rate,
                "stems_per_plant": m.stems_per_plant,
            }
            for m in mass_adjs
        ]
        mass_override_cache = {}

        week_adjustments_by_block = defaultdict(dict)
        for w in week_adjs:
            pm_norm = "VERONICA" if "VERONICA" in w.product_master_norm else w.product_master_norm
            b_key = (w.activity, pm_norm, w.variety_norm, w.block_norm, w.source_week)
            week_adjustments_by_block[b_key][w.harvest_week] = {
                "agronomo_estimate": w.agronomo_estimate,
                "real_closed": w.real_closed,
                "dump_stems": w.dump_stems or 0,
                "is_dump": bool(w.is_dump or (w.dump_stems and w.dump_stems > 0)),
            }

        block_closures = {
            (
                b.activity,
                "VERONICA" if "VERONICA" in b.product_master_norm else b.product_master_norm,
                b.variety_norm,
                b.block_norm,
                b.source_week,
            ): b.is_closed
            for b in block_closes
        }

        # Consolidar registros TPSR a nivel de bloque único (activity, product_master, variety, block, source_week)
        consolidated_dict: dict[tuple[str, str, str, str, int], dict[str, Any]] = {}
        for rec in tpsr_recs:
            activity = normalize_text(rec.activity)
            if activity not in {"SIEMBRA", "PODA"}:
                activity = "SIEMBRA"

            source_week = parse_week_code(rec.source_week) or parse_week_code(rec.source_week_short)
            if not source_week:
                try:
                    sw_int = int(rec.source_week)
                    if sw_int < 100:
                        source_week = 202600 + sw_int
                    elif sw_int < 10000:
                        source_week = 200000 + sw_int if sw_int >= 2400 else 202600 + (sw_int % 100)
                    elif 200001 <= sw_int <= 200053 or 202001 <= sw_int <= 202053:
                        source_week = 202600 + (sw_int % 100)
                    else:
                        source_week = sw_int
                except (ValueError, TypeError):
                    source_week = current_week_code()

            plants = int(rec.plants or 0)
            product_master = str(rec.product_master or "").strip() or "GYPSOPHILA"
            pm_upper = product_master.upper()
            if "GYP" in pm_upper:
                product_master = "GYPSOPHILA"
            elif "VERON" in pm_upper:
                product_master = "VERONICA"
            elif "HYPER" in pm_upper:
                product_master = "HYPERICUM"

            product_master_norm = normalize_text(product_master)
            variety = str(rec.variety or "").strip()
            variety_norm = normalize_text(rec.variety_norm or variety)
            block = str(rec.block or "").strip()
            block_norm = normalize_text(rec.block_norm or block)

            if "VERONICA" in product_master_norm or "SPLASH" in variety_norm:
                product_master = "VERONICA"
                product_master_norm = "VERONICA"

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
                    "bed_locations": [rec.bed_location] if rec.bed_location else [],
                }
            else:
                consolidated_dict[key]["plants"] += plants
                if rec.bed_location and rec.bed_location not in consolidated_dict[key]["bed_locations"]:
                    consolidated_dict[key]["bed_locations"].append(rec.bed_location)

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
            bed_location_str = merge_bed_locations(item.get("bed_locations", []))

            cycle_key = (product_master_norm, variety_norm, activity)
            cycle = cycle_map.get(cycle_key)
            if cycle is None:
                cycle = cycle_map_by_var_act.get((variety_norm, activity))
            if cycle is None:
                clean_v = variety_norm.rstrip("S")
                if len(clean_v) > 3:
                    cycle = cycle_map_by_clean_v.get((clean_v, activity))

            # If no cycle definition exists for this variety in Parámetros, DO NOT project or display it!
            # It will safely remain in TPSR and will only enter projection once its curve is created.
            if cycle is None:
                missing_cycles[(product_master, variety, activity)] = (
                    missing_cycles.get((product_master, variety, activity), 0) + 1
                )
                continue

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
                    "bed_location": bed_location_str,
                    "pruning_number": "",
                    "cycle_found": True,
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

            mass_key = (product_master_norm, variety_norm, activity)
            if mass_key in mass_override_cache:
                mass_override = mass_override_cache[mass_key]
            else:
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
                mass_override_cache[mass_key] = mass_override

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

            closure_key = (activity, product_master_norm, variety_norm, block_norm, source_week)
            block_closed = bool(block_closures.get(closure_key, False))
            source_rows[-1]["block_closed"] = block_closed

            block_week_adjs = week_adjustments_by_block.get(closure_key, {})
            all_harvest_weeks = sorted(set(model_projection_by_week.keys()) | set(block_week_adjs.keys()))

            for harvest_week in all_harvest_weeks:
                model_exportable = int(model_projection_by_week.get(harvest_week, 0))
                week_override = block_week_adjs.get(harvest_week, {})
                exportable_stems = model_exportable
                weekly_status = "MODELO"
                is_dump = bool(week_override.get("is_dump", False))
                if week_override.get("agronomo_estimate") is not None:
                    exportable_stems = int(week_override["agronomo_estimate"])
                    if exportable_stems > 0:
                        weekly_status = "AGRONOMO"
                if week_override.get("real_closed") is not None:
                    exportable_stems = int(week_override["real_closed"])
                    if exportable_stems > 0:
                        weekly_status = "REAL"
                elif block_closed:
                    exportable_stems = 0
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
                        "bed_location": bed_location_str,
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
                        "dump_stems": int(week_override.get("dump_stems", 0) or 0),
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


def get_snapshot(force_refresh: bool = False, year: int = 2026) -> dict[str, Any]:
    current_version = get_data_version()
    now = time.time()
    # Cache is strictly version-controlled with PostgreSQL data_version.
    # Fallback TTL is 600s (10 min) to avoid wasteful recalculations every 30s.
    cache_expired = (now - _WORKER_CACHE.get("timestamp", 0.0)) > 600.0
    if force_refresh or cache_expired or _WORKER_CACHE.get("version") != current_version:
        _WORKER_CACHE["snapshots"] = {}
        _WORKER_CACHE["snapshot"] = None
        _WORKER_CACHE["version"] = current_version
        _WORKER_CACHE["timestamp"] = now

    if "snapshots" not in _WORKER_CACHE:
        _WORKER_CACHE["snapshots"] = {}

    if year not in _WORKER_CACHE["snapshots"]:
        _WORKER_CACHE["snapshots"][year] = build_projection_snapshot_from_db(year=year)
    _WORKER_CACHE["snapshot"] = _WORKER_CACHE["snapshots"][year]
    return _WORKER_CACHE["snapshots"][year]


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
        if "WHITE" in variety_upper:
            order = 2 if "SPLASH" in variety_upper else 1
        elif "BLUE" in variety_upper:
            order = 4 if "SPLASH" in variety_upper else 3
        elif "PINK" in variety_upper:
            order = 6 if "SPLASH" in variety_upper else 5
        elif "DUO" in variety_upper or "CRYSTAL" in variety_upper:
            order = 7
        else:
            order = 8
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
    selected_product_master = (selected_product_master or "").strip()
    selected_norm = normalize_text(selected_product_master) if selected_product_master else None

    week_window = build_week_window_with_lookback(
        selected_week, horizon_weeks_back, horizon_weeks_forward
    )
    week_window_set = set(week_window)
    week_columns = [
        {"week_code": week_code, "label": format_short_week(week_code)} for week_code in week_window
    ]
    week_label_map = {column["week_code"]: column["label"] for column in week_columns}

    totals_by_activity = {"SIEMBRA": 0, "PODA": 0}
    totals_by_week = {column["label"]: 0 for column in week_columns}
    totals_by_week_by_activity = {
        "SIEMBRA": {column["label"]: 0 for column in week_columns},
        "PODA": {column["label"]: 0 for column in week_columns},
    }

    lifetime_metrics: dict[tuple[str, str, str, int, str], dict[str, int]] = {}
    projection_lookup: dict[tuple[str, str, str, int, str, int], int] = {}
    projection_status_lookup: dict[tuple[str, str, str, int, str, int], str] = {}
    projection_dump_lookup: dict[tuple[str, str, str, int, str, int], bool] = {}
    projection_dump_stems_lookup: dict[tuple[str, str, str, int, str, int], int] = {}
    source_keys: set[tuple[str, str, str, int, str]] = set()

    for row in snapshot.get("projection_rows", []):
        if selected_norm and row.get("product_master_norm") != selected_norm:
            continue

        pm = str(row.get("product_master") or "")
        var = str(row.get("variety") or "")
        act = str(row.get("activity") or "")
        sw = int(row.get("source_week") or 0)
        blk = str(row.get("block") or "")
        hw = int(row.get("harvest_week") or 0)
        exp = int(row.get("exportable_stems") or 0)
        dumps = int(row.get("dump_stems") or 0)

        # Lifetime accumulation
        b_key = (pm, var, act, sw, blk)
        m = lifetime_metrics.get(b_key)
        if m is None:
            lifetime_metrics[b_key] = {
                "total_production": exp + dumps,
                "total_exportable": exp,
                "total_dumps": dumps,
            }
        else:
            m["total_production"] += exp + dumps
            m["total_exportable"] += exp
            m["total_dumps"] += dumps

        # Window projection
        if hw in week_window_set:
            source_keys.add(b_key)
            lbl = week_label_map[hw]
            totals_by_week[lbl] += exp
            if act in totals_by_activity:
                totals_by_activity[act] += exp
                totals_by_week_by_activity[act][lbl] += exp

            p_key = (pm, var, act, sw, blk, hw)
            projection_lookup[p_key] = exp
            projection_status_lookup[p_key] = str(row.get("weekly_status") or "MODELO")
            projection_dump_lookup[p_key] = bool(row.get("is_dump", False))
            projection_dump_stems_lookup[p_key] = dumps

    min_window_week = min(week_window) if week_window else selected_week
    recent_source_weeks = set(build_week_window_with_lookback(selected_week, 40, 26))

    # Filter and sort source rows
    candidates = []
    for row in snapshot.get("source_rows", []):
        if selected_norm and row.get("product_master_norm") != selected_norm:
            continue
        try:
            sw_val = int(row.get("source_week") or 0)
        except (ValueError, TypeError):
            continue
        pm = str(row.get("product_master") or "")
        var = str(row.get("variety") or "")
        act = str(row.get("activity") or "")
        blk = str(row.get("block") or "")
        b_key = (pm, var, act, sw_val, blk)

        is_closed = bool(row.get("block_closed", False))
        eff_cycle = int(row.get("cycle_weeks") or 0)
        h_start = shift_week(sw_val, eff_cycle) if eff_cycle else sw_val

        # Include if it has harvests in the current window OR is an active/growing batch whose harvest is current/upcoming
        if b_key in source_keys or (not is_closed and sw_val in recent_source_weeks and h_start >= min_window_week):
            candidates.append(row)

    candidates.sort(key=lambda r: (r.get("product_master") or "", r.get("variety") or "", r.get("activity") or "", r.get("source_week") or 0, r.get("block") or ""))

    matrix_rows: list[dict[str, Any]] = []
    for row in candidates:
        pm = str(row.get("product_master") or "")
        var = str(row.get("variety") or "")
        act = str(row.get("activity") or "")
        sw = int(row.get("source_week") or 0)
        blk = str(row.get("block") or "")

        weekly_projection = {}
        weekly_status = {}
        is_dump_by_week = {}
        dump_stems_by_week = {}
        week_code_by_label = {}
        window_total_prod = 0

        for col in week_columns:
            lbl = col["label"]
            hw_code = col["week_code"]
            week_code_by_label[lbl] = hw_code
            p_key = (pm, var, act, sw, blk, hw_code)

            exp_val = projection_lookup.get(p_key, 0)
            st_val = projection_status_lookup.get(p_key, "MODELO")
            dump_flag = projection_dump_lookup.get(p_key, False)
            dump_st_val = projection_dump_stems_lookup.get(p_key, 0)

            weekly_projection[lbl] = exp_val
            weekly_status[lbl] = st_val
            is_dump_by_week[lbl] = dump_flag
            dump_stems_by_week[lbl] = dump_st_val
            window_total_prod += exp_val + dump_st_val

        eff_cycle = int(row.get("cycle_weeks") or 0)
        harvest_start_week = shift_week(sw, eff_cycle) if eff_cycle else sw
        harvest_start_week_short = format_short_week(harvest_start_week)

        b_key = (pm, var, act, sw, blk)
        b_lifetime = lifetime_metrics.get(b_key, {"total_production": 0, "total_exportable": 0, "total_dumps": 0})
        plants_raw = int(row.get("plants") or 0)
        plants_count = plants_raw if plants_raw > 0 else 1
        lifetime_total_prod = b_lifetime["total_production"]
        lifetime_real_stems = round(lifetime_total_prod / plants_count, 1) if plants_raw > 0 else 0.0

        is_closed = bool(row.get("block_closed", False))

        matrix_rows.append(
            {
                "variety": var,
                "variety_display": format_variety_display(var),
                "product_master": pm,
                "activity": act,
                "bed_location": str(row.get("bed_location") or "").strip(),
                "source_week": sw,
                "source_week_short": row.get("source_week_short") or format_short_week(sw),
                "harvest_start_week": harvest_start_week,
                "harvest_start_week_short": harvest_start_week_short,
                "block": blk,
                "plants": plants_raw,
                "cycle_weeks": row.get("cycle_weeks"),
                "ideal_cycle_weeks": row.get("ideal_cycle_weeks"),
                "curve_csv": str(row.get("curve_csv") or ""),
                "stems_per_plant": float(row.get("stems_per_plant") or 0.0),
                "ideal_stems_per_plant": float(row.get("ideal_stems_per_plant") or 0.0),
                "waste_rate": float(row.get("waste_rate") or 0.0),
                "ideal_waste_rate": float(row.get("ideal_waste_rate") or 0.0),
                "program_total": int(round(float(row.get("program_total") or 0))),
                "weekly_projection": weekly_projection,
                "weekly_status": weekly_status,
                "is_dump": is_dump_by_week,
                "dump_stems": dump_stems_by_week,
                "week_code_by_label": week_code_by_label,
                "window_total": int(window_total_prod),
                "total_production": lifetime_total_prod,
                "block_closed": is_closed,
                "real_stems_per_plant": lifetime_real_stems,
                "lifetime_total_production": lifetime_total_prod,
                "lifetime_real_stems_per_plant": lifetime_real_stems,
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
        "matrix_rows": matrix_rows,
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


ALL_VIEWS_ORDER = ["agro", "reales", "estadistica", "tpsr", "parametros"]

PAGE_TO_VIEW = {
    "/": "agro",
    "/reales": "reales",
    "/estadistica": "estadistica",
    "/tpsr": "tpsr",
    "/parametros": "parametros",
}

VIEW_TO_URL = {
    "agro": "/",
    "reales": "/reales",
    "estadistica": "/estadistica",
    "tpsr": "/tpsr",
    "parametros": "/parametros",
}


@app.context_processor
def inject_auth_context():
    role = session.get("role", "visita")
    views = session.get("allowed_views") or []
    if role == "admin":
        views = ALL_VIEWS_ORDER
    return {
        "current_user": session.get("username"),
        "user_role": role,
        "user_name": session.get("full_name"),
        "user_views": views,
    }


@app.before_request
def require_login():
    # Allow public endpoints and static files
    if request.path.startswith("/static") or request.path in {"/login", "/logout"}:
        return None

    # Check if user is logged in
    if not session.get("username"):
        if request.path.startswith("/api/"):
            return jsonify({"ok": False, "message": "Sesión requerida."}), 401
        target_next = request.full_path if request.method == "GET" and request.full_path != "/?" else "/"
        return redirect(url_for("login_view", next=target_next))

    user_role = session.get("role", "visita")
    user_views = session.get("allowed_views") or []
    if user_role == "admin":
        user_views = ALL_VIEWS_ORDER

    # Admin only routes for User Management
    if request.path.startswith("/usuarios") or request.path.startswith("/api/usuarios"):
        if user_role != "admin":
            if request.path.startswith("/api/"):
                return jsonify({"ok": False, "message": "Acceso restringido: requiere rol de Administrador."}), 403
            for v in ALL_VIEWS_ORDER:
                if v in user_views:
                    return redirect(VIEW_TO_URL[v])
            return redirect(url_for("logout_view"))
        return None

    # Page view restrictions
    if request.path in PAGE_TO_VIEW:
        required_view = PAGE_TO_VIEW[request.path]
        if required_view not in user_views and user_role != "admin":
            for v in ALL_VIEWS_ORDER:
                if v in user_views:
                    return redirect(VIEW_TO_URL[v])
            flash("No tienes vistas asignadas en este momento. Contacta al administrador.", "warning")
            return redirect(url_for("logout_view"))

    # API view restrictions
    if request.path.startswith("/api/"):
        # TPSR APIs
        if request.path.startswith("/api/tpsr"):
            if "tpsr" not in user_views and user_role != "admin":
                return jsonify({"ok": False, "message": "Acceso denegado a módulo TPSR."}), 403
        # Parametros / Ciclos APIs
        elif request.path.startswith("/api/ciclos") or request.path.startswith("/api/ajustes-masivos"):
            if "parametros" not in user_views and user_role != "admin":
                return jsonify({"ok": False, "message": "Acceso denegado a módulo de Configuración de Ciclos."}), 403
        # Estadistica APIs
        elif request.path.startswith("/api/estadistica"):
            if "estadistica" not in user_views and user_role != "admin":
                return jsonify({"ok": False, "message": "Acceso denegado a módulo Estadístico."}), 403
        # Reales specific APIs
        elif request.path in {"/api/bloques-cerrados", "/api/bloques-cerrados/batch"}:
            if "reales" not in user_views and user_role != "admin":
                return jsonify({"ok": False, "message": "Acceso denegado a módulo de Datos Reales."}), 403
        # Agro specific APIs
        elif request.path == "/api/reset-week":
            if "agro" not in user_views and user_role != "admin":
                return jsonify({"ok": False, "message": "Acceso denegado a Llenado Agrónomo."}), 403
        # Shared operational APIs (proyecciones, ajustes-fila, ajustes-semanales, semanas, exportar matriz)
        elif request.path in {"/api/proyecciones", "/api/ajustes-fila", "/api/ajustes-semanales", "/api/semanas", "/api/matriz/exportar"}:
            if ("agro" not in user_views and "reales" not in user_views) and user_role != "admin":
                return jsonify({"ok": False, "message": "Acceso denegado a matriz y exportación."}), 403

    return None


@app.after_request
def compress_response(response: Response) -> Response:
    try:
        # Never attempt compression on direct_passthrough, streamed, or non-2xx responses
        if getattr(response, "direct_passthrough", False) or response.is_streamed or response.status_code >= 300:
            return response

        accept_encoding = request.headers.get("Accept-Encoding", "")
        if "gzip" not in accept_encoding.lower():
            return response

        if "Content-Encoding" in response.headers:
            return response

        mimetype = response.mimetype or ""
        if not (mimetype.startswith("text/") or mimetype in ("application/json", "application/javascript")):
            return response

        data = response.get_data()
        if not data or len(data) <= 1024:
            return response

        compressed = gzip.compress(data, compresslevel=6)
        response.set_data(compressed)
        response.headers["Content-Encoding"] = "gzip"
        response.headers["Content-Length"] = str(len(compressed))
        response.headers["Vary"] = "Accept-Encoding"
        return response
    except Exception:
        # Gracefully return uncompressed response on any issue
        return response


@app.route("/login", methods=["GET", "POST"])
def login_view():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        user_info = verify_user(username, password)
        if user_info:
            session["user_id"] = user_info["id"]
            session["username"] = user_info["username"]
            session["role"] = user_info["role"]
            session["full_name"] = user_info["full_name"]
            session["allowed_views"] = user_info.get("allowed_views", [])
            session.permanent = True

            next_url = request.args.get("next")
            views = user_info.get("allowed_views", [])
            if user_info["role"] == "admin":
                if next_url and next_url.startswith("/") and not next_url.startswith("//"):
                    return redirect(next_url)
                return redirect(url_for("index"))

            # Non-admin: check if next_url is allowed
            if next_url and next_url in PAGE_TO_VIEW and PAGE_TO_VIEW[next_url] in views:
                return redirect(next_url)

            # Redirect to their first allowed view
            for v in ALL_VIEWS_ORDER:
                if v in views:
                    return redirect(VIEW_TO_URL[v])

            flash("El usuario no tiene vistas asignadas. Comuníquese con el administrador.", "warning")
            return redirect(url_for("login_view"))
        else:
            flash("Usuario o contraseña incorrectos. Verifique sus credenciales.", "danger")

    return render_template("login.html")


@app.route("/logout")
def logout_view():
    session.clear()
    return redirect(url_for("login_view"))


# ---------------- USER MANAGEMENT ROUTES (ADMIN ONLY) ---------------- #

@app.route("/usuarios")
def usuarios_view():
    if session.get("role") != "admin":
        flash("Acceso restringido: requiere rol de Administrador.", "danger")
        return redirect(url_for("index"))
    users_list = list_users()
    return render_template("usuarios.html", users=users_list, active_view="usuarios")


@app.post("/api/usuarios/crear")
def crear_usuario_api():
    if session.get("role") != "admin":
        return jsonify({"ok": False, "message": "Acceso restringido: solo administrador."}), 403
    payload = request.get_json(silent=True) or {}
    username = payload.get("username", "")
    password = payload.get("password", "")
    full_name = payload.get("full_name", "")
    role = payload.get("role", "visita")
    allowed_views = payload.get("allowed_views", [])
    ok, msg, new_user = create_user(username, password, full_name, role, allowed_views)
    if not ok:
        return jsonify({"ok": False, "message": msg}), 400
    return jsonify({"ok": True, "message": msg, "user": new_user})


@app.post("/api/usuarios/actualizar/<int:user_id>")
def actualizar_usuario_api(user_id: int):
    if session.get("role") != "admin":
        return jsonify({"ok": False, "message": "Acceso restringido: solo administrador."}), 403
    payload = request.get_json(silent=True) or {}
    username = payload.get("username")
    password = payload.get("password")
    full_name = payload.get("full_name")
    role = payload.get("role")
    allowed_views = payload.get("allowed_views")
    ok, msg = update_user(user_id, username, password, full_name, role, allowed_views)
    if not ok:
        return jsonify({"ok": False, "message": msg}), 400

    # If updating currently logged in user session
    if session.get("user_id") == user_id:
        if full_name:
            session["full_name"] = full_name
        if role:
            session["role"] = role
        if allowed_views is not None:
            session["allowed_views"] = parse_allowed_views(allowed_views, role or session.get("role"))

    return jsonify({"ok": True, "message": msg})


@app.post("/api/usuarios/eliminar/<int:user_id>")
def eliminar_usuario_api(user_id: int):
    if session.get("role") != "admin":
        return jsonify({"ok": False, "message": "Acceso restringido: solo administrador."}), 403
    ok, msg = delete_user(user_id, current_user_id=session.get("user_id"))
    if not ok:
        return jsonify({"ok": False, "message": msg}), 400
    return jsonify({"ok": True, "message": msg})


@app.route("/")
def index() -> str:
    requested_year_str = request.args.get("ano", request.args.get("year", "2026")).strip()
    try:
        requested_year = int(requested_year_str) if requested_year_str else 2026
    except ValueError:
        requested_year = 2026

    requested_week = request.args.get("semana", "")
    requested_horizon = request.args.get("horizonte", "")
    requested_horizon_back = request.args.get("horizonte_atras", "")
    requested_horizon_forward = request.args.get("horizonte_adelante", "")
    force_refresh = (request.args.get("refresh", "0") == "1") or (request.args.get("sync_db", "0") == "1")
    if force_refresh:
        invalidate_snapshot_cache()

    snapshot: dict[str, Any] = get_snapshot(force_refresh=force_refresh, year=requested_year)
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
        active_view="agro",
        selected_year=str(requested_year),
        year_options=["2026", "2025", "2024"],
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
    requested_year_str = request.args.get("ano", request.args.get("year", "2026")).strip()
    try:
        requested_year = int(requested_year_str) if requested_year_str else 2026
    except ValueError:
        requested_year = 2026

    requested_week = request.args.get("semana", "")
    requested_horizon = request.args.get("horizonte", "")
    requested_horizon_back = request.args.get("horizonte_atras", "")
    requested_horizon_forward = request.args.get("horizonte_adelante", "")
    force_refresh = (request.args.get("refresh", "0") == "1") or (request.args.get("sync_db", "0") == "1")
    if force_refresh:
        invalidate_snapshot_cache()

    snapshot: dict[str, Any] = get_snapshot(force_refresh=force_refresh, year=requested_year)
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
        active_view="reales",
        selected_year=str(requested_year),
        year_options=["2026", "2025", "2024"],
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
            active_view="parametros",
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
            active_view="tpsr",
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


def format_week_range_display(weeks_list: list[str]) -> str:
    """Formats a list of week numbers into a compact human-readable range, e.g., 'S15-S21 (7s)' or 'S15'."""
    if not weeks_list:
        return "-"
    cleaned = [w.strip() for w in weeks_list if w and w.strip()]
    if not cleaned:
        return "-"
    try:
        sorted_weeks = sorted(list(set(cleaned)), key=lambda x: int(re.sub(r"[^\d]", "", x) or 0))
    except Exception:
        sorted_weeks = sorted(list(set(cleaned)))
    if not sorted_weeks:
        return "-"
    if len(sorted_weeks) == 1:
        w0 = sorted_weeks[0]
        return f"S{w0}" if not str(w0).startswith("S") else str(w0)
    w_first = sorted_weeks[0]
    w_last = sorted_weeks[-1]
    count = len(sorted_weeks)
    prefix_first = f"S{w_first}" if not str(w_first).startswith("S") else str(w_first)
    prefix_last = f"S{w_last}" if not str(w_last).startswith("S") else str(w_last)
    return f"{prefix_first}-{prefix_last} ({count}s)"


def consolidate_production_cluster(block: str, prod_num: int, cluster_rows: list[dict[str, Any]]) -> dict[str, Any]:
    first = cluster_rows[0]
    activity = first["activity"]
    pm = first["product_master"]
    
    distinct_vars = sorted(list(set(r.get("variety_display") or r["variety"] for r in cluster_rows if r.get("variety"))))
    var_disp = " / ".join(distinct_vars) if distinct_vars else (first.get("variety_display") or first["variety"])
    raw_vars = sorted(list(set(r["variety"] for r in cluster_rows if r.get("variety"))))
    var = ", ".join(raw_vars)

    source_weeks = sorted(list(set(r["source_week"] for r in cluster_rows)))
    source_weeks_short = [format_short_week(w) for w in source_weeks]
    source_weeks_str = ", ".join(source_weeks_short)
    source_weeks_display = format_week_range_display(source_weeks_short)
    source_week_max = max(source_weeks)
    source_week_max_short = format_short_week(source_week_max)

    tot_plants = sum(r.get("plants", 0) for r in cluster_rows)
    tot_exportable = sum(r.get("exportable_stems", r.get("real_stems", 0)) for r in cluster_rows)
    tot_prod = sum(r.get("total_production", 0) for r in cluster_rows)
    tot_real = sum(r.get("real_stems", 0) for r in cluster_rows)
    tot_dumps = sum(r.get("dump_stems", 0) for r in cluster_rows)

    real_stems_pp = round(tot_prod / tot_plants, 2) if tot_plants > 0 else 0.0
    weighted_ideal_sum = sum(r["ideal_stems_per_plant"] * r["plants"] for r in cluster_rows)
    ideal_stems_pp = round(weighted_ideal_sum / tot_plants, 2) if tot_plants > 0 else first["ideal_stems_per_plant"]
    is_below_target = (real_stems_pp < ideal_stems_pp)

    # Per-variety sub-breakdown within this production wave
    variety_breakdown: dict[str, dict[str, Any]] = {}
    variety_hw_maps: dict[str, dict[int, dict[str, Any]]] = {}
    for r in cluster_rows:
        v_name = r.get("variety_display") or r["variety"]
        if v_name not in variety_breakdown:
            variety_breakdown[v_name] = {
                "variety": v_name,
                "plants": 0,
                "production": 0,
                "dump": 0,
                "weighted_ideal_sum": 0.0,
                "ideal_stems_pp": r.get("ideal_stems_per_plant", 0),
            }
            variety_hw_maps[v_name] = {}
        variety_breakdown[v_name]["plants"] += r["plants"]
        variety_breakdown[v_name]["production"] += r["total_production"]
        variety_breakdown[v_name]["dump"] += r.get("dump_stems", 0)
        variety_breakdown[v_name]["weighted_ideal_sum"] += (r.get("ideal_stems_per_plant", 0) * r["plants"])

        for hb in r.get("harvest_breakdown", []):
            if hb.get("stems", 0) <= 0:
                continue
            hw = hb["harvest_week"]
            if hw not in variety_hw_maps[v_name]:
                variety_hw_maps[v_name][hw] = {
                    "harvest_week": hw,
                    "harvest_week_short": hb["harvest_week_short"],
                    "stems": 0,
                    "exportable_stems": 0,
                    "dump_stems": 0,
                    "source": "MODELO",
                }
            v_hw = variety_hw_maps[v_name][hw]
            v_hw["stems"] += hb["stems"]
            v_hw["exportable_stems"] += hb.get("exportable_stems", 0)
            v_hw["dump_stems"] += hb.get("dump_stems", 0)
            if hb.get("source") == "REAL":
                v_hw["source"] = "REAL"
            elif hb.get("source") == "AGRONOMO" and v_hw["source"] != "REAL":
                v_hw["source"] = "AGRONOMO"

    for v_name, v_info in variety_breakdown.items():
        v_plants = v_info["plants"]
        v_prod = v_info["production"]
        v_info["real_stems_pp"] = round(v_prod / v_plants, 2) if v_plants > 0 else 0.0
        if v_plants > 0 and v_info.get("weighted_ideal_sum"):
            v_info["ideal_stems_pp"] = round(v_info["weighted_ideal_sum"] / v_plants, 2)
        v_hw_list = sorted([item for item in variety_hw_maps[v_name].values() if item["stems"] > 0], key=lambda x: x["harvest_week"])
        for item in v_hw_list:
            item["stems_pp"] = round(item["stems"] / v_plants, 2) if v_plants > 0 else 0.0
            item["pct_of_total"] = round((item["stems"] / v_prod * 100.0), 1) if v_prod > 0 else 0.0
            item["activity"] = activity
            item["variety"] = v_name
        v_info["harvest_breakdown"] = v_hw_list
        v_info["harvest_weeks_str"] = ", ".join(i["harvest_week_short"] for i in v_hw_list)
        v_info["harvest_weeks_display"] = format_week_range_display([i["harvest_week_short"] for i in v_hw_list])

    # Consolidated harvest breakdown by harvest_week
    hw_map: dict[int, dict[str, Any]] = {}
    for r in cluster_rows:
        for hb in r.get("harvest_breakdown", []):
            if hb.get("stems", 0) <= 0:
                continue
            hw = hb["harvest_week"]
            if hw not in hw_map:
                hw_map[hw] = {
                    "harvest_week": hw,
                    "harvest_week_short": hb["harvest_week_short"],
                    "stems": 0,
                    "exportable_stems": 0,
                    "dump_stems": 0,
                    "source": "MODELO",
                }
            hw_map[hw]["stems"] += hb["stems"]
            hw_map[hw]["exportable_stems"] += hb.get("exportable_stems", 0)
            hw_map[hw]["dump_stems"] += hb.get("dump_stems", 0)
            if hb.get("source") == "REAL":
                hw_map[hw]["source"] = "REAL"
            elif hb.get("source") == "AGRONOMO" and hw_map[hw]["source"] != "REAL":
                hw_map[hw]["source"] = "AGRONOMO"

    sorted_hw_list = sorted([item for item in hw_map.values() if item["stems"] > 0], key=lambda x: x["harvest_week"])
    for item in sorted_hw_list:
        item["stems_pp"] = round(item["stems"] / tot_plants, 2) if tot_plants > 0 else 0.0
        item["pct_of_total"] = round((item["stems"] / tot_prod * 100.0), 1) if tot_prod > 0 else 0.0

    harvest_weeks_str = ", ".join(i["harvest_week_short"] for i in sorted_hw_list)
    harvest_weeks_display = format_week_range_display([i["harvest_week_short"] for i in sorted_hw_list])

    is_closed = any(r["is_closed"] for r in cluster_rows)
    has_real = any(r["has_real_data"] for r in cluster_rows)
    status_label = "CERRADO" if is_closed else ("EN COSECHA" if has_real else "PENDIENTE")

    act_label = "Siembra" if activity == "SIEMBRA" else "Poda"
    prod_label = f"Producción {prod_num} ({act_label})"

    first_hw = sorted_hw_list[0]["harvest_week"] if sorted_hw_list else None
    real_cycle_weeks = calculate_week_difference(source_weeks[0], first_hw) if first_hw else first["real_cycle_weeks"]
    real_cycle_weeks = max(1, real_cycle_weeks)

    uid = f"prod-{block}-{prod_num}-{source_week_max}"

    return {
        "uid": uid,
        "production_num": prod_num,
        "production_label": prod_label,
        "activity": activity,
        "product_master": pm,
        "variety": var,
        "variety_display": var_disp,
        "variety_breakdown": list(variety_breakdown.values()),
        "block": block,
        "source_weeks": source_weeks,
        "source_weeks_short": source_weeks_short,
        "source_weeks_str": source_weeks_str,
        "source_weeks_display": source_weeks_display,
        "source_week_max": source_week_max,
        "source_week_max_short": source_week_max_short,
        "is_staggered": len(source_weeks) > 1,
        "staggered_count": len(source_weeks),
        "plants": tot_plants,
        "exportable_stems": tot_exportable,
        "total_production": tot_prod,
        "real_stems": tot_real,
        "dump_stems": tot_dumps,
        "real_stems_per_plant": real_stems_pp,
        "ideal_stems_per_plant": ideal_stems_pp,
        "is_below_target": is_below_target,
        "real_cycle_weeks": real_cycle_weeks,
        "ideal_cycle_weeks": first.get("ideal_cycle_weeks", first.get("real_cycle_weeks", 0)),
        "waste_pct": first.get("waste_pct", 0.0),
        "is_closed": is_closed,
        "status": status_label,
        "has_real_data": has_real,
        "harvest_weeks_str": harvest_weeks_str,
        "harvest_weeks_display": harvest_weeks_display,
        "harvest_breakdown": sorted_hw_list,
    }


def build_consolidated_block_productions(flat_block_rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    by_block: dict[str, list[dict[str, Any]]] = {}
    for r in flat_block_rows:
        blk = str(r["block"]).strip()
        by_block.setdefault(blk, []).append(r)

    def block_sort_key(b_str: str):
        nums = re.findall(r"\d+", b_str)
        return (int(nums[0]) if nums else 999999, b_str)

    sorted_blocks = sorted(by_block.keys(), key=block_sort_key)
    consolidated_blocks = []
    flat_productions = []

    for blk in sorted_blocks:
        rows = by_block[blk]
        rows.sort(key=lambda x: (x["source_week"], x["activity"], x["product_master"]))

        productions: list[dict[str, Any]] = []
        current_cluster: list[dict[str, Any]] = []

        for r in rows:
            if not current_cluster:
                current_cluster.append(r)
            else:
                prev = current_cluster[-1]
                is_same_act = (r["activity"] == prev["activity"])
                is_same_crop = (r["product_master"] == prev["product_master"])
                # Staggered planting window for beds of the same crop and block is up to 4 weeks
                window_weeks = 4
                is_close_week = (calculate_week_difference(prev["source_week"], r["source_week"]) <= window_weeks)

                if is_same_act and is_same_crop and is_close_week:
                    current_cluster.append(r)
                else:
                    prod = consolidate_production_cluster(blk, len(productions) + 1, current_cluster)
                    productions.append(prod)
                    flat_productions.append(prod)
                    current_cluster = [r]

        if current_cluster:
            prod = consolidate_production_cluster(blk, len(productions) + 1, current_cluster)
            productions.append(prod)
            flat_productions.append(prod)

        # Build chronological stages (Etapas de cultivo) for this physical block
        stages: list[dict[str, Any]] = []
        for p in productions:
            pm_curr = p.get("product_master", "OTRO")
            if not stages or stages[-1]["product_master"] != pm_curr:
                stg = {
                    "stage_num": len(stages) + 1,
                    "product_master": pm_curr,
                    "variety_display": p.get("variety_display", ""),
                    "productions": [p],
                    "total_plants": p.get("plants", 0),
                    "total_production": p.get("total_production", 0),
                    "is_closed": p.get("is_closed", False),
                    "has_real_data": p.get("has_real_data", False),
                    "source_weeks_display": p.get("source_weeks_display", ""),
                }
                stages.append(stg)
            else:
                stg = stages[-1]
                stg["productions"].append(p)
                stg["total_plants"] += p.get("plants", 0)
                stg["total_production"] += p.get("total_production", 0)
                if p.get("is_closed"):
                    stg["is_closed"] = True
                if p.get("has_real_data"):
                    stg["has_real_data"] = True

        for stg in stages:
            stg_pl = stg["total_plants"]
            stg_prod = stg["total_production"]
            stg["real_stems_per_plant"] = round(stg_prod / stg_pl, 2) if stg_pl > 0 else 0.0
            stg["status"] = "CERRADO" if stg["is_closed"] else ("EN COSECHA" if stg["has_real_data"] else "PENDIENTE")
            d_vars = []
            for p in stg["productions"]:
                if p.get("variety_display") and p["variety_display"] not in d_vars:
                    d_vars.append(p["variety_display"])
            stg["variety_display"] = " / ".join(d_vars)

        # Consolidate block-level summary metrics across its productions
        total_prod_sum = sum(p["total_production"] for p in productions)
        total_dump_sum = sum(p.get("dump_stems", 0) for p in productions)
        
        # Physical plant capacity across coexisting varieties in the block
        var_max_plants: dict[str, int] = {}
        for p in productions:
            if p.get("variety_breakdown"):
                for vb in p["variety_breakdown"]:
                    v_key = vb["variety"]
                    var_max_plants[v_key] = max(var_max_plants.get(v_key, 0), vb.get("plants", 0))
            else:
                v_key = p.get("variety_display") or p.get("variety") or "default"
                var_max_plants[v_key] = max(var_max_plants.get(v_key, 0), p.get("plants", 0))

        block_capacity = sum(var_max_plants.values())
        max_plants = max(block_capacity, max((p["plants"] for p in productions), default=0))
        
        tot_plants_all_cycles = sum(p["plants"] for p in productions)
        if tot_plants_all_cycles > 0:
            weighted_real_sum = sum(p["real_stems_per_plant"] * p["plants"] for p in productions)
            real_stems_pp = round(weighted_real_sum / tot_plants_all_cycles, 2)
        else:
            real_stems_pp = round(total_prod_sum / max_plants, 2) if max_plants > 0 else 0.0
        
        weighted_ideal_sum = sum(p["ideal_stems_per_plant"] * p["plants"] for p in productions)
        weighted_plants_sum = sum(p["plants"] for p in productions)
        ideal_stems_pp = round(weighted_ideal_sum / weighted_plants_sum, 2) if weighted_plants_sum > 0 else 0.0
        
        is_below = (real_stems_pp < ideal_stems_pp) if (real_stems_pp > 0 and ideal_stems_pp > 0) else False
        
        pms = sorted(list(set(p["product_master"] for p in productions if p.get("product_master"))))
        var_disps = sorted(list(set(p["variety_display"] for p in productions if p.get("variety_display"))))
        vars_raw = sorted(list(set(p["variety"] for p in productions if p.get("variety"))))
        acts = sorted(list(set(p["activity"] for p in productions if p.get("activity"))))
        
        if "SIEMBRA" in acts and "PODA" in acts:
            acts_str = "Siembra + Poda"
        elif "SIEMBRA" in acts:
            acts_str = "Siembra"
        elif "PODA" in acts:
            acts_str = "Poda"
        else:
            acts_str = ", ".join(acts) if acts else "-"
            
        sw_list = []
        sw_display_list = []
        for p in productions:
            act_prefix = "S" if p["activity"] == "SIEMBRA" else "P"
            sw_list.append(f"{act_prefix}:{p['source_weeks_str']}")
            p_disp = p.get("source_weeks_display") or format_week_range_display(p.get("source_weeks_short", []))
            clean_disp = p_disp.lstrip("S") if p_disp.startswith("S") and not p_disp.startswith("Sem") else p_disp
            sw_display_list.append(f"{act_prefix}:{clean_disp}")
        source_weeks_summary = " | ".join(sw_list)
        source_weeks_display = " | ".join(sw_display_list)
        
        avg_cycle = round(sum(p["real_cycle_weeks"] for p in productions) / len(productions), 1) if productions else 0.0
        avg_waste = round(sum(p["waste_pct"] for p in productions) / len(productions), 1) if productions else 0.0
        
        all_hw = sorted(list(set(hw.strip() for p in productions for hw in p.get("harvest_weeks_str", "").split(",") if hw.strip())))
        harvest_weeks_summary = ", ".join(all_hw)
        harvest_weeks_display = format_week_range_display(all_hw)
        
        has_any_closed = any(p["is_closed"] for p in productions) if productions else False
        is_all_closed = all(p["is_closed"] for p in productions) if productions else False
        has_any_real = any(p.get("has_real_data") for p in productions) if productions else False
        overall_status = "CERRADO" if (is_all_closed or has_any_closed) else ("EN COSECHA" if has_any_real else "PENDIENTE")
        
        block_uid = f"blk-cons-{blk}"

        consolidated_blocks.append({
            "block": blk,
            "uid": block_uid,
            "product_master": ", ".join(pms),
            "variety": ", ".join(vars_raw),
            "variety_display": " / ".join(var_disps),
            "varieties_list": var_disps,
            "activities": acts,
            "activities_str": acts_str,
            "source_weeks_str": source_weeks_summary,
            "source_weeks_display": source_weeks_display,
            "source_week_max_short": productions[-1]["source_week_max_short"] if productions else "",
            "total_plants": max_plants,
            "total_production": total_prod_sum,
            "total_dump": total_dump_sum,
            "real_stems_per_plant": real_stems_pp,
            "ideal_stems_per_plant": ideal_stems_pp,
            "is_below_target": is_below,
            "avg_cycle_weeks": avg_cycle,
            "avg_waste_pct": avg_waste,
            "harvest_weeks_str": harvest_weeks_summary,
            "harvest_weeks_display": harvest_weeks_display,
            "is_closed": has_any_closed,
            "is_all_closed": is_all_closed,
            "has_real_data": has_any_real,
            "status": overall_status,
            "productions": productions,
            "total_productions": len(productions),
            "stages": stages,
            "total_stages": len(stages),
        })

    return consolidated_blocks, flat_productions


def get_statistics_data(
    year_filter: str | None = None,
    pm_filter: str | None = None,
    variety_filter: str = "",
    activity_filter: str = "",
    block_filter: str = "",
    weeks_filter: list[str] | str | None = None,
    week_from: str | None = None,
    week_to: str | None = None,
) -> dict[str, Any]:
    current_version = get_data_version()
    if _WORKER_CACHE.get("version") != current_version:
        _STATS_CACHE.clear()

    var_key = tuple(sorted(str(v) for v in (variety_filter if isinstance(variety_filter, list) else [variety_filter])))
    wk_key = tuple(sorted(str(w) for w in (weeks_filter if isinstance(weeks_filter, list) else [weeks_filter])))
    cache_key = (
        current_version,
        str(year_filter or ""),
        str(pm_filter or ""),
        var_key,
        str(activity_filter or ""),
        str(block_filter or ""),
        wk_key,
        str(week_from or ""),
        str(week_to or ""),
    )
    if cache_key in _STATS_CACHE:
        return _STATS_CACHE[cache_key]

    db = SessionLocal()
    try:
        # 1. Base mappings for cascading selectors
        pm_var_blk_rows = db.query(TpsrRecord.product_master, TpsrRecord.variety, TpsrRecord.block).distinct().all()
        pm_varieties_map: dict[str, list[str]] = {}
        pm_blocks_map: dict[str, list[str]] = {}
        for pm, var, blk in pm_var_blk_rows:
            if pm:
                pm_clean = "VERONICA" if "VERONICA" in pm.upper() else pm
                if var:
                    pm_varieties_map.setdefault(pm_clean, []).append(var)
                if blk:
                    pm_blocks_map.setdefault(pm_clean, []).append(blk.strip())
        for pm in pm_varieties_map:
            pm_varieties_map[pm] = sorted(list(set(pm_varieties_map[pm])))
        for pm in pm_blocks_map:
            pm_blocks_map[pm] = sorted(list(set(pm_blocks_map[pm])))

        all_blocks_query = db.query(TpsrRecord.block).distinct().all()
        available_blocks = sorted(list(set(b[0].strip() for b in all_blocks_query if b[0] and str(b[0]).strip())))

        sws = db.query(TpsrRecord.source_week).distinct().all()
        all_source_weeks_raw = sorted(list(set(s[0] for s in sws if s[0])))
        extracted_years = sorted(
            list(set(int(str(s[0])[:4]) for s in sws if s[0] and len(str(s[0])) >= 4)),
            reverse=True,
        )
        available_years = [str(y) for y in extracted_years]

        selected_weeks_set: set[int] = set()
        if isinstance(weeks_filter, list):
            for w in weeks_filter:
                w_str = str(w).strip()
                if w_str and w_str.upper() not in ("ALL", "TODAS", "(TODAS)", "(TODAS LAS SEMANAS)"):
                    for sub in w_str.split(","):
                        sub_str = sub.strip()
                        if sub_str:
                            c = parse_week_code(sub_str)
                            if c: selected_weeks_set.add(c)
        elif isinstance(weeks_filter, str) and weeks_filter.strip():
            w_str = weeks_filter.strip()
            if w_str.upper() not in ("ALL", "TODAS", "(TODAS)", "(TODAS LAS SEMANAS)"):
                for sub in w_str.split(","):
                    sub_str = sub.strip()
                    if sub_str:
                        c = parse_week_code(sub_str)
                        if c: selected_weeks_set.add(c)

        parsed_sw_from = parse_week_code(week_from) if week_from and str(week_from).strip() else None
        parsed_sw_to = parse_week_code(week_to) if week_to and str(week_to).strip() else None
        if parsed_sw_from and parsed_sw_to and parsed_sw_from > parsed_sw_to:
            parsed_sw_from, parsed_sw_to = parsed_sw_to, parsed_sw_from
        # 2. Cycle Definitions Map & Valid Curve PMs
        cycles_list = db.query(CycleDefinitionDB).all()
        cycles_map: dict[tuple[str, str, str], dict[str, Any]] = {}
        valid_curve_pms: set[str] = set()
        for c in cycles_list:
            c_pm_norm = "VERONICA" if "VERONICA" in c.product_master_norm else c.product_master_norm
            c_pm = "VERONICA" if "VERONICA" in (c.product_master or "").upper() else c.product_master
            key = (c_pm_norm, c.variety_norm, normalize_text(c.activity))
            c_curve = parse_curve_csv(c.curve)
            cycles_map[key] = {
                "cycle_weeks": c.cycle_weeks,
                "waste_rate": c.waste_rate,
                "stems_per_plant": c.stems_per_plant,
                "curve": c_curve,
            }
            cycles_map[(c.product_master_norm, c.variety_norm, normalize_text(c.activity))] = cycles_map[key]
            if c_pm and c_curve and any(x > 0 for x in c_curve):
                valid_curve_pms.add(c_pm)
                valid_curve_pms.add(c.product_master)

        available_pms = sorted([pm for pm in pm_varieties_map.keys() if pm in valid_curve_pms])
        available_varieties = sorted(list(set(v for vars in pm_varieties_map.values() for v in vars)))

        # Resolve defaults:
        # Default year is 2026 (or most recent available year)
        if year_filter is None:
            resolved_year = "2026" if "2026" in available_years else (available_years[0] if available_years else "2026")
        elif str(year_filter).strip().upper() in ("", "ALL"):
            resolved_year = ""
        else:
            resolved_year = str(year_filter).strip()

        # Build year to weeks mapping and filter available_weeks for resolved_year
        year_weeks_map: dict[str, list[str]] = {}
        for w in all_source_weeks_raw:
            y_str = str(w)[:4]
            year_weeks_map.setdefault(y_str, []).append(format_short_week(w))
        for y_str in year_weeks_map:
            year_weeks_map[y_str] = sorted(list(set(year_weeks_map[y_str])))

        if resolved_year and resolved_year in year_weeks_map:
            available_weeks = year_weeks_map[resolved_year]
        else:
            available_weeks = [format_short_week(w) for w in all_source_weeks_raw]

        # Default product master is first available product (same logic as agronomo view)
        if pm_filter is None:
            resolved_pm = available_pms[0] if available_pms else "HYPERICUM"
        elif str(pm_filter).strip().upper() in ("", "ALL"):
            resolved_pm = ""
        else:
            resolved_pm = str(pm_filter).strip()

        selected_varieties = []
        if isinstance(variety_filter, list):
            selected_varieties = [str(v).strip() for v in variety_filter if v and str(v).strip() and str(v).strip().upper() not in ("ALL", "TODAS")]
        elif isinstance(variety_filter, str) and variety_filter.strip():
            parts = [p.strip() for p in variety_filter.split(",") if p.strip()]
            selected_varieties = [v for v in parts if v.upper() not in ("ALL", "TODAS")]

        selected_varieties_norm = [normalize_text(v) for v in selected_varieties]

        activity_filter = (activity_filter or "").strip()
        if activity_filter.upper() in ("ALL", "TODAS"):
            activity_filter = ""

        block_filter = (block_filter or "").strip()
        if block_filter.upper() in ("ALL", "TODOS", "(TODOS LOS BLOQUES)"):
            block_filter = ""

        # 3. Query Block Closures across all crops (unrestricted to preserve multi-crop lifecycle)
        closure_query = db.query(BlockClosureDB)
        closures_list = closure_query.all()
        closed_blocks_set: set[tuple[str, str, int, str, str]] = set()
        closed_blocks_by_lot: set[tuple[str, int, str, str]] = set()
        for c in closures_list:
            if c.is_closed:
                c_pm = normalize_text(c.product_master_norm)
                c_var = normalize_text(c.variety_norm)
                c_ac = normalize_text(c.activity)
                c_blk = normalize_text(c.block_norm or c.block)
                closed_blocks_set.add((c_pm, c_var, c.source_week, c_ac, c_blk))
                closed_blocks_by_lot.add((c_pm, c.source_week, c_ac, c_blk))

        # 4. Query TPSR records across all crops
        tpsr_records = db.query(TpsrRecord).all()

        # 5. Query Week Adjustments across all crops
        week_adjustments = db.query(WeekAdjustmentDB).all()
        adj_map_by_block: dict[tuple[str, str, int, str, str], list[WeekAdjustmentDB]] = {}
        for adj in week_adjustments:
            adj_pm_norm = "VERONICA" if "VERONICA" in adj.product_master_norm else adj.product_master_norm
            k = (adj_pm_norm, adj.variety_norm, adj.source_week, adj.activity, adj.block_norm)
            adj_map_by_block.setdefault(k, []).append(adj)
            adj_map_by_block.setdefault((adj.product_master_norm, adj.variety_norm, adj.source_week, adj.activity, adj.block_norm), []).append(adj)

        # 6. Group TPSR records by individual block: (product_master, variety, source_week, activity, block)
        grouped_block_lots: dict[tuple[str, str, int, str, str], dict[str, Any]] = {}
        for r in tpsr_records:
            pm = "VERONICA" if "VERONICA" in (r.product_master or "").upper() else r.product_master
            var = r.variety

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

        # 7. Process individual block rows and calculate metrics for all lots
        all_flat_block_rows: list[dict[str, Any]] = []

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
            adj_dump_map = {a.harvest_week: (a.dump_stems or 0) for a in adjs}

            is_closed = (k_norm in closed_blocks_set) or ((pm_norm, sw, ac_norm, blk_norm) in closed_blocks_by_lot)

            # Build complete curve weeks: if closed, ONLY real/agro/dump weeks are included; never fake model weeks
            all_hw_set = set()
            curve_hw_map = {}
            if not is_closed:
                for idx, pct in enumerate(c_def["curve"]):
                    hw = shift_week(sw, effective_cycle_weeks + idx)
                    all_hw_set.add(hw)
                    curve_hw_map[hw] = round(base_exportable * pct)

            for hw in adj_agro_map:
                all_hw_set.add(hw)
            for hw in adj_real_map:
                all_hw_set.add(hw)
            for hw, d_val in adj_dump_map.items():
                if d_val > 0:
                    all_hw_set.add(hw)

            harvest_breakdown = []
            block_total_production = 0
            block_real_stems = 0
            block_dumps = 0

            for hw in sorted(all_hw_set):
                hw_short = format_short_week(hw)
                val = 0
                src = "MODELO"
                if hw in adj_real_map:
                    val = adj_real_map[hw]
                    src = "REAL"
                    if val > 0:
                        block_real_stems += val
                elif hw in adj_agro_map:
                    val = adj_agro_map[hw]
                    # Historical closures or weeks prior to week 2630 must always show as REAL
                    if is_closed or hw < 202630:
                        src = "REAL"
                        if val > 0:
                            block_real_stems += val
                    else:
                        src = "AGRONOMO"
                elif not is_closed:
                    if hw < 202630:
                        val = 0
                        src = "REAL"
                    else:
                        val = curve_hw_map.get(hw, 0)
                        src = "MODELO"
                else:
                    val = 0
                    src = "REAL"

                d_stems = adj_dump_map.get(hw, 0)
                hw_total = val + d_stems

                # Skip zero-stem weeks: if there is no harvest (explicitly zeroed out or no production), it should not be in the curve
                if hw_total <= 0:
                    continue

                block_dumps += d_stems
                block_total_production += hw_total
                hw_pp = round(hw_total / plants, 2) if plants > 0 else 0.0

                harvest_breakdown.append({
                    "harvest_week": hw,
                    "harvest_week_short": hw_short,
                    "stems": hw_total,
                    "exportable_stems": val,
                    "dump_stems": d_stems,
                    "stems_pp": hw_pp,
                    "source": src,
                })

            for item in harvest_breakdown:
                item["pct_of_total"] = round((item["stems"] / block_total_production * 100.0), 1) if block_total_production > 0 else 0.0

            has_real_data = (block_real_stems > 0) or (block_dumps > 0)
            block_t_pl = round(block_total_production / plants, 1) if plants > 0 else 0.0
            is_below_target = (block_t_pl < effective_stems_pp)

            real_cycle_weeks = effective_cycle_weeks
            if adj_real_map:
                min_hw = min(adj_real_map.keys())
                real_cycle_weeks = max(1, calculate_week_difference(sw, min_hw))

            if plants >= 50:
                # Skip abandoned historical phantom lots (pre-2026 with no closure and no harvest stems)
                if sw < 202601 and not is_closed and not has_real_data:
                    continue

                block_dumps = sum(item["dump_stems"] for item in harvest_breakdown)
                harvest_weeks_str = ", ".join(item["harvest_week_short"] for item in harvest_breakdown)
                harvest_weeks_display = format_week_range_display([item["harvest_week_short"] for item in harvest_breakdown])
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
                    "dump_stems": block_dumps,
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
                    "harvest_weeks_display": harvest_weeks_display,
                    "harvest_breakdown": harvest_breakdown,
                }
                all_flat_block_rows.append(block_item)

        # Build consolidated productions per physical block across the entire farm
        # (This ensures every physical block retains its complete multi-crop and multi-year lifecycle,
        # without being truncated by active crop or week slicers)
        consolidated_blocks, consolidated_productions = build_consolidated_block_productions(all_flat_block_rows)

        # Filter block rows for Tab 1 (Visual Dashboard KPIs & Curves) and Tab 2 (Hierarchical Weekly Matrix)
        flat_block_rows: list[dict[str, Any]] = []
        chart_candidate_rows: list[dict[str, Any]] = []

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

        for b in all_flat_block_rows:
            # Crop match for Tab 1 & Tab 2
            pm_matches = True
            if resolved_pm:
                if "VERONICA" in resolved_pm.upper():
                    pm_matches = ("VERONICA" in b["product_master"].upper())
                else:
                    pm_matches = (b["product_master"] == resolved_pm)
            if not pm_matches:
                continue

            # Variety match
            if selected_varieties_norm and normalize_text(b["variety"]) not in selected_varieties_norm:
                continue

            # Activity match
            if activity_filter and b["activity"] != activity_filter:
                continue

            # Block match
            if block_filter and normalize_text(b["block"]) != normalize_text(block_filter):
                continue

            # Candidate for multi-year chart (across all available years for this crop/variety)
            chart_candidate_rows.append(b)

            # Year filter for operational active view
            if resolved_year:
                try:
                    y_int = int(resolved_year)
                    sw_val = b.get("source_week") or 0
                    b_year = b.get("year") or (sw_val // 100)
                    if b_year != y_int:
                        continue
                except ValueError:
                    pass

            # Week range or list filter
            if selected_weeks_set:
                if b["source_week"] not in selected_weeks_set:
                    continue
            else:
                if parsed_sw_from and b["source_week"] < parsed_sw_from:
                    continue
                if parsed_sw_to and b["source_week"] > parsed_sw_to:
                    continue

            flat_block_rows.append(b)

            # Accumulate metrics for active filtered scope
            total_plants_sum += b["plants"]
            total_production_sum += b["total_production"]
            total_real_closed_sum += b["real_stems"]
            total_dumps_sum += b["dump_stems"]
            real_stems_per_plant_list.append(b["real_stems_per_plant"])
            ideal_stems_per_plant_list.append(b["ideal_stems_per_plant"])
            waste_list.append(b["waste_pct"])
            cycle_real_list.append(b["real_cycle_weeks"])

            # Chronological harvest accumulation
            for item in b.get("harvest_breakdown", []):
                hw_short = item["harvest_week_short"]
                if item["source"] == "REAL":
                    chronological_real_stems[hw_short] = chronological_real_stems.get(hw_short, 0) + item["exportable_stems"]
                elif item["source"] == "AGRONOMO":
                    chronological_agronomo_stems[hw_short] = chronological_agronomo_stems.get(hw_short, 0) + item["exportable_stems"]

            # Relative curve accumulation for closed blocks
            if b["is_closed"] and b["real_stems"] > 0:
                real_hws = [item for item in b.get("harvest_breakdown", []) if item["source"] == "REAL"]
                if real_hws:
                    start_hw = real_hws[0]["harvest_week"]
                    for h_item in real_hws:
                        offset = max(0, min(11, calculate_week_difference(start_hw, h_item["harvest_week"])))
                        pct = (h_item["exportable_stems"] / b["real_stems"]) * 100.0 if b["real_stems"] > 0 else 0.0
                        relative_curve_real_sums[offset] += pct
                        relative_curve_real_counts[offset] += 1

            pm_norm = normalize_text(b["product_master"])
            var_norm = normalize_text(b["variety"])
            ac_norm = normalize_text(b["activity"])
            c_def_b = find_cycle_definition(cycles_map, pm_norm, var_norm, ac_norm)
            for idx, pct in enumerate(c_def_b["curve"]):
                if idx < 12:
                    relative_curve_ideal_sums[idx] += (pct * 100.0)
                    relative_curve_ideal_counts[idx] += 1

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
                        "source_week": sw,
                        "total_plants": 0,
                        "total_production": 0,
                        "siem_plants": 0,
                        "siem_prod": 0,
                        "siem_blocks": 0,
                        "poda_plants": 0,
                        "poda_prod": 0,
                        "poda_blocks": 0,
                        "ideal_target": sw_ideal_t_pl,
                        "blocks_count": 0,
                        "closed_blocks": 0,
                    }
                chart3_weeks_map[chart3_label]["total_plants"] += sw_plants
                chart3_weeks_map[chart3_label]["total_production"] += sw_prod
                chart3_weeks_map[chart3_label]["blocks_count"] += len(blocks)
                chart3_weeks_map[chart3_label]["closed_blocks"] += sw_closed
                if ac == "SIEMBRA":
                    chart3_weeks_map[chart3_label]["siem_plants"] += sw_plants
                    chart3_weeks_map[chart3_label]["siem_prod"] += sw_prod
                    chart3_weeks_map[chart3_label]["siem_blocks"] += len(blocks)
                elif ac == "PODA":
                    chart3_weeks_map[chart3_label]["poda_plants"] += sw_plants
                    chart3_weeks_map[chart3_label]["poda_prod"] += sw_prod
                    chart3_weeks_map[chart3_label]["poda_blocks"] += len(blocks)

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

        # Overall KPIs (Strictly for CLOSED blocks as requested by user)
        closed_blocks = [r for r in flat_block_rows if r["is_closed"]]
        closed_plants_sum = sum(r["plants"] for r in closed_blocks)
        closed_prod_sum = sum(r["total_production"] for r in closed_blocks)
        closed_real_stems_sum = sum(r["real_stems"] for r in closed_blocks)
        closed_dumps_sum = sum(r["dump_stems"] for r in closed_blocks)

        closed_stems_pp_list = [r["real_stems_per_plant"] for r in closed_blocks]
        closed_ideal_pp_list = [r["ideal_stems_per_plant"] for r in closed_blocks]
        closed_cycle_list = [r["real_cycle_weeks"] for r in closed_blocks]
        closed_waste_list = [r["waste_pct"] for r in closed_blocks]

        avg_real_stems_pp = round(closed_prod_sum / closed_plants_sum, 1) if closed_plants_sum > 0 else 0.0
        avg_ideal_stems_pp = round(sum(closed_ideal_pp_list) / len(closed_ideal_pp_list), 1) if closed_ideal_pp_list else (round(sum(ideal_stems_per_plant_list) / len(ideal_stems_per_plant_list), 1) if ideal_stems_per_plant_list else (4.0 if (pm_filter and "VERONICA" in pm_filter.upper()) else 9.5))
        avg_cycle = round(sum(closed_cycle_list) / len(closed_cycle_list), 1) if closed_cycle_list else 0.0
        avg_waste = round(sum(closed_waste_list) / len(closed_waste_list), 1) if closed_waste_list else 10.0
        total_closed_blocks = len(closed_blocks)

        # Siembra vs Poda individual breakdown for closed blocks
        s_closed = [r for r in closed_blocks if r.get("activity") == "SIEMBRA"]
        p_closed = [r for r in closed_blocks if r.get("activity") == "PODA"]

        s_plants = sum(r["plants"] for r in s_closed)
        s_prod = sum(r["total_production"] for r in s_closed)
        s_tp = round(s_prod / s_plants, 2) if s_plants > 0 else 0.0
        s_cycle = round(sum(r["real_cycle_weeks"] for r in s_closed) / len(s_closed), 1) if s_closed else 0.0
        s_waste = round(sum(r["waste_pct"] for r in s_closed) / len(s_closed), 1) if s_closed else 0.0

        p_plants = sum(r["plants"] for r in p_closed)
        p_prod = sum(r["total_production"] for r in p_closed)
        p_tp = round(p_prod / p_plants, 2) if p_plants > 0 else 0.0
        p_cycle = round(sum(r["real_cycle_weeks"] for r in p_closed) / len(p_closed), 1) if p_closed else 0.0
        p_waste = round(sum(r["waste_pct"] for r in p_closed) / len(p_closed), 1) if p_closed else 0.0

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

        # Chart 3: Multi-Year Real Stems / Plant Comparison by Calendar Week (Format 45, 46 etc. comparing 2026 vs 2025 vs 2024)
        year_week_stems: dict[str, dict[int, dict[str, Any]]] = {}
        all_calendar_weeks_set: set[int] = set()

        for b in chart_candidate_rows:
            if not b["is_closed"]:
                continue
            sw_val = b.get("source_week")
            if not sw_val:
                continue
            sw_str = str(sw_val)
            if len(sw_str) < 6:
                continue
            try:
                y_str = sw_str[:4]
                wk_num = int(sw_str[-2:])
            except ValueError:
                continue

            b_plants = b.get("plants", 0)
            b_prod = b.get("total_production", 0)
            if b_plants <= 0 or b_prod <= 0:
                continue

            all_calendar_weeks_set.add(wk_num)
            if y_str not in year_week_stems:
                year_week_stems[y_str] = {}
            if wk_num not in year_week_stems[y_str]:
                year_week_stems[y_str][wk_num] = {
                    "plants": 0,
                    "prod": 0,
                    "siem_plants": 0,
                    "siem_prod": 0,
                    "poda_plants": 0,
                    "poda_prod": 0,
                    "blocks_count": 0,
                }
            acc = year_week_stems[y_str][wk_num]
            acc["plants"] += b_plants
            acc["prod"] += b_prod
            acc["blocks_count"] += 1
            if b.get("activity") == "SIEMBRA":
                acc["siem_plants"] += b_plants
                acc["siem_prod"] += b_prod
            elif b.get("activity") == "PODA":
                acc["poda_plants"] += b_plants
                acc["poda_prod"] += b_prod

        # Filter by selected weeks if specified
        if selected_weeks_set:
            wk_filter_set = {int(str(w)[-2:]) for w in selected_weeks_set if len(str(w)) >= 2}
            sorted_calendar_weeks = sorted([w for w in all_calendar_weeks_set if w in wk_filter_set])
        else:
            sorted_calendar_weeks = sorted(list(all_calendar_weeks_set))

        chart3_sorted_labels = [f"Sem {w}" for w in sorted_calendar_weeks]
        series_by_year: dict[str, list[float | None]] = {}
        details_by_year: dict[str, list[dict[str, Any] | None]] = {}
        available_chart_years = sorted(list(year_week_stems.keys()), reverse=True)

        for y_str in available_chart_years:
            y_data = []
            y_details = []
            for wk_num in sorted_calendar_weeks:
                if wk_num in year_week_stems[y_str]:
                    cell = year_week_stems[y_str][wk_num]
                    t_pl = round(cell["prod"] / cell["plants"], 2) if cell["plants"] > 0 else None
                    y_data.append(t_pl)
                    siem_pp = round(cell["siem_prod"] / cell["siem_plants"], 2) if cell["siem_plants"] > 0 else None
                    poda_pp = round(cell["poda_prod"] / cell["poda_plants"], 2) if cell["poda_plants"] > 0 else None
                    y_details.append({
                        "year": y_str,
                        "week": wk_num,
                        "total_t_pl": t_pl,
                        "siem_t_pl": siem_pp,
                        "poda_t_pl": poda_pp,
                        "total_plants": cell["plants"],
                        "total_production": cell["prod"],
                        "blocks_count": cell["blocks_count"],
                        "target": avg_ideal_stems_pp,
                    })
                else:
                    y_data.append(None)
                    y_details.append(None)
            series_by_year[y_str] = y_data
            details_by_year[y_str] = y_details

        primary_year = resolved_year if resolved_year in series_by_year else ("2026" if "2026" in series_by_year else (available_chart_years[0] if available_chart_years else None))
        chart_stems_pp_real = series_by_year.get(primary_year, []) if primary_year else []
        chart_stems_pp_target = [avg_ideal_stems_pp] * len(chart3_sorted_labels)
        chart_stems_pp_details = details_by_year.get(primary_year, []) if primary_year else []
        last_closed_label = chart3_sorted_labels[-1] if chart3_sorted_labels else None
        last_closed_index = len(chart3_sorted_labels) - 1 if chart3_sorted_labels else -1

        # Chart 4: Hybrid Lifecycle Curve (Closed Real Stems + Future Projected Estimates)
        hybrid_weeks = sorted_chrono_weeks
        chart_hybrid_real = [chronological_real_stems.get(w, 0) for w in hybrid_weeks]
        chart_hybrid_projected = [chronological_agronomo_stems.get(w, 0) for w in hybrid_weeks]

        harvest_breakdown_map: dict[str, list[dict[str, Any]]] = {}
        for b in all_flat_block_rows:
            uid = f"{b['product_master']}-{b['source_week']}-{b['activity']}-{b['block']}"
            if b.get("harvest_breakdown"):
                harvest_breakdown_map[uid] = [
                    {
                        "hw": h["harvest_week_short"],
                        "sw": b["source_week_short"],
                        "stems": h["stems"],
                        "pp": h["stems_pp"],
                        "dump": h.get("dump_stems", 0),
                        "pct": h.get("pct_of_total", 0),
                        "src": h.get("source", "MODELO"),
                        "activity": b.get("activity", "SIEMBRA"),
                    }
                    for h in b["harvest_breakdown"]
                    if h.get("stems", 0) > 0
                ]

        for prod in consolidated_productions:
            uid = prod["uid"]
            if prod.get("harvest_breakdown"):
                harvest_breakdown_map[uid] = [
                    {
                        "hw": h["harvest_week_short"],
                        "sw": prod["source_week_max_short"],
                        "stems": h["stems"],
                        "pp": h["stems_pp"],
                        "dump": h.get("dump_stems", 0),
                        "pct": h.get("pct_of_total", 0),
                        "src": h.get("source", "MODELO"),
                        "activity": prod.get("activity", "SIEMBRA"),
                        "production_label": prod.get("production_label", ""),
                        "production_num": prod.get("production_num", 1),
                    }
                    for h in prod["harvest_breakdown"]
                    if h.get("stems", 0) > 0
                ]
            for vb in prod.get("variety_breakdown", []):
                v_key = f"{uid}_var_{vb['variety']}"
                if vb.get("harvest_breakdown"):
                    harvest_breakdown_map[v_key] = [
                        {
                            "hw": h["harvest_week_short"],
                            "sw": prod["source_week_max_short"],
                            "stems": h["stems"],
                            "pp": h["stems_pp"],
                            "dump": h.get("dump_stems", 0),
                            "pct": h.get("pct_of_total", 0),
                            "src": h.get("source", "MODELO"),
                            "activity": prod.get("activity", "SIEMBRA"),
                            "variety": vb["variety"],
                            "production_label": prod.get("production_label", ""),
                        }
                        for h in vb["harvest_breakdown"]
                        if h.get("stems", 0) > 0
                    ]

        for cblock in consolidated_blocks:
            cuid = cblock["uid"]
            combined_hw = {}
            for prod in cblock["productions"]:
                for h in prod.get("harvest_breakdown", []):
                    if h.get("stems", 0) <= 0:
                        continue
                    hw_short = h["harvest_week_short"]
                    if hw_short not in combined_hw:
                        combined_hw[hw_short] = {
                            "hw": hw_short,
                            "sw": cblock["source_week_max_short"],
                            "stems": 0,
                            "dump": 0,
                            "src": h.get("source", "MODELO"),
                            "activity": prod.get("activity", "SIEMBRA"),
                            "production_label": prod.get("production_label", ""),
                        }
                    combined_hw[hw_short]["stems"] += h["stems"]
                    combined_hw[hw_short]["dump"] += h.get("dump_stems", 0)
                    if h.get("source") == "REAL":
                        combined_hw[hw_short]["src"] = "REAL"
                    elif h.get("source") == "AGRONOMO" and combined_hw[hw_short]["src"] != "REAL":
                        combined_hw[hw_short]["src"] = "AGRONOMO"

            tot_b_prod = cblock["total_production"]
            tot_b_plants = cblock["total_plants"]
            sorted_cb_hw = sorted([item for item in combined_hw.values() if item["stems"] > 0], key=lambda x: int(x["hw"]) if x["hw"].isdigit() else 999)
            for item in sorted_cb_hw:
                item["pct"] = round((item["stems"] / tot_b_prod * 100.0), 1) if tot_b_prod > 0 else 0.0
                item["pp"] = round(item["stems"] / tot_b_plants, 2) if tot_b_plants > 0 else 0.0
            
            harvest_breakdown_map[cuid] = sorted_cb_hw

            # Also register per-variety combined breakdown across all cycles in the block
            cblock_var_map: dict[str, dict[str, Any]] = {}
            for prod in cblock["productions"]:
                for vb in prod.get("variety_breakdown", []):
                    v_name = vb["variety"]
                    if v_name not in cblock_var_map:
                        cblock_var_map[v_name] = {"plants": 0, "prod": 0, "hw": {}}
                    cblock_var_map[v_name]["plants"] = max(cblock_var_map[v_name]["plants"], vb.get("plants", 0))
                    cblock_var_map[v_name]["prod"] += vb.get("production", 0)
                    for h in vb.get("harvest_breakdown", []):
                        if h.get("stems", 0) <= 0:
                            continue
                        hws = h["harvest_week_short"]
                        if hws not in cblock_var_map[v_name]["hw"]:
                            cblock_var_map[v_name]["hw"][hws] = {
                                "hw": hws,
                                "sw": cblock["source_week_max_short"],
                                "stems": 0,
                                "dump": 0,
                                "src": h.get("source", "MODELO"),
                                "activity": prod.get("activity", "SIEMBRA"),
                                "variety": v_name,
                            }
                        entry = cblock_var_map[v_name]["hw"][hws]
                        entry["stems"] += h["stems"]
                        entry["dump"] += h.get("dump_stems", 0)
                        if h.get("source") == "REAL":
                            entry["src"] = "REAL"
                        elif h.get("source") == "AGRONOMO" and entry["src"] != "REAL":
                            entry["src"] = "AGRONOMO"

            for v_name, cv_info in cblock_var_map.items():
                v_key = f"{cuid}_var_{v_name}"
                sorted_vhw = sorted(cv_info["hw"].values(), key=lambda x: int(x["hw"]) if x["hw"].isdigit() else 999)
                cv_plants = cv_info["plants"]
                cv_prod = cv_info["prod"]
                for item in sorted_vhw:
                    item["pp"] = round(item["stems"] / cv_plants, 2) if cv_plants > 0 else 0.0
                    item["pct"] = round((item["stems"] / cv_prod * 100.0), 1) if cv_prod > 0 else 0.0
                harvest_breakdown_map[v_key] = sorted_vhw

            # Also register per-crop combined breakdown across cycles of that crop in the block
            cblock_crop_map: dict[str, dict[str, Any]] = {}
            for prod in cblock["productions"]:
                crop_name = prod.get("product_master") or "OTRO"
                if crop_name not in cblock_crop_map:
                    cblock_crop_map[crop_name] = {"plants": 0, "prod": 0, "hw": {}}
                cblock_crop_map[crop_name]["plants"] += prod.get("plants", 0)
                cblock_crop_map[crop_name]["prod"] += prod.get("total_production", 0)
                for h in prod.get("harvest_breakdown", []):
                    if h.get("stems", 0) <= 0:
                        continue
                    hws = h["harvest_week_short"]
                    if hws not in cblock_crop_map[crop_name]["hw"]:
                        cblock_crop_map[crop_name]["hw"][hws] = {
                            "hw": hws,
                            "sw": prod.get("source_week_max_short", ""),
                            "stems": 0,
                            "dump": 0,
                            "src": h.get("source", "MODELO"),
                            "activity": prod.get("activity", "SIEMBRA"),
                            "product_master": crop_name,
                        }
                    entry = cblock_crop_map[crop_name]["hw"][hws]
                    entry["stems"] += h["stems"]
                    entry["dump"] += h.get("dump_stems", 0)
                    if h.get("source") == "REAL":
                        entry["src"] = "REAL"
                    elif h.get("source") == "AGRONOMO" and entry["src"] != "REAL":
                        entry["src"] = "AGRONOMO"

            for crop_name, cc_info in cblock_crop_map.items():
                crop_key = f"{cuid}_crop_{crop_name}"
                sorted_chw = sorted(cc_info["hw"].values(), key=lambda x: int(x["hw"]) if x["hw"].isdigit() else 999)
                cc_plants = cc_info["plants"]
                cc_prod = cc_info["prod"]
                for item in sorted_chw:
                    item["pp"] = round(item["stems"] / cc_plants, 2) if cc_plants > 0 else 0.0
                    item["pct"] = round((item["stems"] / cc_prod * 100.0), 1) if cc_prod > 0 else 0.0
                harvest_breakdown_map[crop_key] = sorted_chw

        res = {
            "kpis": {
                "total_real_stems": closed_prod_sum,
                "total_plants": closed_plants_sum,
                "avg_real_stems_per_plant": avg_real_stems_pp,
                "avg_ideal_stems_per_plant": avg_ideal_stems_pp,
                "avg_cycle_weeks": avg_cycle,
                "avg_waste_pct": avg_waste,
                "total_dumps": closed_dumps_sum,
                "total_lots": len(flat_block_rows),
                "total_closed_blocks": total_closed_blocks,
                "siembra": {
                    "plants": s_plants,
                    "production": s_prod,
                    "stems_per_plant": s_tp,
                    "cycle_weeks": s_cycle,
                    "waste_pct": s_waste,
                    "count": len(s_closed),
                },
                "poda": {
                    "plants": p_plants,
                    "production": p_prod,
                    "stems_per_plant": p_tp,
                    "cycle_weeks": p_cycle,
                    "waste_pct": p_waste,
                    "count": len(p_closed),
                },
            },
            "table_rows": flat_block_rows,
            "consolidated_blocks": consolidated_blocks,
            "consolidated_productions": consolidated_productions,
            "grouped_products": grouped_products,
            "harvest_breakdown_map": harvest_breakdown_map,
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
                "calendar_weeks": sorted_calendar_weeks,
                "series_by_year": series_by_year,
                "years": available_chart_years,
                "details_by_year": details_by_year,
                "real_stems_pp": chart_stems_pp_real,
                "target_stems_pp": chart_stems_pp_target,
                "details": chart_stems_pp_details,
                "last_closed_label": last_closed_label,
                "last_closed_index": last_closed_index,
            },
            "chart_hybrid": {
                "labels": hybrid_weeks,
                "real_stems": chart_hybrid_real,
                "projected_stems": chart_hybrid_projected,
            },
            "available_years": available_years,
            "available_weeks": available_weeks,
            "available_pms": available_pms,
            "available_varieties": available_varieties,
            "available_blocks": available_blocks,
            "pm_varieties_map": pm_varieties_map,
            "pm_blocks_map": pm_blocks_map,
            "year_weeks_map": year_weeks_map,
            "selected_year": resolved_year,
            "selected_weeks": [format_short_week(w) for w in sorted(selected_weeks_set)],
            "selected_weeks_str": ", ".join(format_short_week(w) for w in sorted(selected_weeks_set)),
            "selected_week_from": format_short_week(parsed_sw_from) if parsed_sw_from else (week_from or ""),
            "selected_week_to": format_short_week(parsed_sw_to) if parsed_sw_to else (week_to or ""),
            "selected_pm": resolved_pm,
            "selected_varieties": selected_varieties,
            "selected_variety": selected_varieties[0] if len(selected_varieties) == 1 else (", ".join(selected_varieties) if selected_varieties else ""),
            "selected_activity": activity_filter,
            "selected_block": block_filter,
        }
        _STATS_CACHE[cache_key] = res
        return res
    finally:
        db.close()


@app.route("/estadistica")
def estadistica_view() -> str:
    year_filter = request.args.get("year")
    weeks_list = request.args.getlist("semana")
    if not weeks_list and request.args.get("semana"):
        weeks_list = [w.strip() for w in request.args.get("semana").split(",") if w.strip()]
    week_from = request.args.get("week_from", "").strip()
    week_to = request.args.get("week_to", "").strip()
    pm_filter = request.args.get("pm")
    varieties = request.args.getlist("variedad")
    if not varieties and request.args.get("variedad"):
        varieties = [v.strip() for v in request.args.get("variedad").split(",") if v.strip()]
    activity_filter = request.args.get("ac", "").strip()
    block_filter = request.args.get("bloque", "").strip()

    stats_data = get_statistics_data(
        year_filter=year_filter,
        pm_filter=pm_filter,
        variety_filter=varieties,
        activity_filter=activity_filter,
        block_filter=block_filter,
        weeks_filter=weeks_list,
        week_from=week_from,
        week_to=week_to,
    )

    return render_template(
        "estadistica.html",
        stats=stats_data,
        workbook_path="PostgreSQL Database",
        active_view="estadistica",
    )


@app.route("/api/estadistica")
def estadistica_api():
    year_filter = request.args.get("year")
    weeks_list = request.args.getlist("semana")
    if not weeks_list and request.args.get("semana"):
        weeks_list = [w.strip() for w in request.args.get("semana").split(",") if w.strip()]
    week_from = request.args.get("week_from", "").strip()
    week_to = request.args.get("week_to", "").strip()
    pm_filter = request.args.get("pm")
    varieties = request.args.getlist("variedad")
    if not varieties and request.args.get("variedad"):
        varieties = [v.strip() for v in request.args.get("variedad").split(",") if v.strip()]
    activity_filter = request.args.get("ac", "").strip()
    block_filter = request.args.get("bloque", "").strip()

    stats_data = get_statistics_data(
        year_filter=year_filter,
        pm_filter=pm_filter,
        variety_filter=varieties,
        activity_filter=activity_filter,
        block_filter=block_filter,
        weeks_filter=weeks_list,
        week_from=week_from,
        week_to=week_to,
    )
    return jsonify(stats_data)


@app.route("/api/estadistica/export-csv")
def estadistica_export_csv():
    year_filter = request.args.get("year")
    weeks_list = request.args.getlist("semana")
    if not weeks_list and request.args.get("semana"):
        weeks_list = [w.strip() for w in request.args.get("semana").split(",") if w.strip()]
    week_from = request.args.get("week_from", "").strip()
    week_to = request.args.get("week_to", "").strip()
    pm_filter = request.args.get("pm")
    varieties = request.args.getlist("variedad")
    if not varieties and request.args.get("variedad"):
        varieties = [v.strip() for v in request.args.get("variedad").split(",") if v.strip()]
    activity_filter = request.args.get("ac", "").strip()
    block_filter = request.args.get("bloque", "").strip()

    stats_data = get_statistics_data(
        year_filter=year_filter,
        pm_filter=pm_filter,
        variety_filter=varieties,
        activity_filter=activity_filter,
        block_filter=block_filter,
        weeks_filter=weeks_list,
        week_from=week_from,
        week_to=week_to,
    )

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "Producto Maestro",
        "Semana Origen",
        "Actividad",
        "Bloque",
        "Variedad",
        "Plantas",
        "Tallos Exportables",
        "Total Produccion",
        "Tallos DUMP",
        "Real Tallos/Planta",
        "Ideal Tallos/Planta",
        "Ciclo Real (sem)",
        "Ciclo Ideal (sem)",
        "Desperdicio %",
        "Semanas Cosecha",
        "Estado",
        "Detalle Cosechas Semanales",
    ])
    for r in stats_data.get("table_rows", []):
        harvest_detail = "; ".join([
            f"Sem {hw['harvest_week_short']}: {hw['stems']} {('(' + str(hw['dump_stems']) + ' Dump)') if hw.get('dump_stems') else ''} [{hw['source']}]"
            for hw in r.get("harvest_breakdown", [])
        ])
        writer.writerow([
            r.get("product_master", ""),
            r.get("source_week_short", ""),
            r.get("activity", ""),
            r.get("block", ""),
            r.get("variety_display") or r.get("variety", ""),
            r.get("plants", 0),
            r.get("exportable_stems", 0),
            r.get("total_production", 0),
            r.get("dump_stems", 0),
            r.get("real_stems_per_plant", 0),
            r.get("ideal_stems_per_plant", 0),
            r.get("real_cycle_weeks", 0),
            r.get("ideal_cycle_weeks", 0),
            r.get("waste_pct", 0),
            r.get("harvest_weeks_str", ""),
            r.get("status", ""),
            harvest_detail,
        ])

    csv_bytes = ("\ufeff" + output.getvalue()).encode("utf-8")
    filename = f"reporte_estadistica_{stats_data.get('selected_year') or 'todos'}_{datetime.now().strftime('%Y%m%d')}.csv"
    return Response(
        csv_bytes,
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.route("/api/matriz/exportar")
def exportar_matriz_api():
    if not session.get("username"):
        return jsonify({"ok": False, "message": "Sesión requerida."}), 401

    requested_year_str = request.args.get("ano", request.args.get("year", "2026")).strip()
    try:
        requested_year = int(requested_year_str) if requested_year_str else 2026
    except ValueError:
        requested_year = 2026

    snapshot: dict[str, Any] = get_snapshot(force_refresh=False, year=requested_year)
    prods = available_products(snapshot)

    requested_week = request.args.get("semana", "").strip()
    selected_week = resolve_selected_week(requested_week, snapshot)

    requested_product_master = request.args.get("producto", "").strip()
    if not requested_product_master and prods:
        requested_product_master = prods[0]

    requested_horizon_back = request.args.get("horizonte_atras", "")
    requested_horizon_forward = request.args.get("horizonte_adelante", "")
    formato = request.args.get("formato", "excel").lower().strip()

    filter_var = request.args.get("variedad", "").strip().upper()
    filter_blk = request.args.get("bloque", "").strip().upper()
    filter_ac = request.args.get("ac", "").strip().upper()
    filter_ubi = request.args.get("ubi", "").strip().upper()
    filter_origen = request.args.get("origen", "").strip().upper()
    filter_cosecha = request.args.get("cosecha", "").strip().upper()
    filter_solo_cosechas = request.args.get("solo_cosechas", "0") == "1"

    horizon_weeks_back = max(0, min(20, int(requested_horizon_back or 4)))
    horizon_weeks_forward = max(1, min(20, int(requested_horizon_forward or 8)))

    weekly_view = aggregate_for_week(
        snapshot,
        selected_week,
        horizon_weeks_back,
        horizon_weeks_forward,
        selected_product_master="" if requested_product_master.upper() in {"TODOS", "ALL"} else requested_product_master,
    )

    # Flatten all rows from varieties (siembras + podas)
    matrix_rows = list(weekly_view.get("matrix_rows", []))
    if not matrix_rows:
        for v in weekly_view.get("varieties", []):
            matrix_rows.extend(v.get("siembras", []))
            matrix_rows.extend(v.get("podas", []))

    week_columns = list(weekly_view.get("week_columns", []))

    # Apply inline filters if present
    if filter_var:
        matrix_rows = [r for r in matrix_rows if filter_var in (r.get("variety_display") or r.get("variety", "")).upper()]
    if filter_blk:
        matrix_rows = [r for r in matrix_rows if filter_blk in str(r.get("block", "")).upper()]
    if filter_ac:
        matrix_rows = [r for r in matrix_rows if r.get("activity") == filter_ac or (filter_ac == 'S' and r.get("activity") == 'SIEMBRA') or (filter_ac == 'P' and r.get("activity") == 'PODA')]
    if filter_ubi:
        matrix_rows = [r for r in matrix_rows if filter_ubi in str(r.get("bed_location", "")).upper()]
    if filter_origen:
        matrix_rows = [r for r in matrix_rows if filter_origen in str(r.get("source_week_short", "")).upper()]
    if filter_cosecha:
        matrix_rows = [r for r in matrix_rows if filter_cosecha in str(r.get("harvest_start_week_short", "")).upper()]
    if filter_solo_cosechas:
        matrix_rows = [r for r in matrix_rows if r.get("window_total", 0) > 0]

    filename_base = f"Matriz_Agronomo_Sem{format_short_week(selected_week)}_{requested_product_master or 'TODOS'}"

    # Calculate totals as pure numerical values (never formulas)
    total_plants = sum(r.get("plants", 0) for r in matrix_rows)
    total_prog = sum(r.get("program_total", 0) for r in matrix_rows)
    total_window = sum(r.get("window_total", 0) for r in matrix_rows)
    total_prod = sum(r.get("total_production", 0) for r in matrix_rows)
    total_by_col = {
        w["label"]: sum(
            (r.get("weekly_projection", {}).get(w["label"], 0) or 0)
            + (r.get("dump_stems", {}).get(w["label"], 0) or 0)
            for r in matrix_rows
        )
        for w in week_columns
    }

    if formato == "csv":
        output = io.StringIO()
        writer = csv.writer(output, delimiter=";")
        fixed_headers = [
            "CIERRE",
            "AC",
            "UBI",
            "SEM. ORIGEN",
            "BLOQUE",
            "PRODUCTO",
            "VARIEDAD",
            "SEM. COSECHA",
            "PLANTAS",
            "CICLO",
            "T/PL",
            "TOT PROG",
            "TOT VENTANA",
            "TOT PRODUCCION",
        ]
        headers = fixed_headers + [f"SEM {w['label']}" for w in week_columns]
        writer.writerow(headers)

        for r in matrix_rows:
            cierre_txt = "CERRADO" if r.get("block_closed") else "ABIERTO"
            row_data = [
                cierre_txt,
                "S" if r.get("activity") == "SIEMBRA" else ("P" if r.get("activity") == "PODA" else r.get("activity", "")),
                r.get("bed_location", ""),
                r.get("source_week_short", ""),
                r.get("block", ""),
                r.get("product_master", ""),
                r.get("variety_display") or r.get("variety", ""),
                r.get("harvest_start_week_short", ""),
                r.get("plants", 0),
                r.get("cycle_weeks", ""),
                r.get("stems_per_plant", 0),
                r.get("program_total", 0),
                r.get("window_total", 0),
                r.get("total_production", 0),
            ]
            for w in week_columns:
                proj_val = r.get("weekly_projection", {}).get(w["label"], 0) or 0
                dump_val = r.get("dump_stems", {}).get(w["label"], 0) or 0
                row_data.append(proj_val + dump_val)
            writer.writerow(row_data)

        # Totals row in CSV
        totals_row = [
            "TOTALES", "", "", "", "", "", "", "",
            total_plants, "", "",
            total_prog, total_window, total_prod,
        ]
        for w in week_columns:
            totals_row.append(total_by_col.get(w["label"], 0))
        writer.writerow(totals_row)

        csv_bytes = ("\ufeff" + output.getvalue()).encode("utf-8")
        return Response(
            csv_bytes,
            mimetype="text/csv",
            headers={"Content-Disposition": f"attachment; filename={filename_base}.csv"},
        )
    else:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.comments import Comment
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Matriz Sem {format_short_week(selected_week)}"
        ws.views.sheetView[0].showGridLines = True

        navy_header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        cyan_week_fill = PatternFill(start_color="0369A1", end_color="0369A1", fill_type="solid")
        amber_cutoff_fill = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
        totals_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        closed_row_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

        white_bold = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        black_regular = Font(name="Calibri", size=9, bold=False, color="000000")
        totals_font = Font(name="Calibri", size=9, bold=True, color="92400E")

        thin_side = Side(border_style="thin", color="CBD5E1")
        cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        double_bottom = Border(left=thin_side, right=thin_side, top=thin_side, bottom=Side(border_style="double", color="92400E"))

        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        # Title & Meta Info
        ws.merge_cells("A1:G1")
        ws["A1"] = f"🌿 MATRIZ DE PROYECCIÓN Y COSECHA AGRÓNOMO — SEMANA {format_short_week(selected_week)}"
        ws["A1"].font = Font(name="Calibri", size=12, bold=True, color="0369A1")
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells("A2:N2")
        ws["A2"] = f"Producto: {requested_product_master or 'TODOS'} | Ventana: -{horizon_weeks_back} a +{horizon_weeks_forward} semanas | Bloques: {len(matrix_rows)} | Exportado: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        ws["A2"].font = Font(name="Calibri", size=8, italic=True, color="64748B")

        fixed_headers = [
            "CIERRE",
            "AC",
            "UBI",
            "SEM. ORIGEN",
            "BLOQUE",
            "PRODUCTO",
            "VARIEDAD",
            "SEM. COSECHA",
            "PLANTAS",
            "CICLO",
            "T/PL",
            "TOT PROG",
            "TOT VENTANA",
            "TOT PRODUCCIÓN",
        ]
        headers = fixed_headers + [f"SEM {w['label']}" for w in week_columns]
        header_row_idx = 4
        ws.row_dimensions[header_row_idx].height = 24

        for col_idx, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=header_row_idx, column=col_idx, value=h_text)
            cell.border = cell_border
            cell.alignment = align_center
            if col_idx <= len(fixed_headers):
                cell.fill = navy_header_fill
                cell.font = white_bold
            else:
                w_lbl = week_columns[col_idx - len(fixed_headers) - 1]["label"]
                if w_lbl == format_short_week(selected_week):
                    cell.fill = amber_cutoff_fill
                else:
                    cell.fill = cyan_week_fill
                cell.font = white_bold

        ws.freeze_panes = "H5"

        current_row_idx = header_row_idx + 1
        for r in matrix_rows:
            ws.row_dimensions[current_row_idx].height = 18
            is_closed = bool(r.get("block_closed"))
            cierre_txt = "CERRADO" if is_closed else "ABIERTO"
            ac_label = "S" if r.get("activity") == "SIEMBRA" else ("P" if r.get("activity") == "PODA" else r.get("activity", ""))
            row_vals = [
                cierre_txt,
                ac_label,
                r.get("bed_location", ""),
                int(r["source_week_short"]) if str(r.get("source_week_short", "")).isdigit() else r.get("source_week_short", ""),
                str(r.get("block", "")),
                r.get("product_master", ""),
                r.get("variety_display") or r.get("variety", ""),
                int(r["harvest_start_week_short"]) if str(r.get("harvest_start_week_short", "")).isdigit() else r.get("harvest_start_week_short", ""),
                int(r.get("plants", 0)),
                int(r["cycle_weeks"]) if r.get("cycle_weeks") is not None else "",
                float(r.get("stems_per_plant", 0.0)),
                int(r.get("program_total", 0)),
                int(r.get("window_total", 0)),
                int(r.get("total_production", 0)),
            ]
            week_comments = {}
            for w_idx, w in enumerate(week_columns):
                proj_val = r.get("weekly_projection", {}).get(w["label"], 0) or 0
                dump_val = r.get("dump_stems", {}).get(w["label"], 0) or 0
                total_val = proj_val + dump_val
                row_vals.append(int(total_val) if total_val else 0)
                if dump_val > 0:
                    target_col = len(fixed_headers) + 1 + w_idx
                    if proj_val > 0:
                        week_comments[target_col] = f"Cosecha: {proj_val:,} + Descarte (Dump): {dump_val:,} = Total: {total_val:,}"
                    else:
                        week_comments[target_col] = f"Descarte (Dump): {dump_val:,}"

            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=current_row_idx, column=col_idx, value=val)
                cell.border = cell_border
                cell.font = black_regular
                if is_closed:
                    cell.fill = closed_row_fill

                if col_idx in week_comments:
                    cell.comment = Comment(week_comments[col_idx], "Sistema Proyecciones")

                if col_idx in [1, 2, 4, 8]:
                    cell.alignment = align_center
                elif col_idx in [3, 5, 6, 7]:
                    cell.alignment = align_left
                elif col_idx == 11:
                    cell.alignment = align_right
                    cell.number_format = "0.0"
                elif col_idx in [9, 10, 12, 13, 14] or col_idx > len(fixed_headers):
                    cell.alignment = align_right
                    cell.number_format = "#,##0"

            current_row_idx += 1

        # Totals row: pure numerical values (no formulas to avoid circular references)
        ws.row_dimensions[current_row_idx].height = 20

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=current_row_idx, column=col_idx)
            cell.border = double_bottom
            cell.fill = totals_fill
            cell.font = totals_font

            if col_idx == 1:
                cell.value = "TOTALES"
                cell.alignment = align_left
            elif col_idx == 9:
                cell.value = total_plants
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif col_idx == 12:
                cell.value = total_prog
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif col_idx == 13:
                cell.value = total_window
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif col_idx == 14:
                cell.value = total_prod
                cell.alignment = align_right
                cell.number_format = "#,##0"
            elif col_idx > len(fixed_headers):
                w_lbl = week_columns[col_idx - len(fixed_headers) - 1]["label"]
                cell.value = total_by_col.get(w_lbl, 0)
                cell.alignment = align_right
                cell.number_format = "#,##0"

        # Adjust column widths
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row in [1, 2]:
                    continue
                v_str = str(cell.value or "")
                if len(v_str) > max_len:
                    max_len = len(v_str)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 8)

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        return send_file(
            bio,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"{filename_base}.xlsx",
        )


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

    requested_year_str = request.args.get("ano", request.args.get("year", "2026")).strip()
    try:
        requested_year = int(requested_year_str) if requested_year_str else 2026
    except ValueError:
        requested_year = 2026

    snapshot: dict[str, Any] = get_snapshot(force_refresh=force_refresh, year=requested_year)
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
        v_norm = normalize_text(variety)
        b_norm = normalize_text(block)

        if "VERONICA" in product_master.upper():
            product_master = "VERONICA"
            pm_norm = "VERONICA"
            existing = db.query(RowAdjustmentDB).filter(
                RowAdjustmentDB.activity == activity,
                RowAdjustmentDB.product_master_norm.in_(["VERONICA", "VERONICA SPRAY"]),
                RowAdjustmentDB.variety_norm == v_norm,
                RowAdjustmentDB.block_norm == b_norm,
                RowAdjustmentDB.source_week == source_week,
            ).first()
        else:
            pm_norm = normalize_text(product_master)
            existing = db.query(RowAdjustmentDB).filter_by(
                activity=activity,
                product_master_norm=pm_norm,
                variety_norm=v_norm,
                block_norm=b_norm,
                source_week=source_week,
            ).first()

        if existing:
            existing.product_master_norm = pm_norm
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
    dump_stems_raw = payload.get("dump_stems")
    dump_stems = to_int_or_none(dump_stems_raw)
    mode = normalize_text(payload.get("mode"))
    is_dump = payload.get("is_dump")
    is_reset = bool(payload.get("reset")) or (raw_value in (None, "") and is_dump is None and dump_stems is None)
    reset_all = bool(payload.get("reset_all"))

    if activity not in {"SIEMBRA", "PODA"}:
        return jsonify({"ok": False, "message": "Actividad invalida."}), 400
    if mode not in {"AGRONOMO", "REAL"}:
        return jsonify({"ok": False, "message": "Modo invalido."}), 400
    if not product_master or not variety or not block or source_week is None or harvest_week is None:
        return jsonify({"ok": False, "message": "Faltan llaves."}), 400

    db = SessionLocal()
    try:
        v_norm = normalize_text(variety)
        b_norm = normalize_text(block)

        if "VERONICA" in product_master.upper():
            product_master = "VERONICA"
            pm_norm = "VERONICA"
            existing = db.query(WeekAdjustmentDB).filter(
                WeekAdjustmentDB.activity == activity,
                WeekAdjustmentDB.product_master_norm.in_(["VERONICA", "VERONICA SPRAY"]),
                WeekAdjustmentDB.variety_norm == v_norm,
                WeekAdjustmentDB.block_norm == b_norm,
                WeekAdjustmentDB.source_week == source_week,
                WeekAdjustmentDB.harvest_week == harvest_week,
            ).first()
        else:
            pm_norm = normalize_text(product_master)
            existing = db.query(WeekAdjustmentDB).filter_by(
                activity=activity,
                product_master_norm=pm_norm,
                variety_norm=v_norm,
                block_norm=b_norm,
                source_week=source_week,
                harvest_week=harvest_week,
            ).first()

        prev_val = None
        if existing:
            prev_val = existing.agronomo_estimate if mode == "AGRONOMO" else existing.real_closed
            if reset_all:
                existing.agronomo_estimate = None
                existing.real_closed = None
                existing.dump_stems = 0
                existing.is_dump = False
            elif mode == "AGRONOMO":
                existing.agronomo_estimate = value
                if dump_stems is not None:
                    existing.dump_stems = dump_stems
                    existing.is_dump = (dump_stems > 0) or bool(is_dump)
                elif is_reset:
                    existing.is_dump = False
                    existing.dump_stems = 0
            else:
                existing.real_closed = value
                if dump_stems is not None:
                    existing.dump_stems = dump_stems
                    existing.is_dump = (dump_stems > 0) or bool(is_dump)
                elif is_dump is not None:
                    existing.is_dump = bool(is_dump)
                    if is_dump and (not existing.dump_stems or existing.dump_stems == 0) and value:
                        existing.dump_stems = value
                elif is_reset:
                    existing.is_dump = False
                    existing.dump_stems = 0
            
            existing.updated_at = datetime.utcnow()

            if existing.agronomo_estimate is None and existing.real_closed is None and not existing.is_dump and (not existing.dump_stems or existing.dump_stems == 0):
                db.delete(existing)
        else:
            if not is_reset and not reset_all:
                if value is not None or (is_dump is not None and bool(is_dump)) or (dump_stems is not None and dump_stems > 0):
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
                        dump_stems=dump_stems if dump_stems is not None else (value if is_dump else 0),
                        is_dump=bool(is_dump or (dump_stems and dump_stems > 0)),
                    )
                    db.add(new_adj)

        # Record action in EstimateAuditDB
        action_type = "RESET" if (reset_all or is_reset) else "EDIT"
        new_val = None if (reset_all or is_reset) else value
        if prev_val != new_val or is_dump is not None or dump_stems is not None:
            audit_entry = EstimateAuditDB(
                product_master=product_master,
                product_master_norm=pm_norm,
                variety=variety,
                variety_norm=v_norm,
                block=block,
                block_norm=b_norm,
                activity=activity,
                source_week=source_week,
                harvest_week=harvest_week,
                mode=mode,
                action=action_type,
                previous_value=prev_val,
                new_value=new_val,
                dump_stems=dump_stems if dump_stems is not None else 0,
                notes=f"Guardado {mode} ({'Reseteo' if is_reset else 'Edición'})",
                created_at=datetime.utcnow(),
            )
            db.add(audit_entry)

        db.commit()
        invalidate_snapshot_cache()

        weekly_status = "MODELO"
        is_dump_val = False
        res_dump_stems = 0
        check = db.query(WeekAdjustmentDB).filter_by(
            activity=activity,
            product_master_norm=pm_norm,
            variety_norm=v_norm,
            block_norm=b_norm,
            source_week=source_week,
            harvest_week=harvest_week,
        ).first()
        if check:
            if check.is_dump or (check.dump_stems and check.dump_stems > 0):
                weekly_status = "DUMP"
                is_dump_val = True
            elif check.real_closed is not None and check.real_closed > 0:
                weekly_status = "REAL"
            elif check.agronomo_estimate is not None and check.agronomo_estimate > 0:
                weekly_status = "AGRONOMO"
            else:
                weekly_status = "MODELO"
            res_dump_stems = check.dump_stems or 0

        return jsonify({
            "ok": True,
            "weekly_status": weekly_status,
            "is_dump": is_dump_val,
            "dump_stems": res_dump_stems,
            "value": check.real_closed if (check and check.real_closed is not None) else (check.agronomo_estimate if check else None),
        })
    except SQLAlchemyError as exc:
        db.rollback()
        return jsonify({"ok": False, "message": str(exc)}), 500
    finally:
        db.close()


@app.get("/api/historial-estimaciones")
def get_estimate_history_api():
    block_filter = request.args.get("block", "").strip()
    source_week_filter = parse_week_code(request.args.get("source_week"))
    harvest_week_filter = parse_week_code(request.args.get("harvest_week"))
    product_filter = request.args.get("product", "").strip()
    variety_filter = request.args.get("variety", "").strip()
    mode_filter = request.args.get("mode", "").strip().upper()
    limit = max(1, min(500, int(request.args.get("limit") or 100)))

    db = SessionLocal()
    try:
        query = db.query(EstimateAuditDB)
        if block_filter:
            query = query.filter(EstimateAuditDB.block_norm == normalize_text(block_filter))
        if source_week_filter:
            query = query.filter(EstimateAuditDB.source_week == source_week_filter)
        if harvest_week_filter:
            query = query.filter(EstimateAuditDB.harvest_week == harvest_week_filter)
        if product_filter and product_filter.upper() != "ALL":
            if "VERONICA" in product_filter.upper():
                query = query.filter(EstimateAuditDB.product_master_norm.like("%VERONICA%"))
            else:
                query = query.filter(EstimateAuditDB.product_master_norm == normalize_text(product_filter))
        if variety_filter and variety_filter.upper() != "ALL":
            query = query.filter(EstimateAuditDB.variety_norm == normalize_text(variety_filter))
        if mode_filter and mode_filter in ("AGRONOMO", "REAL"):
            query = query.filter(EstimateAuditDB.mode == mode_filter)

        audits = query.order_by(EstimateAuditDB.id.desc()).limit(limit).all()

        results = [
            {
                "id": a.id,
                "product_master": a.product_master,
                "variety": a.variety,
                "variety_display": format_variety_display(a.variety),
                "block": a.block,
                "activity": a.activity,
                "source_week": a.source_week,
                "source_week_short": format_short_week(a.source_week),
                "harvest_week": a.harvest_week,
                "harvest_week_short": format_short_week(a.harvest_week),
                "mode": a.mode,
                "action": a.action,
                "previous_value": a.previous_value,
                "new_value": a.new_value,
                "dump_stems": a.dump_stems,
                "notes": a.notes or "",
                "created_at": a.created_at.strftime("%Y-%m-%d %H:%M:%S") if a.created_at else "",
            }
            for a in audits
        ]
        return jsonify({"ok": True, "count": len(results), "history": results})
    except SQLAlchemyError as exc:
        return jsonify({"ok": False, "message": str(exc)}), 500
    finally:
        db.close()


@app.post("/api/restaurar-estimacion")
def restore_estimate_api():
    payload = request.get_json(silent=True) or {}
    audit_id = to_int_or_none(payload.get("audit_id"))
    
    activity = normalize_text(payload.get("activity"))
    product_master = str(payload.get("product_master") or "").strip()
    variety = str(payload.get("variety") or "").strip()
    block = str(payload.get("block") or "").strip()
    source_week = parse_week_code(payload.get("source_week"))
    harvest_week = parse_week_code(payload.get("harvest_week"))
    value = to_int_or_none(payload.get("value"))
    mode = normalize_text(payload.get("mode") or "AGRONOMO")
    if mode not in ("AGRONOMO", "REAL"):
        mode = "AGRONOMO"

    db = SessionLocal()
    try:
        if audit_id:
            audit = db.query(EstimateAuditDB).filter_by(id=audit_id).first()
            if not audit:
                return jsonify({"ok": False, "message": "Registro de auditoría no encontrado."}), 404
            activity = audit.activity
            product_master = audit.product_master
            variety = audit.variety
            block = audit.block
            source_week = audit.source_week
            harvest_week = audit.harvest_week
            if value is None:
                value = audit.previous_value if audit.previous_value is not None else audit.new_value
            mode = audit.mode

        if not activity or not product_master or not variety or not block or source_week is None or harvest_week is None:
            return jsonify({"ok": False, "message": "Faltan parámetros requeridos para restaurar."}), 400

        v_norm = normalize_text(variety)
        b_norm = normalize_text(block)

        if "VERONICA" in product_master.upper():
            product_master = "VERONICA"
            pm_norm = "VERONICA"
            existing = db.query(WeekAdjustmentDB).filter(
                WeekAdjustmentDB.activity == activity,
                WeekAdjustmentDB.product_master_norm.in_(["VERONICA", "VERONICA SPRAY"]),
                WeekAdjustmentDB.variety_norm == v_norm,
                WeekAdjustmentDB.block_norm == b_norm,
                WeekAdjustmentDB.source_week == source_week,
                WeekAdjustmentDB.harvest_week == harvest_week,
            ).first()
        else:
            pm_norm = normalize_text(product_master)
            existing = db.query(WeekAdjustmentDB).filter_by(
                activity=activity,
                product_master_norm=pm_norm,
                variety_norm=v_norm,
                block_norm=b_norm,
                source_week=source_week,
                harvest_week=harvest_week,
            ).first()

        prev_val = None
        if existing:
            prev_val = existing.agronomo_estimate if mode == "AGRONOMO" else existing.real_closed
            if mode == "AGRONOMO":
                existing.agronomo_estimate = value
            else:
                existing.real_closed = value
            existing.updated_at = datetime.utcnow()
        else:
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
                dump_stems=0,
                is_dump=False,
            )
            db.add(new_adj)

        # Log restoration in audit
        restore_audit = EstimateAuditDB(
            product_master=product_master,
            product_master_norm=pm_norm,
            variety=variety,
            variety_norm=v_norm,
            block=block,
            block_norm=b_norm,
            activity=activity,
            source_week=source_week,
            harvest_week=harvest_week,
            mode=mode,
            action="RESTORE",
            previous_value=prev_val,
            new_value=value,
            dump_stems=0,
            notes=f"Restaurado a {value} tallos",
            created_at=datetime.utcnow(),
        )
        db.add(restore_audit)

        db.commit()
        invalidate_snapshot_cache()

        return jsonify({
            "ok": True,
            "message": f"Valor {value} tallos restaurado con éxito para bloque {block} en semana {format_short_week(harvest_week)}.",
            "value": value,
            "weekly_status": mode,
            "block": block,
            "source_week": source_week,
            "harvest_week": harvest_week,
            "activity": activity,
            "variety": variety,
            "product_master": product_master,
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
        v_norm = normalize_text(variety)
        b_norm = normalize_text(block)

        if "VERONICA" in product_master.upper():
            product_master = "VERONICA"
            pm_norm = "VERONICA"
            existing = db.query(BlockClosureDB).filter(
                BlockClosureDB.activity == activity,
                BlockClosureDB.product_master_norm.in_(["VERONICA", "VERONICA SPRAY"]),
                BlockClosureDB.variety_norm == v_norm,
                BlockClosureDB.block_norm == b_norm,
                BlockClosureDB.source_week == source_week,
            ).first()
        else:
            pm_norm = normalize_text(product_master)
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
        v_norm = normalize_optional_text(variety)

        if "VERONICA" in product_master.upper():
            product_master = "VERONICA"
            pm_norm = "VERONICA"
            existing = db.query(MassAdjustmentDB).filter(
                MassAdjustmentDB.product_master_norm.in_(["VERONICA", "VERONICA SPRAY"]),
                MassAdjustmentDB.variety_norm == v_norm,
                MassAdjustmentDB.activity == activity,
            ).first()
        else:
            pm_norm = normalize_optional_text(product_master)
            existing = db.query(MassAdjustmentDB).filter_by(
                product_master_norm=pm_norm,
                variety_norm=v_norm,
                activity=activity,
            ).first()

        if existing:
            existing.product_master_norm = pm_norm
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
            pm_upper = pm.upper()
            if "GYP" in pm_upper:
                pm = "GYPSOPHILA"
            elif "VERON" in pm_upper:
                pm = "VERONICA"
            elif "HYPER" in pm_upper:
                pm = "HYPERICUM"
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

    pm_upper = pm.upper()
    if "GYP" in pm_upper:
        pm = "GYPSOPHILA"
    elif "VERON" in pm_upper:
        pm = "VERONICA"
    elif "HYPER" in pm_upper:
        pm = "HYPERICUM"

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
    seed_data_from_excel_if_empty(WORKBOOK_PATH)
    debug_mode = os.environ.get("FLASK_DEBUG", "0").lower() in ("1", "true")
    app.run(host="0.0.0.0", port=5000, debug=debug_mode)
