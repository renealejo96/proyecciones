import os
import re
import time
import gzip
import json
import unicodedata
from datetime import datetime, date
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
from openpyxl import load_workbook
from sqlalchemy import (
    Boolean, Column, DateTime, Float, Index, Integer, String, Text, create_engine, text
)
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.orm import declarative_base, sessionmaker

DATABASE_URL = os.environ.get(
    "DATABASE_URL", "postgresql://postgres:admin@localhost:5432/proyecciones"
)

# Connect to database with robust connection pool configuration
engine = create_engine(
    DATABASE_URL,
    pool_size=10,
    max_overflow=20,
    pool_recycle=1800,
    pool_pre_ping=True,
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


class TpsrRecord(Base):
    __tablename__ = "tpsr_records"

    id = Column(Integer, primary_key=True, autoincrement=True)
    year = Column(Integer, nullable=True)
    source_week = Column(Integer, nullable=False, index=True)
    source_week_short = Column(String(10), nullable=True)
    date_val = Column(String(50), nullable=True)
    pruning_number = Column(String(50), default="")
    activity = Column(String(50), nullable=False, index=True)  # 'SIEMBRA' or 'PODA'
    product_master = Column(String(100), nullable=False)
    product_master_norm = Column(String(100), nullable=False, index=True)
    product = Column(String(100), nullable=True)
    variety = Column(String(100), nullable=False)
    variety_norm = Column(String(100), nullable=False, index=True)
    color = Column(String(50), nullable=True)
    material_origin = Column(String(100), nullable=True)
    block = Column(String(50), nullable=False, index=True)
    block_norm = Column(String(50), nullable=False, index=True)
    bed_location = Column(String(100), default="")
    total_beds = Column(Float, nullable=True)
    plants = Column(Integer, nullable=False, default=0)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class CycleDefinitionDB(Base):
    __tablename__ = "cycle_definitions"

    id = Column(Integer, primary_key=True, autoincrement=True)
    product_master = Column(String(100), nullable=False)
    product_master_norm = Column(String(100), nullable=False, index=True)
    variety = Column(String(100), nullable=False)
    variety_norm = Column(String(100), nullable=False, index=True)
    activity = Column(String(50), nullable=False, index=True)
    cycle_weeks = Column(Integer, nullable=False, default=0)
    waste_rate = Column(Float, nullable=False, default=0.0)
    stems_per_plant = Column(Float, nullable=False, default=0.0)
    curve = Column(Text, nullable=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class RowAdjustmentDB(Base):
    __tablename__ = "row_adjustments"

    id = Column(Integer, primary_key=True, autoincrement=True)
    activity = Column(String(50), nullable=False)
    product_master_norm = Column(String(100), nullable=False)
    variety_norm = Column(String(100), nullable=False)
    block = Column(String(50), nullable=False)
    block_norm = Column(String(50), nullable=False)
    source_week = Column(Integer, nullable=False)
    cycle_weeks = Column(Integer, nullable=True)
    waste_rate = Column(Float, nullable=True)
    stems_per_plant = Column(Float, nullable=True)
    updated_by = Column(String(120), nullable=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    __table_args__ = (
        Index("ix_row_adjustments_lookup", "activity", "product_master_norm", "variety_norm", "block_norm", "source_week"),
    )


class MassAdjustmentDB(Base):
    __tablename__ = "mass_adjustments"

    id = Column(Integer, primary_key=True, autoincrement=True)
    product_master_norm = Column(String(100), nullable=False)
    variety_norm = Column(String(100), nullable=False)
    activity = Column(String(50), nullable=False)
    cycle_weeks = Column(Integer, nullable=True)
    waste_rate = Column(Float, nullable=True)
    stems_per_plant = Column(Float, nullable=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class WeekAdjustmentDB(Base):
    __tablename__ = "week_adjustments"

    id = Column(Integer, primary_key=True, autoincrement=True)
    activity = Column(String(50), nullable=False)
    product_master_norm = Column(String(100), nullable=False)
    variety_norm = Column(String(100), nullable=False)
    block = Column(String(50), nullable=False)
    block_norm = Column(String(50), nullable=False)
    source_week = Column(Integer, nullable=False)
    harvest_week = Column(Integer, nullable=False)
    agronomo_estimate = Column(Integer, nullable=True)
    real_closed = Column(Integer, nullable=True)
    dump_stems = Column(Integer, default=0, nullable=True)
    is_dump = Column(Boolean, default=False, nullable=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    __table_args__ = (
        Index("ix_week_adjustments_lookup", "activity", "product_master_norm", "variety_norm", "block_norm", "source_week", "harvest_week"),
    )


class BlockClosureDB(Base):
    __tablename__ = "block_closures"

    id = Column(Integer, primary_key=True, autoincrement=True)
    activity = Column(String(50), nullable=False)
    product_master_norm = Column(String(100), nullable=False)
    variety_norm = Column(String(100), nullable=False)
    block = Column(String(50), nullable=False)
    block_norm = Column(String(50), nullable=False)
    source_week = Column(Integer, nullable=False)
    is_closed = Column(Boolean, default=False)
    closed_by = Column(String(120), nullable=True)
    closed_at = Column(DateTime, nullable=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    __table_args__ = (
        Index("ix_block_closures_lookup", "activity", "product_master_norm", "variety_norm", "block_norm", "source_week"),
    )


class RowDeactivationDB(Base):
    __tablename__ = "row_deactivations"

    id = Column(Integer, primary_key=True, autoincrement=True)
    activity = Column(String(50), nullable=False)
    product_master_norm = Column(String(100), nullable=False)
    variety_norm = Column(String(100), nullable=False)
    block = Column(String(50), nullable=False)
    block_norm = Column(String(50), nullable=False)
    source_week = Column(Integer, nullable=False)
    is_deactivated = Column(Boolean, default=True, nullable=False)
    reason = Column(String(255), nullable=True)
    updated_by = Column(String(120), nullable=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    __table_args__ = (
        Index("ix_row_deactivations_lookup", "activity", "product_master_norm", "variety_norm", "block_norm", "source_week"),
    )


class EstimateAuditDB(Base):
    __tablename__ = "estimate_audits"

    id = Column(Integer, primary_key=True, autoincrement=True)
    product_master = Column(String(100), nullable=False)
    product_master_norm = Column(String(100), nullable=False, index=True)
    variety = Column(String(100), nullable=False)
    variety_norm = Column(String(100), nullable=False, index=True)
    block = Column(String(50), nullable=False, index=True)
    block_norm = Column(String(50), nullable=False, index=True)
    activity = Column(String(50), nullable=False, index=True)
    source_week = Column(Integer, nullable=False, index=True)
    harvest_week = Column(Integer, nullable=False, index=True)
    mode = Column(String(50), default="AGRONOMO")  # 'AGRONOMO' or 'REAL'
    action = Column(String(50), default="EDIT")    # 'EDIT', 'RESTORE', 'RESET', 'CLOSE_BLOCK', 'REOPEN_BLOCK', 'FACTORY_RESET', 'ROW_PARAM', 'DEACTIVATE_ROW', 'REACTIVATE_ROW'
    previous_value = Column(Integer, nullable=True)
    new_value = Column(Integer, nullable=True)
    dump_stems = Column(Integer, nullable=True, default=0)
    username = Column(String(80), nullable=True, index=True)
    user_full_name = Column(String(120), nullable=True)
    user_role = Column(String(50), nullable=True)
    ip_address = Column(String(64), nullable=True)
    notes = Column(String(255), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, index=True)


class AppMetaDB(Base):
    __tablename__ = "app_meta"

    key = Column(String(50), primary_key=True)
    value = Column(String(255), nullable=False)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class UserDB(Base):
    __tablename__ = "app_users"

    id = Column(Integer, primary_key=True, autoincrement=True)
    username = Column(String(50), unique=True, nullable=False, index=True)
    password_hash = Column(String(255), nullable=False)
    role = Column(String(50), nullable=False, default="visita")  # 'agronomo', 'visita', 'admin'
    full_name = Column(String(100), nullable=True)
    allowed_views = Column(String(255), nullable=True)  # Comma-separated: "agro,reales,estadistica,tpsr,parametros"
    created_at = Column(DateTime, default=datetime.utcnow)


def get_data_version(db_session=None) -> int:
    close_session = False
    if db_session is None:
        db_session = SessionLocal()
        close_session = True
    try:
        meta = db_session.query(AppMetaDB).filter_by(key="data_version").first()
        if meta and meta.value:
            return int(meta.value)
        return 0
    except Exception:
        return 0
    finally:
        if close_session:
            db_session.close()


def bump_data_version(db_session=None) -> int:
    close_session = False
    if db_session is None:
        db_session = SessionLocal()
        close_session = True
    try:
        try:
            res = db_session.execute(
                text(
                    "INSERT INTO app_meta (key, value, updated_at) "
                    "VALUES ('data_version', '1', CURRENT_TIMESTAMP) "
                    "ON CONFLICT (key) DO UPDATE SET "
                    "value = (COALESCE(NULLIF(app_meta.value, ''), '0')::bigint + 1)::text, "
                    "updated_at = CURRENT_TIMESTAMP "
                    "RETURNING value;"
                )
            ).scalar()
            db_session.commit()
            if res is not None:
                return int(res)
        except Exception:
            db_session.rollback()
        meta = db_session.query(AppMetaDB).filter_by(key="data_version").first()
        if not meta:
            meta = AppMetaDB(key="data_version", value="1")
            db_session.add(meta)
        else:
            new_val = int(meta.value or 0) + 1
            meta.value = str(new_val)
            meta.updated_at = datetime.utcnow()
        db_session.commit()
        return int(meta.value)
    except Exception:
        db_session.rollback()
        return int(time.time())
    finally:
        if close_session:
            db_session.close()


def normalize_text(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(character for character in text if not unicodedata.combining(character))
    text = text.upper().strip()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_block_name_db(raw_block: Any) -> str:
    if not raw_block:
        return ""
    b = str(raw_block).strip().upper()
    b = re.sub(r"^(?:BLOQUE|BL\.?)\s*", "", b, flags=re.I).strip()
    return b


def normalize_variety_name_db(raw_var: Any, pm: str = "") -> str:
    if not raw_var:
        return ""
    v = str(raw_var).strip().upper()
    if ("GYP" in str(pm).upper() or not pm) and ("XCELENCE" in v or "XELENCE" in v or v == "XLENCE"):
        return "XLENCE"
    return v


def canonical_norms(activity: Any, product_master: Any, variety: Any, block: Any) -> tuple[str, str, str, str]:
    act_raw = normalize_text(activity)
    act_norm = "PODA" if "PODA" in act_raw else ("SIEMBRA" if act_raw else "")
    pm_str = str(product_master or "").strip().upper()
    if "GYP" in pm_str:
        pm_norm = "GYPSOPHILA"
    elif "VERON" in pm_str or "SPLASH" in normalize_text(variety):
        pm_norm = "VERONICA"
    elif "HYPER" in pm_str:
        pm_norm = "HYPERICUM"
    else:
        pm_norm = normalize_text(pm_str)
    v_clean = normalize_variety_name_db(variety, pm_norm)
    v_norm = normalize_text(v_clean)
    b_clean = normalize_block_name_db(block)
    b_norm = normalize_text(b_clean)
    return act_norm, pm_norm, v_norm, b_norm


def _shift_week_db(week_code: int, delta_weeks: int) -> int:
    from datetime import timedelta
    year = week_code // 100
    week = max(1, min(53, week_code % 100))
    try:
        monday = date.fromisocalendar(year, week, 1)
    except ValueError:
        monday = date.fromisocalendar(year, 52, 1)
    shifted = monday + timedelta(weeks=delta_weeks)
    iso_year, iso_week, _ = shifted.isocalendar()
    return iso_year * 100 + iso_week


def parse_week_code(raw_value: Any) -> int | None:
    if raw_value is None:
        return None
    digits = re.sub(r"\D", "", str(raw_value).strip())
    if not digits:
        return None
    num = int(digits)
    # Case 1: 1 or 2 digits, e.g. 28 -> week 28 of 2026
    if len(digits) in (1, 2) or num <= 53:
        year = date.today().isocalendar()[0]
        week = num
    elif len(digits) == 4:
        # Check if entered as "2028" (20 + week 28)
        if 2001 <= num <= 2053:
            year = date.today().isocalendar()[0]
            week = num % 100
        else:
            year = 2000 + int(digits[:2])
            week = int(digits[2:])
    elif len(digits) == 6:
        if 200001 <= num <= 200053 or 202001 <= num <= 202053:
            year = date.today().isocalendar()[0]
            week = num % 100
        else:
            year = int(digits[:4])
            week = int(digits[4:])
    else:
        return None
    return year * 100 + week


def format_short_week(week_code: int | None) -> str:
    if week_code is None:
        return ""
    try:
        wc = int(week_code)
    except (ValueError, TypeError):
        return str(week_code)
    if wc < 100:
        return f"{date.today().isocalendar()[0] % 100:02d}{wc:02d}"
    year = wc // 100
    week = wc % 100
    if year in (2000, 2020):
        year = date.today().isocalendar()[0]
    return f"{year % 100:02d}{week:02d}"


def float_or_zero(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    try:
        return float(value)
    except ValueError:
        return 0.0


def restore_snapshot_if_needed():
    snapshot_path = Path(__file__).resolve().parent / "siembras_podas" / "full_db_snapshot.json.gz"
    if not snapshot_path.exists():
        return

    with SessionLocal() as db:
        try:
            closure_2025_count = db.query(BlockClosureDB).filter(BlockClosureDB.source_week < 202600).count()
            tpsr_2025_count = db.query(TpsrRecord).filter(TpsrRecord.source_week < 202600).count()
            if closure_2025_count > 1000 and tpsr_2025_count > 1000:
                return  # Históricos ya presentes, no hacer nada

            print("Restaurando histórico de años anteriores (2024-2025) sin tocar datos de 2026...")
            with gzip.open(snapshot_path, "rt", encoding="utf-8") as f:
                data = json.load(f)

            # 1. TPSR: Solo insertar históricos (< 202600) que no existan
            tpsr_rows = data.get("tpsr_records", [])
            existing_tpsr = set(
                (r.activity, r.product_master_norm, r.variety_norm, r.block_norm, r.source_week, r.bed_location)
                for r in db.query(
                    TpsrRecord.activity,
                    TpsrRecord.product_master_norm,
                    TpsrRecord.variety_norm,
                    TpsrRecord.block_norm,
                    TpsrRecord.source_week,
                    TpsrRecord.bed_location,
                ).filter(TpsrRecord.source_week < 202600).all()
            )
            tpsr_to_insert = []
            for r in tpsr_rows:
                if (r.get("source_week") or 0) < 202600:
                    k = (r.get("activity"), r.get("product_master_norm"), r.get("variety_norm"), r.get("block_norm"), r.get("source_week"), r.get("bed_location"))
                    if k not in existing_tpsr:
                        r_copy = dict(r)
                        r_copy.pop("id", None)
                        if isinstance(r_copy.get("created_at"), str):
                            try: r_copy["created_at"] = datetime.fromisoformat(r_copy["created_at"])
                            except Exception: pass
                        if isinstance(r_copy.get("updated_at"), str):
                            try: r_copy["updated_at"] = datetime.fromisoformat(r_copy["updated_at"])
                            except Exception: pass
                        tpsr_to_insert.append(r_copy)
                        existing_tpsr.add(k)
            if tpsr_to_insert:
                db.bulk_insert_mappings(TpsrRecord, tpsr_to_insert)
                db.commit()
                print(f"Inyectados {len(tpsr_to_insert)} registros históricos TPSR (2024-2025)")

            # 2. Block Closures: Solo cierres históricos (< 202600)
            bc_rows = data.get("block_closures", [])
            existing_bc = set(
                (r.activity, r.product_master_norm, r.variety_norm, r.block_norm, r.source_week)
                for r in db.query(
                    BlockClosureDB.activity,
                    BlockClosureDB.product_master_norm,
                    BlockClosureDB.variety_norm,
                    BlockClosureDB.block_norm,
                    BlockClosureDB.source_week,
                ).filter(BlockClosureDB.source_week < 202600).all()
            )
            bc_to_insert = []
            for r in bc_rows:
                if (r.get("source_week") or 0) < 202600:
                    k = (r.get("activity"), r.get("product_master_norm"), r.get("variety_norm"), r.get("block_norm"), r.get("source_week"))
                    if k not in existing_bc:
                        r_copy = dict(r)
                        r_copy.pop("id", None)
                        if isinstance(r_copy.get("closed_at"), str):
                            try: r_copy["closed_at"] = datetime.fromisoformat(r_copy["closed_at"])
                            except Exception: pass
                        bc_to_insert.append(r_copy)
                        existing_bc.add(k)
            if bc_to_insert:
                db.bulk_insert_mappings(BlockClosureDB, bc_to_insert)
                db.commit()
                print(f"Inyectados {len(bc_to_insert)} cierres históricos de bloques (2024-2025)")

            # 3. Week Adjustments: Solo cortes históricos previos a semana 202630
            wa_rows = data.get("week_adjustments", [])
            existing_wa = set(
                (r.activity, r.product_master_norm, r.variety_norm, r.block_norm, r.source_week, r.harvest_week)
                for r in db.query(
                    WeekAdjustmentDB.activity,
                    WeekAdjustmentDB.product_master_norm,
                    WeekAdjustmentDB.variety_norm,
                    WeekAdjustmentDB.block_norm,
                    WeekAdjustmentDB.source_week,
                    WeekAdjustmentDB.harvest_week,
                ).filter(WeekAdjustmentDB.harvest_week < 202630).all()
            )
            wa_to_insert = []
            for r in wa_rows:
                if (r.get("harvest_week") or 0) < 202630:
                    k = (r.get("activity"), r.get("product_master_norm"), r.get("variety_norm"), r.get("block_norm"), r.get("source_week"), r.get("harvest_week"))
                    if k not in existing_wa:
                        r_copy = dict(r)
                        r_copy.pop("id", None)
                        if isinstance(r_copy.get("created_at"), str):
                            try: r_copy["created_at"] = datetime.fromisoformat(r_copy["created_at"])
                            except Exception: pass
                        if isinstance(r_copy.get("updated_at"), str):
                            try: r_copy["updated_at"] = datetime.fromisoformat(r_copy["updated_at"])
                            except Exception: pass
                        wa_to_insert.append(r_copy)
                        existing_wa.add(k)
            if wa_to_insert:
                db.bulk_insert_mappings(WeekAdjustmentDB, wa_to_insert)
                db.commit()
                print(f"Inyectados {len(wa_to_insert)} ajustes/cortes históricos (pre-semana 30)")

            # 4. Cycle definitions: solo si faltan
            cycle_rows = data.get("cycle_definitions", [])
            existing_cycles = set(
                (r.product_master_norm, r.variety_norm, r.activity)
                for r in db.query(
                    CycleDefinitionDB.product_master_norm,
                    CycleDefinitionDB.variety_norm,
                    CycleDefinitionDB.activity,
                ).all()
            )
            cycle_to_insert = []
            for r in cycle_rows:
                k = (r.get("product_master_norm"), r.get("variety_norm"), r.get("activity"))
                if k not in existing_cycles:
                    r_copy = dict(r)
                    r_copy.pop("id", None)
                    cycle_to_insert.append(r_copy)
                    existing_cycles.add(k)
            if cycle_to_insert:
                db.bulk_insert_mappings(CycleDefinitionDB, cycle_to_insert)
                db.commit()

            print("Restauración histórica completada con éxito. Datos del agrónomo 2026 100% protegidos.")
        except Exception as e:
            db.rollback()
            print(f"Warning: could not restore historical snapshot: {e}")


def cleanup_outlier_hyp_2023_data():
    """
    Elimina datos atípicos erróneos de Hypericum de las semanas 202330 y 202331 (lote bloque 12 con ~20-30 T/pl).
    Es idempotente: si ya se eliminaron, no realiza ninguna acción.
    """
    with SessionLocal() as db:
        try:
            flag = db.query(AppMetaDB).filter_by(key="cleanup_hyp_2330_2331_done").first()
            if flag and flag.value == "1":
                return

            wa_del = db.query(WeekAdjustmentDB).filter(
                WeekAdjustmentDB.source_week.in_([202330, 202331]),
                WeekAdjustmentDB.product_master_norm == "HYPERICUM"
            ).delete(synchronize_session=False)

            bc_del = db.query(BlockClosureDB).filter(
                BlockClosureDB.source_week.in_([202330, 202331]),
                BlockClosureDB.product_master_norm == "HYPERICUM"
            ).delete(synchronize_session=False)

            tpsr_del = db.query(TpsrRecord).filter(
                TpsrRecord.source_week.in_([202330, 202331]),
                TpsrRecord.product_master_norm == "HYPERICUM"
            ).delete(synchronize_session=False)

            if not flag:
                db.add(AppMetaDB(key="cleanup_hyp_2330_2331_done", value="1"))
            else:
                flag.value = "1"

            bump_data_version(db)
            db.commit()
            if wa_del > 0 or bc_del > 0 or tpsr_del > 0:
                print(f"Limpieza de datos atípicos HYP 202330-202331 completada: {wa_del} ajustes, {bc_del} cierres, {tpsr_del} tpsr eliminados.")
        except Exception as e:
            db.rollback()
            print(f"Warning en cleanup_outlier_hyp_2023_data: {e}")


def migrate_homologate_and_pre30_closures():
    """
    1. Homologa nombres de variedad Gypsophila (XCELENCE, XELENCE -> XLENCE).
    2. Homologa nombres de bloques removiendo prefijos 'BL ' o 'Bl '.
    3. Marca como cerrados (is_closed=True) todos los lotes y semanas anteriores a la semana 30 (< 202630).
    4. Inserta registros de cierre faltantes en BlockClosureDB para todo lote histórico pre-30.
    Es completamente idempotente.
    """
    with SessionLocal() as db:
        try:
            flag = db.query(AppMetaDB).filter_by(key="migration_healed_pre30_and_xlence_v1").first()
            if flag and flag.value == "1":
                return

            print("Aplicando migración: Homologación de XLENCE, limpieza de bloques y cierre pre-semana 30...")

            # 1. Homologación XLENCE en Gypsophila
            db.execute(text("""
                UPDATE tpsr_records 
                SET variety = 'XLENCE', variety_norm = 'XLENCE' 
                WHERE (product_master_norm = 'GYPSOPHILA' OR UPPER(product_master) LIKE '%GYP%') 
                AND (variety_norm IN ('XCELENCE', 'XELENCE') OR UPPER(variety) LIKE '%XCELENCE%' OR UPPER(variety) LIKE '%XELENCE%');

                UPDATE week_adjustments 
                SET variety_norm = 'XLENCE' 
                WHERE (product_master_norm = 'GYPSOPHILA' OR UPPER(product_master_norm) LIKE '%GYP%') 
                AND (variety_norm IN ('XCELENCE', 'XELENCE') OR UPPER(variety_norm) LIKE '%XCELENCE%' OR UPPER(variety_norm) LIKE '%XELENCE%');

                UPDATE block_closures 
                SET variety_norm = 'XLENCE' 
                WHERE (product_master_norm = 'GYPSOPHILA' OR UPPER(product_master_norm) LIKE '%GYP%') 
                AND (variety_norm IN ('XCELENCE', 'XELENCE') OR UPPER(variety_norm) LIKE '%XCELENCE%' OR UPPER(variety_norm) LIKE '%XELENCE%');

                UPDATE cycle_definitions 
                SET variety = 'XLENCE', variety_norm = 'XLENCE' 
                WHERE (product_master_norm = 'GYPSOPHILA' OR UPPER(product_master_norm) LIKE '%GYP%') 
                AND (variety_norm IN ('XCELENCE', 'XELENCE') OR UPPER(variety_norm) LIKE '%XCELENCE%' OR UPPER(variety_norm) LIKE '%XELENCE%');

                UPDATE row_adjustments 
                SET variety_norm = 'XLENCE' 
                WHERE (product_master_norm = 'GYPSOPHILA' OR UPPER(product_master_norm) LIKE '%GYP%') 
                AND (variety_norm IN ('XCELENCE', 'XELENCE') OR UPPER(variety_norm) LIKE '%XCELENCE%' OR UPPER(variety_norm) LIKE '%XELENCE%');
            """))
            db.commit()

            # 2. Homologación nombres de bloques (remover 'BL ' o 'Bl ')
            db.execute(text("""
                UPDATE tpsr_records 
                SET block = UPPER(REGEXP_REPLACE(block, '^(BL|Bl|bl)\\s*', '', 'i')), 
                    block_norm = UPPER(REGEXP_REPLACE(block_norm, '^(BL|Bl|bl)\\s*', '', 'i')) 
                WHERE block ~* '^(BL|Bl)\\s*';

                UPDATE week_adjustments 
                SET block = UPPER(REGEXP_REPLACE(block, '^(BL|Bl|bl)\\s*', '', 'i')), 
                    block_norm = UPPER(REGEXP_REPLACE(block_norm, '^(BL|Bl|bl)\\s*', '', 'i')) 
                WHERE block ~* '^(BL|Bl)\\s*';

                UPDATE block_closures 
                SET block = UPPER(REGEXP_REPLACE(block, '^(BL|Bl|bl)\\s*', '', 'i')), 
                    block_norm = UPPER(REGEXP_REPLACE(block_norm, '^(BL|Bl|bl)\\s*', '', 'i')) 
                WHERE block ~* '^(BL|Bl)\\s*';

                UPDATE row_adjustments 
                SET block = UPPER(REGEXP_REPLACE(block, '^(BL|Bl|bl)\\s*', '', 'i')), 
                    block_norm = UPPER(REGEXP_REPLACE(block_norm, '^(BL|Bl|bl)\\s*', '', 'i')) 
                WHERE block ~* '^(BL|Bl)\\s*';
            """))
            db.commit()

            # 3. Marcar como cerrados todos los cierres existentes de semanas pre-30 (< 202630)
            db.execute(text("UPDATE block_closures SET is_closed = TRUE WHERE source_week < 202630;"))
            db.commit()

            # 4. Insertar cierres faltantes para todos los lotes de semanas anteriores a 202630
            existing_closures = set(
                (c.activity, c.product_master_norm, c.variety_norm, c.block_norm, c.source_week)
                for c in db.query(
                    BlockClosureDB.activity,
                    BlockClosureDB.product_master_norm,
                    BlockClosureDB.variety_norm,
                    BlockClosureDB.block_norm,
                    BlockClosureDB.source_week
                ).all()
            )

            tpsr_lots = db.query(
                TpsrRecord.activity,
                TpsrRecord.product_master,
                TpsrRecord.product_master_norm,
                TpsrRecord.variety,
                TpsrRecord.variety_norm,
                TpsrRecord.block,
                TpsrRecord.block_norm,
                TpsrRecord.source_week
            ).filter(TpsrRecord.source_week < 202630).all()

            closures_to_add = []
            seen_keys = set()
            for r in tpsr_lots:
                act = "PODA" if "PODA" in (r.activity or "").upper() else "SIEMBRA"
                pm = "VERONICA" if "VERONICA" in (r.product_master or "").upper() else (r.product_master or "")
                pm_norm = "VERONICA" if "VERONICA" in pm.upper() else (r.product_master_norm or "")
                var_norm = (r.variety_norm or "").strip().upper()
                if "GYP" in pm_norm and ("XCELENCE" in var_norm or "XELENCE" in var_norm):
                    var_norm = "XLENCE"
                blk_clean = re.sub(r'^(BL|Bl|bl)\s*', '', str(r.block_norm or r.block or '').strip(), flags=re.IGNORECASE).upper()
                sw = r.source_week
                if not sw:
                    continue
                k = (act, pm_norm, var_norm, blk_clean, sw)
                if k not in existing_closures and k not in seen_keys:
                    seen_keys.add(k)
                    closures_to_add.append({
                        "activity": act,
                        "product_master_norm": pm_norm,
                        "variety_norm": var_norm,
                        "block": blk_clean,
                        "block_norm": blk_clean,
                        "source_week": sw,
                        "is_closed": True
                    })

            if closures_to_add:
                db.bulk_insert_mappings(BlockClosureDB, closures_to_add)
                db.commit()
                print(f"Insertados {len(closures_to_add)} cierres de bloques para semanas anteriores a 202630.")

            if not flag:
                db.add(AppMetaDB(key="migration_healed_pre30_and_xlence_v1", value="1"))
            else:
                flag.value = "1"

            bump_data_version(db)
            db.commit()
            print("Migración de homologación y cierres pre-semana 30 completada con éxito.")
        except Exception as e:
            db.rollback()
            print(f"Warning en migrate_homologate_and_pre30_closures: {e}")


def ensure_audit_schema_columns():
    with engine.connect() as conn:
        for alter_sql in [
            "ALTER TABLE estimate_audits ADD COLUMN IF NOT EXISTS username VARCHAR(80);",
            "ALTER TABLE estimate_audits ADD COLUMN IF NOT EXISTS user_full_name VARCHAR(120);",
            "ALTER TABLE estimate_audits ADD COLUMN IF NOT EXISTS user_role VARCHAR(50);",
            "ALTER TABLE estimate_audits ADD COLUMN IF NOT EXISTS ip_address VARCHAR(64);",
            "CREATE INDEX IF NOT EXISTS ix_estimate_audits_username ON estimate_audits (username);",
            "ALTER TABLE block_closures ADD COLUMN IF NOT EXISTS closed_by VARCHAR(120);",
            "ALTER TABLE block_closures ADD COLUMN IF NOT EXISTS closed_at TIMESTAMP;",
            "ALTER TABLE row_deactivations ADD COLUMN IF NOT EXISTS updated_by VARCHAR(120);",
            "ALTER TABLE row_adjustments ADD COLUMN IF NOT EXISTS updated_by VARCHAR(120);",
        ]:
            try:
                conn.execute(text(alter_sql))
                conn.commit()
            except Exception:
                try:
                    conn.rollback()
                except Exception:
                    pass


def ensure_audit_and_heal_closures_v2():
    """
    1. Garantiza que las columnas de auditoría de usuario existan en todas las tablas operativas.
    2. Deduplica block_closures usando claves canónicas (respetando reaperturas manuales).
    3. Reabre automáticamente los lotes (especialmente Hypericum de ciclo 19-21 semanas)
       cuya semana de origen era < 202630 pero cuya cosecha real ocurre en semanas >= 202642,
       que habían sido cerrados erróneamente por el corte pre-semana 30 basado en source_week.
    """
    ensure_audit_schema_columns()

    with SessionLocal() as db:
        try:
            flag = db.query(AppMetaDB).filter_by(key="migration_heal_false_pre30_closures_v2").first()
            if flag and flag.value == "1":
                return

            # Build cycle lookup for accurate harvest_end_week computation
            cycles = db.query(CycleDefinitionDB).all()
            cycle_lookup: dict[tuple[str, str, str], tuple[int, int]] = {}
            cycle_by_var: dict[tuple[str, str], tuple[int, int]] = {}
            for c in cycles:
                _, pm_c, v_c, _ = canonical_norms(c.activity, c.product_master, c.variety, "")
                act_c = "PODA" if "PODA" in normalize_text(c.activity) else "SIEMBRA"
                curve_len = max(1, len([x for x in (c.curve or "1").split(",") if x.strip()]))
                cw = int(c.cycle_weeks or 0)
                cycle_lookup[(pm_c, v_c, act_c)] = (cw, curve_len)
                cycle_by_var[(v_c, act_c)] = (cw, curve_len)

            all_closures = db.query(BlockClosureDB).order_by(BlockClosureDB.id.asc()).all()
            grouped_closures: dict[tuple[str, str, str, str, int], list[BlockClosureDB]] = {}
            for bc in all_closures:
                act_n, pm_n, v_n, b_n = canonical_norms(bc.activity, bc.product_master_norm, bc.variety_norm, bc.block_norm or bc.block)
                sw = int(bc.source_week or 0)
                grouped_closures.setdefault((act_n, pm_n, v_n, b_n, sw), []).append(bc)

            reopened_count = 0
            dedup_deleted = 0

            for (act_n, pm_n, v_n, b_n, sw), rows in grouped_closures.items():
                # If multiple duplicate rows exist and ANY row was explicitly reopened (is_closed=False), honor False!
                any_reopened = any(not bool(r.is_closed) for r in rows)
                keeper = rows[-1]
                for extra in rows[:-1]:
                    db.delete(extra)
                    dedup_deleted += 1

                keeper.activity = act_n
                keeper.product_master_norm = pm_n
                keeper.variety_norm = v_n
                keeper.block = normalize_block_name_db(keeper.block or b_n) or b_n
                keeper.block_norm = b_n
                if any_reopened:
                    keeper.is_closed = False

                # Check if this 2026 pre-30 lot actually harvests in active weeks >= 202642
                if 202601 <= sw < 202630 and keeper.is_closed:
                    c_info = cycle_lookup.get((pm_n, v_n, act_n)) or cycle_by_var.get((v_n, act_n))
                    if c_info:
                        cw, c_len = c_info
                    else:
                        cw = 20 if pm_n == "HYPERICUM" else 13
                        c_len = 2
                    harvest_end_week = _shift_week_db(sw, cw + max(0, c_len - 1))
                    if harvest_end_week >= 202642:
                        keeper.is_closed = False
                        keeper.closed_by = "sistema (reapertura auto ciclo activo)"
                        keeper.updated_at = datetime.utcnow()
                        reopened_count += 1

            if not flag:
                db.add(AppMetaDB(key="migration_heal_false_pre30_closures_v2", value="1"))
            else:
                flag.value = "1"

            db.commit()
            bump_data_version(db)
            print(f"Sanación v2 completada: {reopened_count} bloques activos reabiertos (cosecha >= 2642), {dedup_deleted} cierres duplicados limpiados.")
        except Exception as exc:
            db.rollback()
            print(f"Warning en ensure_audit_and_heal_closures_v2: {exc}")


def init_db():
    Base.metadata.create_all(bind=engine)
    ensure_audit_schema_columns()
    restore_snapshot_if_needed()
    cleanup_outlier_hyp_2023_data()
    migrate_homologate_and_pre30_closures()
    ensure_audit_and_heal_closures_v2()

    # 1. Fast check: if migrations and data healing have already been applied, return immediately!
    db = SessionLocal()
    try:
        done_flag = db.query(AppMetaDB).filter_by(key="db_migration_healed_v1").first()
        if done_flag and done_flag.value == "1":
            return
    except Exception:
        pass
    finally:
        db.close()

    # 2. Concurrency guard: Use PostgreSQL advisory lock so only ONE worker can execute migrations
    with engine.connect() as conn:
        is_postgres = False
        try:
            lock_res = conn.execute(text("SELECT pg_try_advisory_lock(982341);")).scalar()
            is_postgres = True
            if not lock_res:
                # Another worker is already executing migrations, do not run concurrently!
                return
        except Exception:
            pass

        try:
            # Re-check under lock in case another worker just finished
            with SessionLocal() as db_chk:
                done_flag = db_chk.query(AppMetaDB).filter_by(key="db_migration_healed_v1").first()
                if done_flag and done_flag.value == "1":
                    return

            # 1. Schema migrations
            for alter_sql in [
                "ALTER TABLE week_adjustments ADD COLUMN IF NOT EXISTS is_dump BOOLEAN DEFAULT FALSE;",
                "ALTER TABLE week_adjustments ADD COLUMN IF NOT EXISTS dump_stems INTEGER DEFAULT 0;",
                "ALTER TABLE app_users ADD COLUMN IF NOT EXISTS allowed_views VARCHAR(255);",
                "CREATE TABLE IF NOT EXISTS row_deactivations (id SERIAL PRIMARY KEY, activity VARCHAR(50) NOT NULL, product_master_norm VARCHAR(100) NOT NULL, variety_norm VARCHAR(100) NOT NULL, block VARCHAR(50) NOT NULL, block_norm VARCHAR(50) NOT NULL, source_week INTEGER NOT NULL, is_deactivated BOOLEAN DEFAULT TRUE, reason VARCHAR(255), updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);",
                "CREATE INDEX IF NOT EXISTS ix_row_deactivations_lookup ON row_deactivations (activity, product_master_norm, variety_norm, block_norm, source_week);",
            ]:
                try:
                    conn.execute(text(alter_sql))
                    conn.commit()
                except Exception:
                    try:
                        conn.rollback()
                    except Exception:
                        pass

            # 2. Data healing and week canonicalization migrations (runs ONCE in history)
            for heal_sql in [
                # Fix 1 or 2 digit weeks (e.g. 28 -> 202628)
                "UPDATE tpsr_records SET source_week = 202600 + source_week WHERE source_week >= 1 AND source_week <= 53;",
                # Fix accidental year 2000 weeks (e.g. 200028 -> 202628)
                "UPDATE tpsr_records SET source_week = 202600 + (source_week % 100) WHERE source_week >= 200001 AND source_week <= 200053;",
                # Fix accidental year 2020 weeks (e.g. 202028 from 200000 + 2028 -> 202628)
                "UPDATE tpsr_records SET source_week = 202600 + (source_week % 100) WHERE source_week >= 202001 AND source_week <= 202053;",
                # Fix 4-digit weeks entered as 20XX (e.g. 2028 -> 202628)
                "UPDATE tpsr_records SET source_week = 202600 + (source_week % 100) WHERE source_week >= 2001 AND source_week <= 2053;",
                # Fix 4-digit YYWW weeks (e.g. 2628 -> 202628, 2428 -> 202428)
                "UPDATE tpsr_records SET source_week = 200000 + source_week WHERE source_week >= 2401 AND source_week <= 2953;",
                # Ensure source_week_short is always 4 digits
                "UPDATE tpsr_records SET source_week_short = LPAD(MOD(source_week, 10000)::text, 4, '0') WHERE source_week IS NOT NULL;",
                # Normalize product_master names
                "UPDATE tpsr_records SET product_master = 'GYPSOPHILA', product_master_norm = 'GYPSOPHILA' WHERE UPPER(product_master) LIKE '%GYP%';",
                "UPDATE tpsr_records SET product_master = 'VERONICA', product_master_norm = 'VERONICA' WHERE UPPER(product_master) LIKE '%VERON%';",
                "UPDATE tpsr_records SET product_master = 'HYPERICUM', product_master_norm = 'HYPERICUM' WHERE UPPER(product_master) LIKE '%HYPER%';",
                "UPDATE tpsr_records SET product_master_norm = UPPER(TRIM(product_master)) WHERE product_master_norm IS NULL OR product_master_norm = '';",
                "UPDATE tpsr_records SET variety_norm = UPPER(TRIM(variety)) WHERE variety_norm IS NULL OR variety_norm = '';",
                "UPDATE tpsr_records SET block_norm = UPPER(TRIM(block)) WHERE block_norm IS NULL OR block_norm = '';",
                "UPDATE tpsr_records SET activity = 'SIEMBRA' WHERE activity IS NULL OR UPPER(activity) LIKE '%SIEMBR%';",
                "UPDATE tpsr_records SET activity = 'PODA' WHERE UPPER(activity) LIKE '%PODA%';",
                # Heal row_adjustments
                "UPDATE row_adjustments SET source_week = 202600 + source_week WHERE source_week >= 1 AND source_week <= 53;",
                "UPDATE row_adjustments SET source_week = 202600 + (source_week % 100) WHERE source_week >= 200001 AND source_week <= 200053;",
                "UPDATE row_adjustments SET source_week = 202600 + (source_week % 100) WHERE source_week >= 202001 AND source_week <= 202053;",
                "UPDATE row_adjustments SET source_week = 202600 + (source_week % 100) WHERE source_week >= 2001 AND source_week <= 2053;",
                "UPDATE row_adjustments SET source_week = 200000 + source_week WHERE source_week >= 2401 AND source_week <= 2953;",
                # Heal week_adjustments
                "UPDATE week_adjustments SET source_week = 202600 + source_week WHERE source_week >= 1 AND source_week <= 53;",
                "UPDATE week_adjustments SET source_week = 202600 + (source_week % 100) WHERE source_week >= 200001 AND source_week <= 200053;",
                "UPDATE week_adjustments SET source_week = 202600 + (source_week % 100) WHERE source_week >= 202001 AND source_week <= 202053;",
                "UPDATE week_adjustments SET source_week = 202600 + (source_week % 100) WHERE source_week >= 2001 AND source_week <= 2053;",
                "UPDATE week_adjustments SET source_week = 200000 + source_week WHERE source_week >= 2401 AND source_week <= 2953;",
                "UPDATE week_adjustments SET harvest_week = 202600 + harvest_week WHERE harvest_week >= 1 AND harvest_week <= 53;",
                "UPDATE week_adjustments SET harvest_week = 202600 + (harvest_week % 100) WHERE harvest_week >= 200001 AND harvest_week <= 200053;",
                "UPDATE week_adjustments SET harvest_week = 202600 + (harvest_week % 100) WHERE harvest_week >= 202001 AND harvest_week <= 202053;",
                "UPDATE week_adjustments SET harvest_week = 202600 + (harvest_week % 100) WHERE harvest_week >= 2001 AND harvest_week <= 2053;",
                "UPDATE week_adjustments SET harvest_week = 200000 + harvest_week WHERE harvest_week >= 2401 AND harvest_week <= 2953;",
                # Heal block_closures
                "UPDATE block_closures SET source_week = 202600 + source_week WHERE source_week >= 1 AND source_week <= 53;",
                "UPDATE block_closures SET source_week = 202600 + (source_week % 100) WHERE source_week >= 200001 AND source_week <= 200053;",
                "UPDATE block_closures SET source_week = 202600 + (source_week % 100) WHERE source_week >= 202001 AND source_week <= 202053;",
                "UPDATE block_closures SET source_week = 202600 + (source_week % 100) WHERE source_week >= 2001 AND source_week <= 2053;",
                "UPDATE block_closures SET source_week = 200000 + source_week WHERE source_week >= 2401 AND source_week <= 2953;",
            ]:
                try:
                    conn.execute(text(heal_sql))
                    conn.commit()
                except Exception:
                    try:
                        conn.rollback()
                    except Exception:
                        pass

            db = SessionLocal()
            try:
                meta = db.query(AppMetaDB).filter_by(key="data_version").first()
                if not meta:
                    meta = AppMetaDB(key="data_version", value="1")
                    db.add(meta)

                # Mark migration as done permanently so it NEVER executes again
                done_meta = db.query(AppMetaDB).filter_by(key="db_migration_healed_v1").first()
                if not done_meta:
                    done_meta = AppMetaDB(key="db_migration_healed_v1", value="1")
                    db.add(done_meta)
                else:
                    done_meta.value = "1"

                # Seed baseline audit history from current adjustments if table is empty
                audit_count = db.query(EstimateAuditDB).count()
                if audit_count == 0:
                    existing_adjs = db.query(WeekAdjustmentDB).all()
                    for a in existing_adjs:
                        val = a.agronomo_estimate if a.agronomo_estimate is not None else a.real_closed
                        if val is not None or a.is_dump or a.dump_stems:
                            db.add(
                                EstimateAuditDB(
                                    product_master=a.product_master_norm,
                                    product_master_norm=a.product_master_norm,
                                    variety=a.variety_norm,
                                    variety_norm=a.variety_norm,
                                    block=a.block,
                                    block_norm=a.block_norm,
                                    activity=a.activity,
                                    source_week=a.source_week,
                                    harvest_week=a.harvest_week,
                                    mode="AGRONOMO" if a.agronomo_estimate is not None else "REAL",
                                    action="INITIAL",
                                    previous_value=None,
                                    new_value=val,
                                    dump_stems=a.dump_stems or 0,
                                    notes="Registro inicial",
                                    created_at=a.updated_at or datetime.utcnow(),
                                )
                            )
                seed_default_users_if_empty(db)
                db.commit()
            except Exception:
                db.rollback()
            finally:
                db.close()

        finally:
            if is_postgres:
                try:
                    conn.execute(text("SELECT pg_advisory_unlock(982341);"))
                    conn.commit()
                except Exception:
                    pass


def parse_allowed_views(allowed_views_raw: Any, role: str = "visita") -> list[str]:
    all_possible = ["agro", "reales", "estadistica", "tpsr", "parametros"]
    if role == "admin":
        return all_possible

    if allowed_views_raw is not None and str(allowed_views_raw).strip() != "":
        parsed = [v.strip().lower() for v in str(allowed_views_raw).split(",") if v.strip().lower() in all_possible]
        if parsed:
            return parsed

    # Default views by role
    if role == "agronomo":
        return ["agro"]
    elif role == "visita":
        return ["estadistica"]
    return ["estadistica"]


def seed_default_users_if_empty(db_session=None):
    from werkzeug.security import generate_password_hash

    close_session = False
    if db_session is None:
        db_session = SessionLocal()
        close_session = True
    try:
        user_count = db_session.query(UserDB).count()
        if user_count == 0:
            defaults = [
                UserDB(
                    username="agronomo",
                    password_hash=generate_password_hash(os.environ.get("AUTH_AGRONOMO_PASS", "agro2026")),
                    role="agronomo",
                    full_name="Agrónomo Operativo",
                    allowed_views="agro",
                ),
                UserDB(
                    username="visita",
                    password_hash=generate_password_hash(os.environ.get("AUTH_VISITA_PASS", "visita2026")),
                    role="visita",
                    full_name="Usuario Visita / Coordinación",
                    allowed_views="estadistica",
                ),
                UserDB(
                    username="gerente",
                    password_hash=generate_password_hash(os.environ.get("AUTH_GERENTE_PASS", "gerente2026")),
                    role="visita",
                    full_name="Gerencia General",
                    allowed_views="estadistica",
                ),
                UserDB(
                    username="coordinador",
                    password_hash=generate_password_hash(os.environ.get("AUTH_COORD_PASS", "coord2026")),
                    role="visita",
                    full_name="Coordinador de Producción",
                    allowed_views="estadistica",
                ),
                UserDB(
                    username="admin",
                    password_hash=generate_password_hash(os.environ.get("AUTH_ADMIN_PASS", "admin2026")),
                    role="admin",
                    full_name="Administrador General",
                    allowed_views="agro,reales,estadistica,tpsr,parametros",
                ),
            ]
            db_session.add_all(defaults)
            db_session.commit()
        else:
            # Migration check: Ensure existing records have default allowed_views if null
            existing_users = db_session.query(UserDB).all()
            for u in existing_users:
                if not u.allowed_views:
                    if u.role == "admin":
                        u.allowed_views = "agro,reales,estadistica,tpsr,parametros"
                    elif u.role == "agronomo":
                        u.allowed_views = "agro"
                    else:
                        u.allowed_views = "estadistica"
            db_session.commit()
    except Exception:
        db_session.rollback()
    finally:
        if close_session:
            db_session.close()


def verify_user(username: str, password: str) -> dict[str, Any] | None:
    from werkzeug.security import check_password_hash

    db = SessionLocal()
    try:
        u_clean = (username or "").strip().lower()
        pwd_clean = (password or "").strip()
        user = db.query(UserDB).filter(UserDB.username == u_clean).first()
        if user and check_password_hash(user.password_hash, pwd_clean):
            return {
                "id": user.id,
                "username": user.username,
                "role": user.role,
                "full_name": user.full_name or user.username,
                "allowed_views": parse_allowed_views(user.allowed_views, user.role),
            }

        fallback_users = {
            "agronomo": ("agro2026", "agronomo", "Agrónomo Operativo", ["agro"]),
            "visita": ("visita2026", "visita", "Usuario Visita / Coordinación", ["estadistica"]),
            "gerente": ("gerente2026", "visita", "Gerencia General", ["estadistica"]),
            "coordinador": ("coord2026", "visita", "Coordinador de Producción", ["estadistica"]),
            "admin": ("admin2026", "admin", "Administrador General", ["agro", "reales", "estadistica", "tpsr", "parametros"]),
        }
        if u_clean in fallback_users and pwd_clean == fallback_users[u_clean][0]:
            return {
                "id": 0,
                "username": u_clean,
                "role": fallback_users[u_clean][1],
                "full_name": fallback_users[u_clean][2],
                "allowed_views": fallback_users[u_clean][3],
            }
        return None
    except Exception:
        return None
    finally:
        db.close()


def list_users() -> list[dict[str, Any]]:
    db = SessionLocal()
    try:
        users = db.query(UserDB).order_by(UserDB.id.asc()).all()
        result = []
        for u in users:
            views_list = parse_allowed_views(u.allowed_views, u.role)
            result.append({
                "id": u.id,
                "username": u.username,
                "full_name": u.full_name or u.username,
                "role": u.role,
                "allowed_views": views_list,
                "allowed_views_str": ",".join(views_list),
                "created_at": u.created_at.strftime("%Y-%m-%d %H:%M") if u.created_at else "",
            })
        return result
    finally:
        db.close()


def get_user_by_id(user_id: int) -> dict[str, Any] | None:
    db = SessionLocal()
    try:
        u = db.query(UserDB).filter_by(id=user_id).first()
        if not u:
            return None
        views_list = parse_allowed_views(u.allowed_views, u.role)
        return {
            "id": u.id,
            "username": u.username,
            "full_name": u.full_name or u.username,
            "role": u.role,
            "allowed_views": views_list,
            "allowed_views_str": ",".join(views_list),
            "created_at": u.created_at.strftime("%Y-%m-%d %H:%M") if u.created_at else "",
        }
    finally:
        db.close()


def create_user(
    username: str,
    password: str,
    full_name: str = "",
    role: str = "visita",
    allowed_views: list[str] | str = None,
) -> tuple[bool, str, dict[str, Any] | None]:
    from werkzeug.security import generate_password_hash

    u_clean = (username or "").strip().lower()
    pwd_clean = (password or "").strip()
    if not u_clean:
        return False, "El nombre de usuario es obligatorio.", None
    if not pwd_clean:
        return False, "La contraseña es obligatoria.", None
    if len(pwd_clean) < 4:
        return False, "La contraseña debe tener al menos 4 caracteres.", None

    role_clean = role.strip().lower() if role else "visita"
    if role_clean not in {"admin", "agronomo", "visita"}:
        role_clean = "visita"

    if isinstance(allowed_views, list):
        views_str = ",".join(v.strip().lower() for v in allowed_views if v.strip())
    else:
        views_str = str(allowed_views or "").strip().lower()

    if not views_str:
        if role_clean == "agronomo":
            views_str = "agro"
        elif role_clean == "admin":
            views_str = "agro,reales,estadistica,tpsr,parametros"
        else:
            views_str = "estadistica"

    db = SessionLocal()
    try:
        existing = db.query(UserDB).filter(UserDB.username == u_clean).first()
        if existing:
            return False, f"El usuario '{u_clean}' ya existe.", None

        new_user = UserDB(
            username=u_clean,
            password_hash=generate_password_hash(pwd_clean),
            full_name=full_name.strip() if full_name else u_clean,
            role=role_clean,
            allowed_views=views_str,
            created_at=datetime.utcnow(),
        )
        db.add(new_user)
        db.commit()
        db.refresh(new_user)
        return True, "Usuario creado exitosamente.", {
            "id": new_user.id,
            "username": new_user.username,
            "full_name": new_user.full_name,
            "role": new_user.role,
            "allowed_views": parse_allowed_views(new_user.allowed_views, new_user.role),
        }
    except Exception as exc:
        db.rollback()
        return False, f"Error al crear usuario: {str(exc)}", None
    finally:
        db.close()


def update_user(
    user_id: int,
    username: str = None,
    password: str = None,
    full_name: str = None,
    role: str = None,
    allowed_views: list[str] | str = None,
) -> tuple[bool, str]:
    from werkzeug.security import generate_password_hash

    db = SessionLocal()
    try:
        user = db.query(UserDB).filter_by(id=user_id).first()
        if not user:
            return False, "Usuario no encontrado."

        if username:
            u_clean = username.strip().lower()
            if u_clean != user.username:
                existing = db.query(UserDB).filter(UserDB.username == u_clean, UserDB.id != user_id).first()
                if existing:
                    return False, f"El nombre de usuario '{u_clean}' ya está en uso."
                user.username = u_clean

        if password is not None and str(password).strip() != "":
            pwd_clean = str(password).strip()
            if len(pwd_clean) < 4:
                return False, "La contraseña debe tener al menos 4 caracteres."
            user.password_hash = generate_password_hash(pwd_clean)

        if full_name is not None:
            user.full_name = full_name.strip()

        if role:
            role_clean = role.strip().lower()
            if role_clean in {"admin", "agronomo", "visita"}:
                user.role = role_clean

        if allowed_views is not None:
            if isinstance(allowed_views, list):
                views_str = ",".join(v.strip().lower() for v in allowed_views if v.strip())
            else:
                views_str = str(allowed_views).strip().lower()
            user.allowed_views = views_str

        db.commit()
        return True, "Usuario actualizado exitosamente."
    except Exception as exc:
        db.rollback()
        return False, f"Error al actualizar: {str(exc)}"
    finally:
        db.close()


def delete_user(user_id: int, current_user_id: int = None) -> tuple[bool, str]:
    db = SessionLocal()
    try:
        user = db.query(UserDB).filter_by(id=user_id).first()
        if not user:
            return False, "Usuario no encontrado."
        if current_user_id and user.id == current_user_id:
            return False, "No puedes eliminar tu propia cuenta en uso."
        if user.username == "admin":
            return False, "No se puede eliminar la cuenta principal de administrador."

        db.delete(user)
        db.commit()
        return True, "Usuario eliminado correctamente."
    except Exception as exc:
        db.rollback()
        return False, f"Error al eliminar usuario: {str(exc)}"
    finally:
        db.close()


def seed_data_from_excel_if_empty(workbook_path: Path):
    init_db()
    db = SessionLocal()
    try:
        count = db.query(TpsrRecord).count()
        if count > 0:
            return  # Already seeded
        from seed_db import seed_data_from_excel
        seed_data_from_excel(workbook_path)
    finally:
        db.close()


def process_tpsr_excel_upload(file_input: Any, commit: bool = True) -> Dict[str, Any]:
    """
    Reads an uploaded TPSR Excel file (.xlsx), validates each row,
    performs deduplication against PostgreSQL, detects plant mismatches/corrections,
    and inserts/updates accordingly.
    """
    wb = load_workbook(file_input, data_only=True)
    sheet_name = "TPSR" if "TPSR" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    db = SessionLocal()
    try:
        # Load existing TPSR records for fast lookup
        existing_records = db.query(TpsrRecord).all()
        # Lookup dictionary: key -> TpsrRecord object
        # Natural key: (source_week, activity_norm, product_master_norm, variety_norm, block_norm, bed_location, pruning_number)
        lookup_dict: Dict[Tuple[int, str, str, str, str, str, str], TpsrRecord] = {}
        # Fallback dictionary without bed_location ONLY when unambiguous (count == 1 in DB and count == 1 in Excel)
        fallback_dict: Dict[Tuple[int, str, str, str, str], TpsrRecord] = {}
        db_fallback_counts: Dict[Tuple[int, str, str, str, str], int] = {}

        for rec in existing_records:
            act_c, pm_c, v_c, b_c = canonical_norms(rec.activity, rec.product_master_norm or rec.product_master, rec.variety_norm or rec.variety, rec.block_norm or rec.block)
            nat_key = (
                rec.source_week,
                act_c,
                pm_c,
                v_c,
                b_c,
                str(rec.bed_location or "").strip(),
                str(rec.pruning_number or "").strip(),
            )
            lookup_dict[nat_key] = rec
            fallback_key = (
                rec.source_week,
                act_c,
                pm_c,
                v_c,
                b_c,
            )
            fallback_dict[fallback_key] = rec
            db_fallback_counts[fallback_key] = db_fallback_counts.get(fallback_key, 0) + 1

        # Pre-count 5-tuple occurrences in uploaded Excel so multi-bed blocks never collide on fallback_dict
        excel_fallback_counts: Dict[Tuple[int, str, str, str, str], int] = {}
        for r in range(2, ws.max_row + 1):
            act_pre = normalize_text(ws.cell(r, 6).value)
            sw_pre = parse_week_code(ws.cell(r, 2).value)
            pm_pre = str(ws.cell(r, 7).value or "").strip()
            var_pre = str(ws.cell(r, 9).value or "").strip()
            blk_pre = str(ws.cell(r, 12).value or "").strip()
            if act_pre in {"SIEMBRA", "PODA"} and sw_pre is not None and pm_pre and var_pre and blk_pre:
                act_c, pm_c, v_c, b_c = canonical_norms(act_pre, pm_pre, var_pre, blk_pre)
                fk = (sw_pre, act_c, pm_c, v_c, b_c)
                excel_fallback_counts[fk] = excel_fallback_counts.get(fk, 0) + 1

        new_count = 0
        updated_count = 0
        skipped_count = 0
        invalid_count = 0

        new_records_list: List[TpsrRecord] = []
        plant_adjustments_details: List[Dict[str, Any]] = []

        for r in range(2, ws.max_row + 1):
            act = normalize_text(ws.cell(r, 6).value)
            if act not in {"SIEMBRA", "PODA"}:
                invalid_count += 1
                continue
            sw = parse_week_code(ws.cell(r, 2).value)
            if sw is None:
                invalid_count += 1
                continue

            pm = str(ws.cell(r, 7).value or "").strip()
            variety = str(ws.cell(r, 9).value or "").strip()
            block = str(ws.cell(r, 12).value or "").strip()
            bed_location = str(ws.cell(r, 13).value or "").strip()
            pruning_number = str(ws.cell(r, 5).value or "").strip()
            plants_cell_val = ws.cell(r, 15).value
            plants = int(round(float_or_zero(plants_cell_val)))

            if not pm or not variety or not block:
                invalid_count += 1
                continue

            act, pm_norm, v_norm, b_norm = canonical_norms(act, pm, variety, block)
            if pm_norm in {"GYPSOPHILA", "VERONICA", "HYPERICUM"}:
                pm = pm_norm
            variety = normalize_variety_name_db(variety, pm)
            block = normalize_block_name_db(block)

            nat_key = (sw, act, pm_norm, v_norm, b_norm, bed_location, pruning_number)
            fallback_key = (sw, act, pm_norm, v_norm, b_norm)

            existing_rec = lookup_dict.get(nat_key)
            if existing_rec is None:
                if db_fallback_counts.get(fallback_key, 0) == 1 and excel_fallback_counts.get(fallback_key, 0) == 1:
                    existing_rec = fallback_dict.get(fallback_key)

            if existing_rec:
                old_plants = existing_rec.plants
                beds_val = float_or_zero(ws.cell(r, 14).value)
                changed = False

                # Protect existing positive plant counts from blank/uncached Excel formula cells (0/None)
                if plants > 0 and existing_rec.plants != plants:
                    if commit:
                        existing_rec.plants = plants
                    changed = True
                    plant_adjustments_details.append({
                        "id": existing_rec.id,
                        "activity": act,
                        "product_master": pm,
                        "variety": variety,
                        "block": block,
                        "source_week": sw,
                        "source_week_short": format_short_week(sw),
                        "old_plants": old_plants,
                        "new_plants": plants,
                        "difference": plants - old_plants,
                    })

                if bed_location and existing_rec.bed_location != bed_location:
                    if commit:
                        existing_rec.bed_location = bed_location
                    changed = True

                if beds_val and existing_rec.total_beds != beds_val:
                    if commit:
                        existing_rec.total_beds = beds_val
                    changed = True

                if changed:
                    if commit:
                        existing_rec.updated_at = datetime.utcnow()
                    updated_count += 1
                else:
                    skipped_count += 1
            else:
                # Brand new record
                new_rec = TpsrRecord(
                    year=int(ws.cell(r, 1).value) if ws.cell(r, 1).value and str(ws.cell(r, 1).value).isdigit() else None,
                    source_week=sw,
                    source_week_short=format_short_week(sw),
                    date_val=str(ws.cell(r, 4).value or "").strip(),
                    pruning_number=pruning_number,
                    activity=act,
                    product_master=pm,
                    product_master_norm=pm_norm,
                    product=str(ws.cell(r, 8).value or "").strip(),
                    variety=variety,
                    variety_norm=v_norm,
                    color=str(ws.cell(r, 10).value or "").strip(),
                    material_origin=str(ws.cell(r, 11).value or "").strip(),
                    block=block,
                    block_norm=b_norm,
                    bed_location=bed_location,
                    total_beds=float_or_zero(ws.cell(r, 14).value),
                    plants=plants,
                )
                new_records_list.append(new_rec)
                lookup_dict[nat_key] = new_rec
                new_count += 1

        if commit:
            if new_records_list:
                db.bulk_save_objects(new_records_list)
            db.commit()

        return {
            "ok": True,
            "new_count": new_count,
            "updated_count": updated_count,
            "skipped_count": skipped_count,
            "invalid_count": invalid_count,
            "total_processed": new_count + updated_count + skipped_count,
            "plant_adjustments": plant_adjustments_details,
            "committed": commit,
        }
    except Exception as exc:
        db.rollback()
        return {"ok": False, "message": str(exc)}
    finally:
        db.close()
