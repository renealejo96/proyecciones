# Import historical block closures and weekly harvest cuts prior to week 2630
# Reference: extractor_siembra.py and Siembras_reales_Variedades_Farms_2642.xlsx

import os
import sys
import re
import time
import openpyxl
from datetime import datetime

sys.path.insert(0, r"D:\todo en vs code\PROYECCIONES")

from db import SessionLocal, TpsrRecord, WeekAdjustmentDB, BlockClosureDB
from app import normalize_text, parse_week_code

excel_path = r"C:\Users\ASUS\Elite Flower\EC-PYGAN-PRODUCION - Documentos\Novel_Producción\Proyección\2025\Variedades\2026\Siembras_reales_Variedades_Farms_2642.xlsx"

SHEET_CONFIG = {
    'Solidago FS1': {'pm': 'SOLIDAGO', 'header_row': 11},
    'Sunflower 2025_inv': {'pm': 'SUNFLOWER', 'header_row': 11},
    'Hypericum_inv': {'pm': 'HYPERICUM', 'header_row': 27},
    'Romantic': {'pm': 'SOLIDAGO', 'header_row': 11},
    'Veronicas_splash': {'pm': 'VERONICA', 'header_row': 11},
    'Veronicas Novel': {'pm': 'VERONICA', 'header_row': 11},
    'Gyp Codigos': {'pm': 'GYPSOPHILA', 'header_row': 12},
    'BILLION LIGHTS': {'pm': 'GYPSOPHILA', 'header_row': 11},
}

def comparable_text(value):
    normalized = normalize_text(value)
    if not normalized:
        return ''
    return re.sub(r'[^A-Z0-9]+', ' ', normalized).strip()

def derive_activity(source_sheet_name, original_value, marker_raw=None):
    original_normalized = normalize_text(original_value)
    if original_normalized and 'PODA' in comparable_text(original_normalized):
        return 'PODA'
    if source_sheet_name.strip().lower() == 'hypericum_inv' and marker_raw:
        marker = normalize_text(marker_raw)
        if marker == 'P':
            return 'PODA'
    return 'SIEMBRA'

def match_canonical_name(value, valid_names):
    comparable_value = comparable_text(value)
    for name in sorted(valid_names, key=len, reverse=True):
        if comparable_text(name) in comparable_value:
            return name
    return normalize_text(value)

def normalize_variedad(source_sheet_name, value):
    normalized = normalize_text(value)
    if not normalized:
        return normalized

    normalized = re.sub(r'\s+(PODA|SIEMBRA)$', '', normalized).strip()
    normalized_sheet_key = source_sheet_name.strip().lower()
    comparable = comparable_text(normalized)

    if normalized_sheet_key == 'hypericum_inv':
        if 'TRIUMPH' in comparable or ('RED' in comparable and 'TRIUMPH' in comparable):
            return 'MAGICAL TRIUMPH'
        if 'RED WINE' in comparable or 'BURGUNDY' in comparable:
            return 'MAGICAL RED WINE'
        if 'ROYAL PRINCESS' in comparable or 'ORANGE' in comparable:
            return 'MAGICAL ROYAL PRINCESS'
        if 'RENATA' in comparable or ('GREEN' in comparable and 'RENATA' in comparable):
            return 'MAGICAL RENATA GREEN'
        if 'SNOW PRINCE' in comparable or 'SNOWPRINCE' in comparable or ('WHITE' in comparable and 'PRINCE' in comparable):
            return 'MAGICAL SNOW PRINCE'
        if 'MAJESTIC PINK' in comparable or 'HOT PINK' in comparable:
            return 'MAGICAL MAJESTIC PINK'
        canonical = match_canonical_name(normalized, [
            'MAGICAL CHERRY',
            'MAGICAL MAJESTIC PINK',
            'MAGICAL RED WINE',
            'MAGICAL RENATA GREEN',
            'MAGICAL ROYAL PRINCESS',
            'MAGICAL SNOW PRINCE',
            'MAGICAL TRIUMPH'
        ])
        return canonical

    if normalized_sheet_key == 'veronicas_splash':
        if 'SPLASH BLUE' in comparable:
            return 'SKYLER SPLASH BLUE'
        if 'SPLASH PINK' in comparable:
            return 'SKYLER SPLASH PINK'
        if 'SPLASH WHITE' in comparable:
            return 'SKYLER SPLASH WHITE'
        return match_canonical_name(normalized, [
            'SKYLER SPLASH BLUE',
            'SKYLER SPLASH PINK',
            'SKYLER SPLASH WHITE'
        ])

    if normalized_sheet_key == 'veronicas novel':
        if 'WHITE' in comparable:
            return 'SKYLER WHITE IMPROVED'
        if 'BLUE' in comparable:
            return 'SKYLER BLUE'
        if 'PINK' in comparable:
            return 'SKYLER PINK'
        return match_canonical_name(normalized, [
            'SKYLER BLUE',
            'SKYLER PINK',
            'SKYLER WHITE IMPROVED'
        ])

    if normalized_sheet_key == 'solidago fs1':
        if 'ROMANTIC' in comparable:
            return 'ROMANTIC GLORY'
        return 'GOLDEN GLORY YELLOW'

    if normalized_sheet_key == 'romantic':
        return 'ROMANTIC GLORY'

    if normalized_sheet_key in {'gyp codigos', 'billion lights'}:
        return 'BILLION LIGHTS'

    if normalized_sheet_key == 'sunflower 2025_inv':
        normalized = re.sub(r'\bPLANTULA\b', '', normalized).strip()
        normalized = re.sub(r'\s+', ' ', normalized)
        comparable = comparable_text(normalized)
        if comparable in {'H10', 'H2'}:
            return comparable
        if 'SUNRICH' in comparable and 'GOLD' in comparable:
            return 'SUNRICH GOLD'
        if 'VINCENT' in comparable and 'CHOICE' in comparable:
            return 'VINCENTS CHOICE'
        return normalized

    return normalized

def clean_block(val):
    if val is None:
        return ""
    b_str = str(val).strip()
    b_str = re.sub(r'^Bl\.?\s*', '', b_str, flags=re.IGNORECASE).strip()
    return b_str

def run_import():
    t0 = time.time()
    print("Iniciando importación histórica desde Excel...")
    
    db = SessionLocal()
    
    # Pre-load existing TpsrRecord keys
    print("Cargando registros existentes en memoria...")
    existing_tpsr = {}
    for r in db.query(TpsrRecord.id, TpsrRecord.product_master_norm, TpsrRecord.variety_norm, TpsrRecord.activity, TpsrRecord.block_norm, TpsrRecord.source_week, TpsrRecord.plants).all():
        k = (r.product_master_norm, r.variety_norm, r.activity, r.block_norm, r.source_week)
        existing_tpsr[k] = (r.id, r.plants)
    print(f"  -> {len(existing_tpsr)} TpsrRecords existentes.")

    existing_adj = set()
    for a in db.query(WeekAdjustmentDB.product_master_norm, WeekAdjustmentDB.variety_norm, WeekAdjustmentDB.activity, WeekAdjustmentDB.block_norm, WeekAdjustmentDB.source_week, WeekAdjustmentDB.harvest_week).all():
        k = (a.product_master_norm, a.variety_norm, a.activity, a.block_norm, a.source_week, a.harvest_week)
        existing_adj.add(k)
    print(f"  -> {len(existing_adj)} WeekAdjustmentDB existentes.")

    existing_closures = {}
    for c in db.query(BlockClosureDB).all():
        k = (c.product_master_norm, c.variety_norm, c.activity, c.block_norm, c.source_week)
        existing_closures[k] = c
    print(f"  -> {len(existing_closures)} BlockClosureDB existentes.")

    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)

    tpsr_to_insert = []
    adj_to_insert = []
    closure_to_insert = []
    
    seen_new_tpsr = set()
    seen_new_adj = set()
    seen_new_closure = set()

    tpsr_updated = 0
    closures_updated = 0

    for sheet_name, cfg in SHEET_CONFIG.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        pm = cfg['pm']
        pm_norm = normalize_text(pm)
        h_row = cfg['header_row']
        
        # Header
        header = None
        for idx, row in enumerate(ws.iter_rows(min_row=h_row, max_row=h_row, values_only=True)):
            header = [str(v).strip().upper() if v is not None else "" for v in row]
            break
        if not header:
            continue
        
        b_idx = [i for i, h in enumerate(header) if h == 'BLOQUE'][0]
        v_idx = [i for i, h in enumerate(header) if 'VARIEDAD' in h][0]
        sw_idx = [i for i, h in enumerate(header) if 'SEMANA' in h and 'SIEMBRA' in h][0]
        
        prog_idx = [i for i, h in enumerate(header) if 'PROGRAMA' in h]
        prog_idx = prog_idx[0] if prog_idx else -1
        
        week_cols = []
        for c_idx, col_name in enumerate(header):
            val_str = col_name.strip()
            if val_str.isdigit() and (2300 <= int(val_str) <= 2700):
                week_cols.append((c_idx, int(val_str)))

        for row in ws.iter_rows(min_row=h_row+1, values_only=True):
            cierre = str(row[0]).strip().upper() if row[0] is not None else ""
            if cierre != 'C':
                continue
            
            blk_raw = row[b_idx] if b_idx < len(row) else None
            blk = clean_block(blk_raw)
            if not blk:
                continue
            blk_norm = normalize_text(blk)
            
            sw_raw = row[sw_idx] if sw_idx < len(row) else None
            if sw_raw is None or str(sw_raw).strip() == "":
                continue
            try:
                sw_int = int(float(sw_raw))
            except:
                continue
            
            sw_full = parse_week_code(sw_int)
            if not sw_full:
                continue
            
            var_raw = row[v_idx] if v_idx < len(row) else ""
            marker_raw = row[1] if sheet_name == 'Hypericum_inv' and len(row) > 1 else None
            act = derive_activity(sheet_name, var_raw, marker_raw)
            var_clean = normalize_variedad(sheet_name, var_raw)
            var_norm = normalize_text(var_clean)
            
            plants = 0
            if prog_idx != -1 and prog_idx < len(row) and row[prog_idx] is not None:
                try:
                    plants = int(round(float(row[prog_idx])))
                except:
                    pass
            
            # Pre-30 harvest cuts
            pre30_cuts = []
            has_post30_cut = False
            for c_i, wk in week_cols:
                val = row[c_i] if c_i < len(row) else None
                if val is not None and str(val).strip() != "":
                    try:
                        stems = float(val)
                        if stems > 0:
                            if wk < 2630:
                                pre30_cuts.append((parse_week_code(wk), int(round(stems))))
                            else:
                                has_post30_cut = True
                    except:
                        pass
            
            if not pre30_cuts:
                continue

            # 1. TpsrRecord
            tpsr_key = (pm_norm, var_norm, act, blk_norm, sw_full)
            if tpsr_key not in existing_tpsr and tpsr_key not in seen_new_tpsr:
                seen_new_tpsr.add(tpsr_key)
                tpsr_to_insert.append(TpsrRecord(
                    year=int(str(sw_full)[:4]),
                    source_week=sw_full,
                    source_week_short=str(sw_full)[2:],
                    activity=act,
                    product_master=pm,
                    product_master_norm=pm_norm,
                    product=pm,
                    variety=var_clean,
                    variety_norm=var_norm,
                    block=blk,
                    block_norm=blk_norm,
                    plants=plants,
                ))

            # 2. WeekAdjustmentDB for pre-30 cuts
            for hw_full, stems in pre30_cuts:
                adj_key = (pm_norm, var_norm, act, blk_norm, sw_full, hw_full)
                if adj_key not in existing_adj and adj_key not in seen_new_adj:
                    seen_new_adj.add(adj_key)
                    adj_to_insert.append(WeekAdjustmentDB(
                        activity=act,
                        product_master_norm=pm_norm,
                        variety_norm=var_norm,
                        block=blk,
                        block_norm=blk_norm,
                        source_week=sw_full,
                        harvest_week=hw_full,
                        agronomo_estimate=stems,
                        real_closed=stems,
                        dump_stems=0,
                        is_dump=False,
                    ))

            # 3. BlockClosureDB (mark closed if no ongoing post-30 cuts)
            if not has_post30_cut:
                cls_key = (pm_norm, var_norm, act, blk_norm, sw_full)
                if cls_key in existing_closures:
                    c_obj = existing_closures[cls_key]
                    if not c_obj.is_closed:
                        c_obj.is_closed = True
                        c_obj.updated_at = datetime.utcnow()
                        closures_updated += 1
                elif cls_key not in seen_new_closure:
                    seen_new_closure.add(cls_key)
                    closure_to_insert.append(BlockClosureDB(
                        activity=act,
                        product_master_norm=pm_norm,
                        variety_norm=var_norm,
                        block=blk,
                        block_norm=blk_norm,
                        source_week=sw_full,
                        is_closed=True,
                    ))

    wb.close()

    print(f"\nGuardando en base de datos PostgreSQL:")
    print(f"  -> TpsrRecord a insertar: {len(tpsr_to_insert)}")
    print(f"  -> WeekAdjustmentDB a insertar: {len(adj_to_insert)}")
    print(f"  -> BlockClosureDB a insertar: {len(closure_to_insert)}")
    print(f"  -> BlockClosureDB actualizados a cerrado: {closures_updated}")

    # Bulk insert in batches of 1000
    batch_size = 1000
    for i in range(0, len(tpsr_to_insert), batch_size):
        db.bulk_save_objects(tpsr_to_insert[i:i+batch_size])
    db.commit()

    for i in range(0, len(adj_to_insert), batch_size):
        db.bulk_save_objects(adj_to_insert[i:i+batch_size])
    db.commit()

    for i in range(0, len(closure_to_insert), batch_size):
        db.bulk_save_objects(closure_to_insert[i:i+batch_size])
    db.commit()

    db.close()
    print(f"Importación completada en {time.time() - t0:.2f} segundos.")

if __name__ == "__main__":
    run_import()
