#!/usr/bin/env python3
"""
Convierte Gastorefacciones.xlsx → datos_YYYY.json + diccionarios.json
Archivos ligeros para cargar en GitHub Pages.

Ejecutar: python convertir_datos.py
Requiere: pip install openpyxl
"""
import openpyxl, json, sys, os
from datetime import datetime, timedelta
from collections import defaultdict

INPUT_FILE  = 'Gastorefacciones.xlsx'
OUTPUT_DIR  = '.'  # Carpeta donde se generan los JSON

# ============================================================
# Columnas del Excel (por índice)
# ============================================================
#  0  Nombre Cta           7  CLASIFICACION       14 DIVISION         21 TIPO CHASIS
#  1  Fecha Contable       8  SISTEMA             15 VALE             22 MARCA MOTOR
#  2  Importe              9  COMPONENTE          16 TIPO UNIDAD      23 TIPO MOTOR
#  3  no unidad           10  TIPO CAPITALIZACION 17 AÑO MODELO       24 TIPO COMBUSTIBLE
#  4  Unidades            11  PLAN                18 MARCA CARROCERIA  25 TIPO MANTENIMIENTO
#  5  concepto            12  Detalle Capital.    19 TIPO CARROCERIA
#  6  originador Batch    13  ZONA                20 MARCA CHASIS

# Columnas que se codifican con diccionario (texto repetitivo)
DICT_COLS_DEF = [
    (0,  'cuenta'),
    (3,  'eco'),
    (6,  'origBatch'),
    (7,  'clasif'),
    (8,  'sistema'),
    (9,  'comp'),
    (10, 'tipoCap'),
    (11, 'plan'),
    (12, 'detCap'),
    (13, 'zona'),
    (14, 'div'),
    (15, 'vale'),
    (16, 'tipoUni'),
    (17, 'anioMod'),
    (18, 'marcaCarr'),
    (19, 'tipoCarr'),
    (20, 'marcaChas'),
    (21, 'tipoChas'),
    (22, 'marcaMot'),
    (23, 'tipoMot'),
    (24, 'tipoComb'),
    (25, 'tipoMant'),
]

# Orden de campos en cada fila de salida (8 valores):
# [cuenta_i, fechaInt, importe, eco_i, origBatch_i, clasif_i, sistema_i, comp_i,
#  tipoCap_i, plan_i, detCap_i, zona_i, div_i, vale_i, tipoUni_i, anioMod_i,
#  marcaCarr_i, tipoCarr_i, marcaChas_i, tipoChas_i, marcaMot_i, tipoMot_i,
#  tipoComb_i, tipoMant_i]

COLUMNS_OUT = [
    'cuenta','fecha','importe','eco','origBatch','clasif','sistema','comp',
    'tipoCap','plan','detCap','zona','div','vale','tipoUni','anioMod',
    'marcaCarr','tipoCarr','marcaChas','tipoChas','marcaMot','tipoMot',
    'tipoComb','tipoMant'
]

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"ERROR: No se encuentra '{INPUT_FILE}'.")
        sys.exit(1)

    print(f"Leyendo {INPUT_FILE}...")
    wb = openpyxl.load_workbook(INPUT_FILE, read_only=True, data_only=True)
    ws = wb['BD']

    # ── Paso 1: Leer y construir diccionarios ──
    print("Paso 1/3: Leyendo filas y construyendo diccionarios...")
    dicts = {name: {} for _, name in DICT_COLS_DEF}
    rows_by_year = defaultdict(list)
    count = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None and row[1] is None and row[2] is None:
            break

        # Fecha
        fecha = row[1]
        if isinstance(fecha, datetime):
            y, m, d = fecha.year, fecha.month, fecha.day
        elif isinstance(fecha, (int, float)):
            dt = datetime(1899, 12, 30) + timedelta(days=int(fecha))
            y, m, d = dt.year, dt.month, dt.day
        else:
            continue
        fecha_int = y * 10000 + m * 100 + d

        # Importe
        importe = round(float(row[2]), 2) if row[2] is not None else 0

        # Diccionarizar todos los campos de texto
        indices = []
        for col_idx, name in DICT_COLS_DEF:
            v = row[col_idx]
            s = str(v).strip() if v is not None else ''
            # Limpiar algunos valores
            if isinstance(v, float) and col_idx in (3, 17):
                s = str(int(v)) if v == int(v) else str(v)
            if s not in dicts[name]:
                dicts[name][s] = len(dicts[name])
            indices.append(dicts[name][s])

        # Fila compacta: [cuenta_i, fecha, importe, eco_i, origBatch_i, ...]
        compact = [indices[0], fecha_int, importe] + indices[1:]
        rows_by_year[y].append(compact)

        count += 1
        if count % 200000 == 0:
            print(f"  ...{count:,} filas")

    wb.close()
    print(f"  Total: {count:,} filas en {len(rows_by_year)} años")

    # ── Paso 2: Invertir diccionarios ──
    print("Paso 2/3: Generando diccionarios invertidos...")
    inv_dicts = {}
    for _, name in DICT_COLS_DEF:
        inv = [''] * len(dicts[name])
        for val, idx in dicts[name].items():
            inv[idx] = val
        inv_dicts[name] = inv

    # Guardar diccionarios
    dict_data = {
        '_version': 3,
        '_columns': COLUMNS_OUT,
        '_dictPositions': {
            'cuenta':0, 'eco':3, 'origBatch':4, 'clasif':5, 'sistema':6,
            'comp':7, 'tipoCap':8, 'plan':9, 'detCap':10, 'zona':11,
            'div':12, 'vale':13, 'tipoUni':14, 'anioMod':15,
            'marcaCarr':16, 'tipoCarr':17, 'marcaChas':18, 'tipoChas':19,
            'marcaMot':20, 'tipoMot':21, 'tipoComb':22, 'tipoMant':23
        },
        'dicts': inv_dicts,
        'years': sorted(rows_by_year.keys())
    }

    dict_path = os.path.join(OUTPUT_DIR, 'diccionarios.json')
    with open(dict_path, 'w', encoding='utf-8') as f:
        json.dump(dict_data, f, ensure_ascii=False, separators=(',', ':'))
    print(f"  diccionarios.json: {os.path.getsize(dict_path)/1024:.0f} KB")

    # ── Paso 3: Guardar un archivo por año ──
    print("Paso 3/3: Generando archivos por año...")
    total_size = os.path.getsize(dict_path)

    for year in sorted(rows_by_year.keys()):
        rows = rows_by_year[year]
        fname = f'datos_{year}.json'
        fpath = os.path.join(OUTPUT_DIR, fname)
        with open(fpath, 'w', encoding='utf-8') as f:
            json.dump(rows, f, ensure_ascii=False, separators=(',', ':'))
        size = os.path.getsize(fpath)
        total_size += size
        print(f"  {fname}: {len(rows):>10,} filas → {size/1048576:.1f} MB")

    print(f"\n{'='*50}")
    print(f"✅ Total: {total_size/1048576:.1f} MB en {len(rows_by_year)+1} archivos")
    print(f"\nArchivos a subir a tu repositorio de GitHub:")
    print(f"  diccionarios.json")
    for y in sorted(rows_by_year.keys()):
        print(f"  datos_{y}.json")
    print(f"\nTodos caben en GitHub (< 25 MB cada uno).")

if __name__ == '__main__':
    main()
