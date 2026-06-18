"""
Michoacan Agricultural Data Filter - SIAP
==========================================
Processes open agricultural production CSV files from SIAP (Servicio de
Informacion Agroalimentaria y Pesquera) and produces clean, aggregated CSVs
for Michoacan, temporal modality only.

What it does, step by step:
    1. Reads each CSV with Latin-1 encoding (SIAP files are not UTF-8).
    2. Normalizes column names that changed between dataset versions:
           'Nomcultivo Sin Um'  ->  'Nomcultivo'       (pre-2020 datasets)
           'Precio'             ->  'Preciomediorural'  (pre-2020 datasets)
    3. Converts area, volume, and price columns to numeric, replacing any
       non-numeric value (dashes, blanks, etc.) with 0.
    4. Filters rows: keeps only Michoacan (partial match on 'Michoac') and
       Modalidad == 'Temporal'. To change the state or modality, edit these
       two conditions in the process_file function.
    5. Keeps only the columns listed below. To add or remove columns, edit
       the column list in the 'Keep only the columns we need' block:
           Anio, Nomestado, Nommunicipio, Nomcicloproductivo, Nommodalidad,
           Nomunidad, Nomcultivo, Sembrada, Cosechada, Siniestrada,
           Volumenproduccion, Valorproduccion, Preciomediorural
    6. Aggregates by crop/municipality/year/cycle: sums areas and volumes,
       averages price across multiple records for the same group.
    7. Computes yield: Rendimiento = Volumenproduccion / Cosechada.
       Set to 0 if no area was harvested (avoids division by zero).
    8. Renames columns to cleaner Spanish display names:
           Anio -> Año,  Nommunicipio -> Municipio,  Nomcultivo -> Cultivo, etc.
    9. Rounds all numeric columns to 2 decimal places.
    10. Exports one clean CSV per input file (UTF-8 with BOM for Excel).

Configuration (set at the bottom of the script):
    input_folder    Folder containing the raw SIAP CSV files.
    output_folder   Folder where cleaned CSVs will be saved.

Usage:
    1. Set input_folder and output_folder.
    2. python Michoacan_agricultural_filter.py

Requirements:
    pip install pandas
"""

import pandas as pd
import os

def process_file(input_path: str, output_path: str):
    df = pd.read_csv(input_path, encoding='latin1', low_memory=False)
    # encoding='latin1': SIAP files are not UTF-8; using 'utf-8' would cause decoding errors
    # low_memory=False: reads the full file at once to infer column types correctly and avoid mixed-type warnings

    # Normalize column names for 2015-2020 files
    df = df.rename(columns={
        'Nomcultivo Sin Um': 'Nomcultivo',
        'Precio':            'Preciomediorural'
    })

    # Force numeric columns (handles mixed types / dirty data)
    # errors='coerce' turns non-numeric values (dashes, blanks) into NaN instead of raising an error
    # fillna(0) then replaces those NaN values with 0 so aggregations work correctly
    numeric_cols = ['Sembrada', 'Cosechada', 'Siniestrada', 'Volumenproduccion', 'Valorproduccion', 'Preciomediorural']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Filter Michoacán + Temporal
    # str.contains('Michoac') is a partial match to cover 'Michoacán de Ocampo'; na=False avoids errors on empty cells
    df = df[
        df['Nomestado'].str.contains('Michoac', na=False) &
        (df['Nommodalidad'] == 'Temporal')
    ]

    # Keep only the columns we need
    df = df[['Anio', 'Nomestado', 'Nommunicipio', 'Nomcicloproductivo', 'Nommodalidad',
             'Nomunidad', 'Nomcultivo', 'Sembrada', 'Cosechada', 'Siniestrada',
             'Volumenproduccion', 'Valorproduccion', 'Preciomediorural']]

    # Aggregate by all identifier columns
    # as_index=False keeps the groupby columns as regular columns instead of moving them to the DataFrame index
    # Areas and volumes are summed; price is averaged across records in the same group
    agg = df.groupby(
        ['Anio', 'Nomestado', 'Nommunicipio', 'Nomcicloproductivo', 'Nommodalidad', 'Nomunidad', 'Nomcultivo'],
        as_index=False
    ).agg(
        Sembrada          = ('Sembrada',          'sum'),
        Cosechada         = ('Cosechada',         'sum'),
        Siniestrada       = ('Siniestrada',       'sum'),
        Volumenproduccion = ('Volumenproduccion',  'sum'),
        Valorproduccion   = ('Valorproduccion',    'sum'),
        Preciomediorural  = ('Preciomediorural',  'mean'),
    )

    # Rendimiento = Volumenproduccion / Cosechada
    # The lambda checks for zero before dividing to avoid a ZeroDivisionError when no area was harvested
    agg['Rendimiento'] = agg.apply(
        lambda r: r['Volumenproduccion'] / r['Cosechada'] if r['Cosechada'] != 0 else 0,
        axis=1
    )
     # Rename columns to Spanish display names
    agg = agg.rename(columns={
        'Anio':               'Año',
        'Nomestado':          'Estado',
        'Nommunicipio':       'Municipio',
        'Nomcicloproductivo': 'Ciclo productivo',
        'Nommodalidad':       'Modalidad',
        'Nomunidad':          'Unidad',
        'Nomcultivo':         'Cultivo'
    })
   # Round all numeric columns to 2 decimals
    numeric_cols = ['Sembrada', 'Cosechada', 'Siniestrada', 'Volumenproduccion',
                    'Valorproduccion', 'Preciomediorural', 'Rendimiento']
    agg[numeric_cols] = agg[numeric_cols].round(2)

    agg.to_csv(output_path, index=False, encoding='utf-8-sig')
    # utf-8-sig adds a BOM (byte order mark) at the start of the file so Excel opens it correctly without garbled characters
    
    print(f"✓ {os.path.basename(input_path)} → {os.path.basename(output_path)}")


# ── Run on all CSVs in a folder ───────────────────────────────────────
input_folder  = r"C:\Users\zomav\Downloads\Raw Data"
output_folder = r"C:\Users\zomav\Downloads\Datos Agricolas Filtrados"

os.makedirs(output_folder, exist_ok=True)

for filename in os.listdir(input_folder):
    if filename.endswith('.csv'):
        process_file(
            os.path.join(input_folder, filename),
            os.path.join(output_folder, filename.replace('.csv', '_procesado.csv'))
        )
