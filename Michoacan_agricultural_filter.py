"""
Filtrado y Agregación de Datos Agrícolas SIAP - Michoacán
==========================================================
Lee archivos CSV del dataset abierto del SIAP, filtra registros de Michoacán
con modalidad Temporal, normaliza columnas entre versiones del dataset (2015-2020),
agrega por cultivo/municipio/ciclo y calcula el Rendimiento.
Genera un CSV limpio por cada archivo de entrada.

Uso:
    Ajusta las rutas input_folder / output_folder al final del script y ejecuta:
    python FilterData_Script.py

Requisitos:
    pip install pandas
"""

import pandas as pd
import os

def process_file(input_path: str, output_path: str):
    df = pd.read_csv(input_path, encoding='latin1', low_memory=False)

    # Normalize column names for 2015-2020 files
    df = df.rename(columns={
        'Nomcultivo Sin Um': 'Nomcultivo',
        'Precio':            'Preciomediorural'
    })

    # Force numeric columns (handles mixed types / dirty data)
    numeric_cols = ['Sembrada', 'Cosechada', 'Siniestrada', 'Volumenproduccion', 'Valorproduccion', 'Preciomediorural']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Filter Michoacán + Temporal
    df = df[
        df['Nomestado'].str.contains('Michoac', na=False) &
        (df['Nommodalidad'] == 'Temporal')
    ]

    # Keep only the columns we need
    df = df[['Anio', 'Nomestado', 'Nommunicipio', 'Nomcicloproductivo', 'Nommodalidad',
             'Nomunidad', 'Nomcultivo', 'Sembrada', 'Cosechada', 'Siniestrada',
             'Volumenproduccion', 'Valorproduccion', 'Preciomediorural']]

    # Aggregate by all identifier columns
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
