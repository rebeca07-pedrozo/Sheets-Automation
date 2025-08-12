import pandas as pd
import numpy as np

emisiones = pd.read_csv('/content/Ventas3Agosto.csv')
emisiones['prima'] = pd.to_numeric(emisiones['prima'], errors='coerce')

agrupado = emisiones.groupby('numero_poliza').agg({
    'clave_agente': 'first',
    'codigo_producto': 'first',
    'fecha_emision': 'first',
    'nombre_producto': 'first',
    'nombre_opcion_poliza': 'first',
    'prima': 'sum',
    'tipo_documento': 'first',
    'FECHA_PROCESO': 'first',
    'NUMERO_DOCUMENTO': lambda x: list(pd.unique(x)),
    'NOMBRE': lambda x: list(pd.unique(x)),
    'CORREO_1': lambda x: list(pd.unique(x)),
    'CELULAR_1': lambda x: list(pd.unique(x)),
}).reset_index()

def expandir_columnas(df, columna_base, prefijo):
    max_items = df[columna_base].apply(len).max()
    nuevas_cols = pd.DataFrame(
        df[columna_base].tolist(),
        columns=[f"{prefijo}_{i+1}" for i in range(max_items)]
    )
    return pd.concat([df.drop(columns=[columna_base]), nuevas_cols], axis=1)

agrupado = expandir_columnas(agrupado, 'NUMERO_DOCUMENTO', 'NUMERO_DOCUMENTO')
agrupado = expandir_columnas(agrupado, 'NOMBRE', 'NOMBRE')
agrupado = expandir_columnas(agrupado, 'CORREO_1', 'CORREO')
agrupado = expandir_columnas(agrupado, 'CELULAR_1', 'CELULAR')

agrupado['prima'] = agrupado['prima'].fillna(0).astype('int64')
for col in agrupado.columns:
    if agrupado[col].dtype == 'float64':
        agrupado[col] = agrupado[col].fillna(0).astype('int64')

agrupado.to_csv("emisionesAgrupadas.csv", index=False)
