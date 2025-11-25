import pandas as pd

valores = pd.read_csv('/content/bquxjob_73cf28b2_19a987b8f5a.csv')
personas = pd.read_csv('/content/ventas 17 nov - EMISIONES SALUD.csv')

print(personas.columns.tolist())

agrupado = personas.groupby('numero_poliza').agg({
    'clave_agente': 'first',
    'codigo_producto': 'first',
    'fecha_emision': 'first',
    'nombre_producto': 'first',
    'nombre_opcion_poliza': 'first',
    'tipo_documento': 'first',
    'NUMERO_DOCUMENTO': lambda x: ', '.join(sorted(set(x.dropna().astype(str)))),
    'NOMBRE': lambda x: ', '.join(sorted(set(x.dropna().astype(str)))),
    'CORREO_1': lambda x: ', '.join(sorted(set(x.dropna().astype(str)))),
    'CELULAR_1': lambda x: ', '.join(sorted(set(x.dropna().astype(str)))),
    'FECHA_PROCESO': 'first'
}).reset_index()

resultado = pd.merge(agrupado, valores[['numero_poliza', 'PRIMA']], on='numero_poliza', how='left')

resultado = resultado.rename(columns={
    'PRIMA': 'Prima_totalizada',
    'FECHA_PROCESO': 'FECG'
})

def expandir_columna(df, columna_base, prefijo):
    df[columna_base] = df[columna_base].fillna('')
    listas = df[columna_base].apply(lambda x: [i.strip() for i in x.split(',') if i.strip()])
    max_items = listas.apply(len).max()
    nuevas_cols = pd.DataFrame(listas.tolist(), columns=[f"{prefijo}_{i+1}" for i in range(max_items)])
    return pd.concat([df.drop(columns=[columna_base]), nuevas_cols], axis=1)

resultado = expandir_columna(resultado, 'NUMERO_DOCUMENTO', 'Numero_documento')

columnas_finales = [
    'emisiones',
    'revision agente',
    'codigo_producto',
    'clave_agente',
    'numero_poliza',
    'fecha_emision',
    'nombre_producto',
    'nombre_opcion_poliza',
    'Prima_totalizada',
    'CORREO_1',
    'Numero_documento_1',
    'tipo_documento',
    'Numero_documento_2',    
    'FECG'
]

for col in columnas_finales:
    if col not in resultado.columns:
        resultado[col] = ''

resultado_final = resultado[columnas_finales]
resultado_final.to_csv('Salud cruzado.csv', index=False, encoding='utf-8-sig')
print("Archivo generado: Consolidado Primas.csv")

