import os
import math
import pandas as pd
import numpy as np
from openpyxl import load_workbook


file_path_men261224 = '/content/Reporte_general__Caracterizacion__novedades_y_requisitos_politica_de_gratuidad__para_las_IES__26_12_2024_cia.xlsx'
output_pathXlsx = '/content/AuditoriaPiam20242Conciliacion.xlsx'

def agregar_mensaje(doc, mensaje):
    print(mensaje)
    doc.add_paragraph(mensaje)

def cargar_archivos_y_dataframes(file_path):
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"{file_path} no encontrado.")
    print(f"Archivo {file_path} encontrado.")
    try:
        dic_insumos = pd.read_excel(file_path, sheet_name=['plantilla_gratuidad_ies'], engine='openpyxl')
        for df in dic_insumos.values():
            df.columns = df.columns.str.strip()
            print('Archivo cargado')
        return dic_insumos['plantilla_gratuidad_ies']
    except Exception as e:
        raise Exception(f"Error al cargar los DataFrames: {e}")

def ajustadorConciliacion(df):
    if 'NUM_DOCUMENTO' not in df.columns or 'PRO_CONSECUTIVO' not in df.columns:
        raise ValueError("El DataFrame no contiene las columnas 'NUM_DOCUMENTO' o 'PRO_CONSECUTIVO'")
    df['ID-SNIES'] = df['NUM_DOCUMENTO'].astype(str) + '-' + df['PRO_CONSECUTIVO'].astype(str)
    columnas = ['ID-SNIES'] + [col for col in df.columns if col != 'ID-SNIES']
    df = df[columnas]
    return df

# Extraccion
conciliacionMen = cargar_archivos_y_dataframes(file_path_men261224)
# Manipulación
## DataFrame  Conciliacion
dfConciliacionMen = ajustadorConciliacion(conciliacionMen)


# Carga
with pd.ExcelWriter(output_pathXlsx, engine='xlsxwriter') as writer:
    if dfConciliacionMen is not None:
        dfConciliacionMen.to_excel(writer, sheet_name='conciliacion2024_2', index=False)

print("Los resultados han sido guardados en el documento y archivo Excel.")