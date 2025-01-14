import os
import math
import pandas as pd
import numpy as np
from openpyxl import load_workbook


file_path_men261224 = '/content/Reporte_general__Caracterizacion__novedades_y_requisitos_politica_de_gratuidad__para_las_IES__26_12_2024_cia.xlsx'
file_path_sq131224 ='/content/SqEnero13_2024_2.xlsx'
output_pathXlsx = '/content/AuditoriaPiam20242Conciliacion.xlsx'

def agregarMensaje(doc, mensaje):
    print(mensaje)
    doc.add_paragraph(mensaje)

def cargarArchivosDataframes(file_path,sheet_name):
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"{file_path} no encontrado.")
    print(f"Archivo {file_path} encontrado.")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        df.columns = df.columns.str.strip()
        print(f"Hoja '{sheet_name}' cargada exitosamente.")
        return df
    except ValueError:
        raise ValueError(f"La hoja '{sheet_name}' no existe en el archivo {file_path}.")
    except Exception as e:
        raise Exception(f"Error al cargar la hoja '{sheet_name}': {e}")

def ajustadorConciliacion(df):
    if 'NUM_DOCUMENTO' not in df.columns or 'PRO_CONSECUTIVO' not in df.columns:
        raise ValueError("El DataFrame no contiene las columnas 'NUM_DOCUMENTO' o 'PRO_CONSECUTIVO'")
    df['ID-SNIES'] = df['NUM_DOCUMENTO'].astype(str) + '-' + df['PRO_CONSECUTIVO'].astype(str)
    columnas = ['ID-SNIES'] + [col for col in df.columns if col != 'ID-SNIES']
    df = df[columnas]
    return df

def ajustadorSq(df):
    ordenColumnas = [
        'Documento', 'Id  factura', 'Tercero', 'Estado Actual', 'Destino', 
        'Nombre de Destino', 'Nombre del Tercero', 'Tipo de Documento', 
        'Fecha', 'Valor Factura', 'Valor Ajuste', 'Valor Pagado', 
        'Valor Anulado', 'Saldo', 'Id Integracion', 'Periodico Academico', 
        'Tipo de Financiacion'
    ]
    columnasFaltantes = [col for col in ordenColumnas if col not in df.columns]
    if columnasFaltantes:
        raise ValueError(f"Las siguientes columnas están ausentes en el DataFrame: {columnasFaltantes}")
    df = df[ordenColumnas]
    return df

# Extraccion
conciliacionMen = cargarArchivosDataframes(file_path_men261224,'plantilla_gratuidad_ies')
sq = cargarArchivosDataframes(file_path_sq131224,'sq')

# Manipulación
## DataFrame  Conciliacion
dfConciliacionMen = ajustadorConciliacion(conciliacionMen)
df_sq_ajustado = ajustadorSq(sq)


# Carga
with pd.ExcelWriter(output_pathXlsx, engine='xlsxwriter') as writer:
    if dfConciliacionMen is not None:
        dfConciliacionMen.to_excel(writer, sheet_name='conciliacion2024_2', index=False)
    if df_sq_ajustado is not None:
        df_sq_ajustado.to_excel(writer, sheet_name='Sq2024_2', index=False)

print("Los resultados han sido guardados en el documento y archivo Excel.")