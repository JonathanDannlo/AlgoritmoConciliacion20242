import os
import math
import pandas as pd
import numpy as np
from openpyxl import load_workbook

pathPiamAll = '/content/PIAM_UNICAUCA6.xlsx'
pathMov251 = '/content/Movilidad25_1.xlsx'
outPathX = '/content/ReporteMovilidad25_1.xlsx'

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

def consolidarEstadosBeneficio(dfEstudiantes, df1, df2, df3, df4, df5, df6):
    dfEstudiantes = dfEstudiantes.rename(columns={"Movilidad": "ID_ESTUDIANTE"})
    df1 = df1.rename(columns={"ID": "ID_ESTUDIANTE", "ESTADO POLITICA": "ESTADO_1", "PERIODOS_APROBADOS_FSE": "INFO_EXTRA_1"})
    df2 = df2.rename(columns={"ID": "ID_ESTUDIANTE", "ESTADO": "ESTADO_2", "PERIODOS_APROBADOS_FSE": "INFO_EXTRA_2"})
    df3 = df3.rename(columns={"IDENTIFICACION": "ID_ESTUDIANTE", "ESTADO": "ESTADO_3", "FONDO": "INFO_EXTRA_3"})
    df4 = df4.rename(columns={"IDENTIFICACION": "ID_ESTUDIANTE", "ESTADO POLITICA": "ESTADO_4", "APRO": "INFO_EXTRA_4", "PFINANCIADOS": "INFO_EXTRA_4B"})
    df5 = df5.rename(columns={"NUM_DOCUMENTO": "ID_ESTUDIANTE", "ESTADO VALIDADO CD": "ESTADO_5", "Paprobados2": "INFO_EXTRA_5", "Pfinaciados2": "INFO_EXTRA_5B"})
    df6 = df6.rename(columns={"IDENTIFICACION": "ID_ESTUDIANTE", "ESTADO_FINAL": "ESTADO_6", "TOTAL_PERIODOS_APROBADOS": "INFO_EXTRA_6", "PERIODOS_FINANCIADOS": "INFO_EXTRA_6B"})
    dfSalida = dfEstudiantes.copy()
    dfSalida = dfSalida.merge(df1[["ID_ESTUDIANTE", "ESTADO_1", "INFO_EXTRA_1"]], on="ID_ESTUDIANTE", how="left")
    dfSalida = dfSalida.merge(df2[["ID_ESTUDIANTE", "ESTADO_2", "INFO_EXTRA_2"]], on="ID_ESTUDIANTE", how="left")
    dfSalida = dfSalida.merge(df3[["ID_ESTUDIANTE", "ESTADO_3", "INFO_EXTRA_3"]], on="ID_ESTUDIANTE", how="left")
    dfSalida = dfSalida.merge(df4[["ID_ESTUDIANTE", "ESTADO_4", "INFO_EXTRA_4", "INFO_EXTRA_4B"]], on="ID_ESTUDIANTE", how="left")
    dfSalida = dfSalida.merge(df5[["ID_ESTUDIANTE", "ESTADO_5", "INFO_EXTRA_5", "INFO_EXTRA_5B"]], on="ID_ESTUDIANTE", how="left")
    dfSalida = dfSalida.merge(df6[["ID_ESTUDIANTE", "ESTADO_6", "INFO_EXTRA_6", "INFO_EXTRA_6B"]], on="ID_ESTUDIANTE", how="left")
    def determinar_estado_y_info(row):
        info_extra_primaria = None
        info_extra_secundaria = None
        for i in range(6, 0, -1):
            estado = row[f"ESTADO_{i}"]
            if pd.notna(estado):
                if (i in [1, 2, 3, 4, 6] and estado == "Beneficiario") or (i == 5 and estado == "B"):
                    info_extra_primaria = row[f"INFO_EXTRA_{i}"] 
                    if i >= 4:
                        info_extra_secundaria = row[f"INFO_EXTRA_{i}B"]
                return estado, info_extra_primaria, info_extra_secundaria
        return None, None, None
    dfSalida[["Ultimo estado", "Ultimo P Aprobado", "Ultimo P Financiado"]] = dfSalida.apply(
        lambda row: pd.Series(determinar_estado_y_info(row)), axis=1)
    dfSalida = dfSalida.rename(columns={
        "ESTADO_1": "2022-1",
        "ESTADO_2": "2022-2",
        "ESTADO_3": "2023-1",
        "ESTADO_4": "2023-2",
        "ESTADO_5": "2024-1",
        "ESTADO_6": "2024-2"
    })
    columnas_finales = [
        "ID_ESTUDIANTE", "Ultimo estado", "Ultimo P Aprobado", "Ultimo P Financiado",
        "2022-1", "2022-2", "2023-1", "2023-2", "2024-1", "2024-2"
    ]
    return dfSalida[columnas_finales]

def cargadorDF(outputPath, dfs, nombresHojas, engine='xlsxwriter'):
    if not isinstance(dfs, list) or not isinstance(nombresHojas, list):
        raise TypeError("Ambos argumentos, 'dfs' y 'nombresHojas', deben ser listas.")
    if len(dfs) != len(nombresHojas):
        raise ValueError("El número de DataFrames debe coincidir con el número de nombres de hojas.")
    with pd.ExcelWriter(outputPath, engine=engine) as writer:
        for df, hoja in zip(dfs, nombresHojas):
            if df is not None:
                df.to_excel(writer, sheet_name=hoja, index=False)
        print(f"Los resultados se han guardado exitosamente en {outputPath}.")

# Extracion
Piam221 = cargarArchivosDataframes(pathPiamAll,'PIAM2022_1')
Piam222 = cargarArchivosDataframes(pathPiamAll,'PIAM2022_2')
Piam231 = cargarArchivosDataframes(pathPiamAll,'PIAM2023_1')
Piam232 = cargarArchivosDataframes(pathPiamAll,'PIAM2023_2')
Piam241 = cargarArchivosDataframes(pathPiamAll,'PIAM24_1_DF')
Piam242 = cargarArchivosDataframes(pathPiamAll,'PIAM2024_2_DF')
movilidad = cargarArchivosDataframes(pathMov251,'Movilidad25-1')

# Manipulacion
mov251 = consolidarEstadosBeneficio(movilidad,Piam221,Piam222,Piam231,Piam232,Piam241,Piam242)

#Carga
cargadorDF(outPathX,[mov251],['Movilidad251'])
