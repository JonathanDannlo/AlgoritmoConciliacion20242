import os
import math
import pandas as pd
import numpy as np
from openpyxl import load_workbook

filePathDarca = '/content/Darca_Conciliacion2024.xlsx'
filePathSq = '/content/SqEnero13_2024_2.xlsx'
filePathMen261224 = '/content/Reporte_general__Caracterizacion__novedades_y_requisitos_politica_de_gratuidad__para_las_IES__26_12_2024_cia.xlsx'
filePathAudotira = '/content/AuditoriaConciliacionPiam20242.xlsx'
filePathPago3 = '/content/PAGO 3 POLITICA DE GRATUIDAD PERIODO 2024-2 OBSERVACIONES.xlsx'
outputPathXlsx = '/content/AuditoriaPiam20242Darca.xlsx'

columnasDfMen = ['ID-SNIES','PERIODO_APROBACION','FONDO_ORIGEN','CRITERIO_NO_ACEPTACION','CRITERIO_NO_RENOVACION','GRADO_PREVIO',
                 'PUNTAJE_SISBEN','ES_VICTIMA','ES_INDIGENA','ESTRATO_INGRESO','CRITERIO_ACEPTACION','TOTAL_PERIODOS_APROBADOS',
                 'PERIODOS_FINANCIADOS','PERIODOS_A_FINANCIAR','RESULTADO_VALIDACION','RESULTAD_APROBACION_RENOVACION','VAL_NETO_DER_MAT_A_CARGO_EST',
                 'VALOR_MATRICULADO','VALOR_A_CUBRIR','ESTADO_GIRO']
matriculaBruta = ['DERECHOS_MATRICULA',
                  'BIBLIOTECA_DEPORTES',
                  'LABORATORIOS',
                  'RECURSOS_COMPUTACIONALES',
                  'SEGURO_ESTUDIANTIL',
                  'VRES_COMPLEMENTARIOS',
                  'RESIDENCIAS',
                  'REPETICIONES']
meritoAcademico = ['CONVENIO_DESCENTRALIZACION',
                   'BECA',
                   'MATRICULA_HONOR',
                   'MEDIA_MATRICULA_HONOR',
                   'TRABAJO_GRADO',
                   'DOS_PROGRAMAS',
                   'DESCUENTO_HERMANO',
                   'ESTIMULO_EMP_DTE_PLANTA',
                   'ESTIMULO_CONYUGE',
                   'EXEN_HIJOS_CONYUGE_CATEDRA',
                   'EXEN_HIJOS_CONYUGE_OCASIONAL',
                   'HIJOS_TRABAJADORES_OFICIALES',
                   'ACTIVIDAES_LUDICAS_DEPOR',
                   'DESCUENTOS',
                   'SERVICIOS_RELIQUIDACION',
                   'DESCUENTO_LEY_1171']
refDuplicados = [10893247, 10893215, 10896951, 10897405, 10897145, 10904986]

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

def auditarRepetidos(df, recibo, id, codigo):
    duplicadosRecibo = df[df[recibo].duplicated(keep=False)]
    totalDuplicadosRecibo = duplicadosRecibo.shape[0]
    duplicadosId = df[df[id].duplicated(keep=False)]
    totalDuplicadosId = duplicadosId.shape[0]
    duplicadosCodigo = duplicadosId[duplicadosId[codigo].duplicated(keep=False)]
    totalDuplicadosCodigo = duplicadosCodigo.shape[0]
    print(f"Total registros duplicados basados en '{recibo}': {totalDuplicadosRecibo}")
    print(f"Total registros duplicados basados en '{id}': {totalDuplicadosId}")
    print(f"De esos, total registros duplicados también por '{codigo}': {totalDuplicadosCodigo}")
    return df

def interpoladorDarcaFacturacionSq(df1, df2):
  df = pd.merge(
      df1,
      df2,
      left_on='RECIBO',
      right_on='Documento',
      how='left')
  df = auditarRepetidos(df,'RECIBO','IDENTIFICACION','CODIGO')
  return df

def integradorDarcaFacturaPago(df1,df2):
  df = pd.merge(
      df1,
      df2[['RECIBO','ID-SNIES','ESTADO_GIRO_df2','ComEjecucionFSE','ValidacionPago',
           'ObservacionEstado','EstadoBeneficio','Pago1','Pago2','Pago3']],
      left_on='RECIBO',
      right_on='RECIBO',
      how='left')
  return df

def ajustadorConciliacion(df):
    if 'NUM_DOCUMENTO' not in df.columns or 'PRO_CONSECUTIVO' not in df.columns:
        raise ValueError("El DataFrame no contiene las columnas 'NUM_DOCUMENTO' o 'PRO_CONSECUTIVO'")
    df['ID-SNIES'] = df['NUM_DOCUMENTO'].astype(str) + '-' + df['PRO_CONSECUTIVO'].astype(str)
    columnas = ['ID-SNIES'] + [col for col in df.columns if col != 'ID-SNIES']
    df = df[columnas]
    return df

def procesarIdSnies(df):
    columnaRequerida = ['ID-SNIES', 'IDENTIFICACION', 'SNIESPROGRAMA']
    columna = [col for col in columnaRequerida if col not in df.columns]
    if columna:
        raise KeyError(f"Las siguientes columnas están ausentes en el DataFrame: {columna}")
    df['ID-SNIESNueva'] = df['ID-SNIES']
    df['ID-SNIESNueva'] = df['ID-SNIESNueva'].fillna(
        df['IDENTIFICACION'].astype(str) + "-" + df['SNIESPROGRAMA'].astype(str)
    )
    df.rename(columns={'ID-SNIES': 'ID-SNIES-EJECUTADO', 'ID-SNIESNueva': 'ID-SNIES'}, inplace=True)
    return df

def interpoladorPiamConciliacion(df1,df2):
  columnasDf2 = ['ID-SNIES','PERIODO_APROBACION','FONDO_ORIGEN','CRITERIO_NO_ACEPTACION','CRITERIO_NO_RENOVACION','GRADO_PREVIO',
                 'PUNTAJE_SISBEN','ES_VICTIMA','ES_INDIGENA','ESTRATO_INGRESO','CRITERIO_ACEPTACION','TOTAL_PERIODOS_APROBADOS',
                 'PERIODOS_FINANCIADOS','PERIODOS_A_FINANCIAR','RESULTADO_VALIDACION','RESULTAD_APROBACION_RENOVACION','VAL_NETO_DER_MAT_A_CARGO_EST',
                 'VALOR_MATRICULADO','VALOR_A_CUBRIR','ESTADO_GIRO']
  df = pd.merge(
      df1,
      df2[columnasDf2],
      on='ID-SNIES',
      how='left',
      suffixes=('_dfa', '_dfb'),
      indicator=True)
  return df

def calcular_matricula(df):
    df['BRUTA'] = df[matriculaBruta].sum(axis=1)
    df['BRUTAORD'] = df['BRUTA'] - df['SEGURO_ESTUDIANTIL']
    df['NETAORD'] = df['BRUTAORD'] - df['VOTO'].abs()
    df['MERITO'] = df[meritoAcademico].sum(axis=1).abs()
    df['MTRNETA'] = df['BRUTA'] - df['VOTO'].abs() - df['MERITO']
    df['NETAAPL'] = df['MTRNETA'] - df['SEGURO_ESTUDIANTIL']
    return df

def procesadorEstado(df):
    columnasNecesarias = ['ESTADO_GIRO', 'CODIGO', 'EstadoBeneficio', 'ValidacionPago',
                           'Estado Actual', 'Valor Pagado', 'SEGURO_ESTUDIANTIL', 'VALOR_A_CUBRIR',
                           'Saldo', 'NETAAPL', 'MERITO']
    for columna in columnasNecesarias:
        if columna not in df.columns:
            raise KeyError(f"El DataFrame no contiene la columna '{columna}'.")
    if 'ESTADO_FINAL' not in df.columns:
        df['ESTADO_FINAL'] = np.nan
    if 'FONDO_FINAL' not in df.columns:
        df['FONDO_FINAL'] = np.nan
    if 'ESTADO_BENEFICIOFINAL' not in df.columns:
        df['ESTADO_BENEFICIOFINAL'] = np.nan
    if 'Pago4' not in df.columns:
        df['Pago4'] = np.nan
    if 'Reintegro' not in df.columns:
        df['Reintegro'] = np.nan
    df['ESTADO_FINAL'] = df['ESTADO_FINAL'].astype(object)
    df['FONDO_FINAL'] = df['FONDO_FINAL'].astype(object)
    df['ESTADO_BENEFICIOFINAL'] = df['ESTADO_BENEFICIOFINAL'].astype(object)
    df['Reintegro'] = df['Reintegro'].astype(object)
    condicionesBeneficiario = df['ESTADO_GIRO'].isin(['Aprobado con giro', 'Renovado con giro'])
    condicionesExcluido = df['ESTADO_GIRO'].isin(['Aprobado valor cero', 'Renovado valor cero', 'No aprobado', 'No renovado'])
    condicionesPotencialExcluido = df['ESTADO_GIRO'].isnull() | (df['ESTADO_GIRO'] == '')
    df.loc[condicionesBeneficiario, ['ESTADO_FINAL', 'FONDO_FINAL']] = ['Beneficiario', 'FSE']
    df.loc[condicionesExcluido, ['ESTADO_FINAL', 'FONDO_FINAL']] = ['Excluido', 'Estudiante']
    df.loc[condicionesPotencialExcluido, ['ESTADO_FINAL', 'FONDO_FINAL']] = ['Potencial Excluido', 'Estudiante']
    duplicados = df.duplicated(subset='CODIGO', keep=False)
    df.loc[duplicados, ['ESTADO_FINAL', 'FONDO_FINAL']] = ['Duplicado', 'Darca']
    condicionBajaBeneficio = df['EstadoBeneficio'] == 'BAJA BENEFICIO'
    df.loc[condicionBajaBeneficio, ['Pago1', 'Pago2', 'Pago3']] = 0
    condicionPagoTotal = (df['ESTADO_FINAL'] == 'Beneficiario') & (df['ValidacionPago'] == 1)
    df.loc[condicionPagoTotal, 'ESTADO_BENEFICIOFINAL'] = 'PAGO FSE TOTAL'
    condicionPagoFse = (
        (df['ESTADO_FINAL'] == 'Beneficiario') &
        (df['ESTADO_BENEFICIOFINAL'].isna()) &
        (df['Estado Actual'] == 'ac') &
        (df['Valor Pagado'] == df['SEGURO_ESTUDIANTIL']) &
        (df['Saldo'] == df['VALOR_A_CUBRIR'])
    )
    df.loc[condicionPagoFse, 'Pago4'] = df['VALOR_A_CUBRIR']
    df.loc[condicionPagoFse, 'ESTADO_BENEFICIOFINAL'] = 'PAGO FSE TOTAL'
    condicionPagoFse1 = (
        (df['ESTADO_FINAL'] == 'Beneficiario') &
        (df['ESTADO_BENEFICIOFINAL'].isna()) &
        (df['Estado Actual'] == 'ac') &
        (df['Valor Pagado'] == df['SEGURO_ESTUDIANTIL']) &
        (df['Saldo'] == df['NETAAPL']) &
        (df['NETAAPL'] == (df['VALOR_A_CUBRIR'] - df['MERITO']))
    )
    df.loc[condicionPagoFse1, 'Pago4'] = df['Saldo']
    df.loc[condicionPagoFse1, 'ESTADO_BENEFICIOFINAL'] = 'PAGO FSE TOTAL'
    condicionPagoFse2 = (
        (df['ESTADO_FINAL'] == 'Beneficiario') &
        (df['ESTADO_BENEFICIOFINAL'].isna()) &
        (df['Estado Actual'] == 'ac') &
        (df['NETAAPL'] == df['VALOR_A_CUBRIR']) &
        (df['Valor Pagado'] != df['SEGURO_ESTUDIANTIL'])
    )
    df.loc[condicionPagoFse2, 'Pago4'] = df['Saldo']
    df.loc[condicionPagoFse2, 'Reintegro'] = (df['Valor Pagado']-df['SEGURO_ESTUDIANTIL'])
    df.loc[condicionPagoFse2, 'ESTADO_BENEFICIOFINAL'] = 'PAGO | REINTEGRO FSE'
    condicionPagoFse3 = (
        (df['ESTADO_FINAL'] == 'Beneficiario') &
        (df['ESTADO_BENEFICIOFINAL'].isna()) &
        (df['Estado Actual'] == 'ac') &
        (df['NETAAPL'] != df['VALOR_A_CUBRIR']) &
        (df['Valor Pagado'] != df['SEGURO_ESTUDIANTIL'])
    )
    df.loc[condicionPagoFse3, 'Pago4'] = df['Saldo']
    df.loc[condicionPagoFse3, 'Reintegro'] = (df['Valor Pagado']-df['SEGURO_ESTUDIANTIL'])
    df.loc[condicionPagoFse3, 'ESTADO_BENEFICIOFINAL'] = 'PAGO | REINTEGRO FSE'
    condicionPagoFse4 = (
        (df['ESTADO_FINAL'] == 'Beneficiario') &
        (df['ESTADO_BENEFICIOFINAL'].isna()) &
        (df['Estado Actual'] == 'ca') &
        (df['Saldo'] == 0)
    )
    df.loc[condicionPagoFse4, 'Reintegro'] = df['NETAAPL']
    df.loc[condicionPagoFse4, 'ESTADO_BENEFICIOFINAL'] = 'REINTEGRO FSE'
    return df

def eliminadorRegistros(df, refDu):
  df = df[~df['Documento'].isin(refDu)]
  df.loc[df['Documento'] == 10885015, ['Pago4', 'ESTADO_FINAL', 'FONDO_FINAL','ESTADO_BENEFICIOFINAL']] = [143000, 'Beneficiario', 'FSE','PAGO FSE TOTAL']
  df.loc[df['ESTADO_FINAL'] == 'Duplicado', ['ESTADO_FINAL', 'FONDO_FINAL']] = ['Excluido', 'Estudiante']
  return df

def validadorFSE(df):
    required_columns = ['Pago1', 'Pago2', 'Pago3', 'Pago4', 'ESTADO_FINAL']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"La columna requerida '{col}' no está presente en el DataFrame.")
    for col in ['Pago1', 'Pago2', 'Pago3', 'Pago4','MERITO']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    df['Pago FSE'] = df.apply(
        lambda row: row['Pago1'] + row['Pago2'] + row['Pago3'] + row['Pago4']
        if row['ESTADO_FINAL'] == 'Beneficiario' else 0, axis=1)
    df['MeritoFSE'] = df.apply(
        lambda row: row['MERITO'] if row['ESTADO_FINAL'] == 'Beneficiario' else None, axis=1)
    return df

def generadorMarcaje(df):
    columnasFiltradas = ['CODIGO', 'ESTADO_FINAL', 'FONDO_FINAL', 'TOTAL_PERIODOS_APROBADOS', 'PERIODOS_FINANCIADOS']
    dFiltrado = df[columnasFiltradas].copy()
    dFiltrado['CODIGO'] = dFiltrado['CODIGO'].apply(lambda x: f"{x:.0f}")
    dFiltrado['ESTADO_FINAL'] = dFiltrado['ESTADO_FINAL'].replace({
        'Beneficiario': 'B',
        'Excluido': 'E',
        'Potencial Excluido': 'PE'
    })
    dFiltrado.loc[dFiltrado['ESTADO_FINAL'] == 'PE', ['TOTAL_PERIODOS_APROBADOS', 'PERIODOS_FINANCIADOS']] = 0
    dFiltrado.loc[dFiltrado['ESTADO_FINAL'] == 'E', ['TOTAL_PERIODOS_APROBADOS', 'PERIODOS_FINANCIADOS']] = 0
    condicion = (dFiltrado['ESTADO_FINAL'] == 'B') & (dFiltrado['TOTAL_PERIODOS_APROBADOS'] - dFiltrado['PERIODOS_FINANCIADOS'] == 0)
    dFiltrado.loc[condicion, ['TOTAL_PERIODOS_APROBADOS', 'PERIODOS_FINANCIADOS']] = 0
    dFiltrado.loc[condicion, 'ESTADO_FINAL'] = 'E'
    dFiltrado.loc[condicion, 'FONDO_FINAL'] = 'Estudiante'
    return dFiltrado

def actualizarPago3(df1, df2):
    dfa = pd.merge(df1,
                   df2[['ID FACTURA', 'APLICADO','SALDO A FAVOR']],
                   left_on='Id  factura',
                   right_on='ID FACTURA',
                   how='left')
    dfa['Pago3'] = dfa['APLICADO'].combine_first(dfa['Pago3'])
    dfa['Reintegro'] = dfa['SALDO A FAVOR'].combine_first(dfa['Reintegro'])
    dfa.loc[dfa['APLICADO'].notna(), 'ESTADO_BENEFICIOFINAL'] = dfa['APLICADO'].apply(
        lambda x: "PAGO | REINTEGRO" if x != 0 else "REINTEGRO"
    )
    dfa.loc[dfa['APLICADO'].notna(), 'Pago FSE'] = dfa['Pago3']
    dfa = dfa.drop(columns=['ID FACTURA', 'APLICADO','SALDO A FAVOR'])
    return dfa

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




#Extracion
darca2024 = cargarArchivosDataframes(filePathDarca,'DARCA20242')
sq20242 = cargarArchivosDataframes(filePathSq,'sq')
auditoria = cargarArchivosDataframes(filePathAudotira,'Piam242ECSP')
men20242 = cargarArchivosDataframes(filePathMen261224,'plantilla_gratuidad_ies')
dfPago3 = cargarArchivosDataframes(filePathPago3,'PARCIALREINTEGRO')

#Manipulacion
interpoladorDarcaFacturacionSq = interpoladorDarcaFacturacionSq(darca2024,sq20242)
integradorDarcaFacturaPago = integradorDarcaFacturaPago(interpoladorDarcaFacturacionSq,auditoria)
procesadorIdSnies = procesarIdSnies(integradorDarcaFacturaPago)
men = ajustadorConciliacion(men20242)
interpoladorPiamConciliacion = interpoladorPiamConciliacion(procesadorIdSnies,men)
piam = calcular_matricula(interpoladorPiamConciliacion)
piam20242 = procesadorEstado(piam)
piam20242f = eliminadorRegistros(piam20242,refDuplicados)
piam20242fi = validadorFSE(piam20242f)
piam20242fii = actualizarPago3(piam20242fi,dfPago3)
piamMarcaje = generadorMarcaje(piam20242fii)

#Carga
cargadorDF(outputPathXlsx,[piam20242fii,piamMarcaje],['piam20242fii','piamMarcaje'])

print("Los resultados han sido guardados en el documento y archivo Excel.")