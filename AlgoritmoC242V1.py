import os
import math
import pandas as pd
import numpy as np
from openpyxl import load_workbook


file_path_men261224 = '/content/Reporte_general__Caracterizacion__novedades_y_requisitos_politica_de_gratuidad__para_las_IES__26_12_2024_cia.xlsx'
file_path_sq131224 ='/content/SqEnero13_2024_2.xlsx'
file_path_piam ='/content/PlantillaCiaFinalPiam2024_2.xlsx'
output_pathXlsx = '/content/AuditoriaPiam20242Conciliacion.xlsx'
output_pathXlsxPagos = '/content/ReportePagoFSE_2024_2_Etapa3_15012025.xlsx'

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

def interpoladorPiamConciliacion(df1,df2):
  columnasDf1 = ['ID-SNIES', 'codigo', 'Tercero', 'RECIBO','Documento', 'Id  factura', 'BRUTA', 'BRUTAORD', 'NETAORD', 'MERITO',
                 'MTRNETA', 'NETAAPL', 'DERECHOS_MATRICULA', 'BIBLIOTECA_DEPORTES', 'LABORATORIOS', 'RECURSOS_COMPUTACIONALES',
                 'SEGURO_ESTUDIANTIL', 'VRES_COMPLEMENTARIOS', 'RESIDENCIAS', 'REPETICIONES', 'VOTO', 'CONVENIO_DESCENTRALIZACION',
                 'BECA', 'MATRICULA_HONOR', 'MEDIA_MATRICULA_HONOR', 'TRABAJO_GRADO', 'DOS_PROGRAMAS', 'DESCUENTO_HERMANO',
                 'ESTIMULO_EMP_DTE_PLANTA', 'ESTIMULO_CONYUGE', 'EXEN_HIJOS_CONYUGE_CATEDRA', 'HIJOS_TRABAJADORES_OFICIALES',
                 'ACTIVIDAES_LUDICAS_DEPOR', 'DESCUENTOS', 'SERVICIOS_RELIQUIDACION', 'DESCUENTO_LEY_1171', 'PROGRAMA', 'TELEFONO',
                 'CELULAR', 'EMAILINSTITUCIONAL', 'Relación de Giro', 'VALOR DEL GIRO ICETEX', 'VALOR PAGO FACTURA APLICADO',
                 'SALDO A FAVOR', 'MERITO UNICAUCA', 'Sublínea Crédito','Estado Beneficio 1','ESTADO BENEFICIO','FONDO','FSE apl',
                 'Merito 1','Pago1','Pago2','ComEjecucionFSE','Saldo_Pago2','ESTADO_GIRO','TOTAL_PERIODOS_APROBADOS','PERIODOS_FINANCIADOS',
                 'PERIODOS_A_FINANCIAR']
  columnasDf2 = ['ID-SNIES','PERIODO_APROBACION','FONDO_ORIGEN','CRITERIO_NO_ACEPTACION','CRITERIO_NO_RENOVACION','GRADO_PREVIO',
                 'PUNTAJE_SISBEN','ES_VICTIMA','ES_INDIGENA','ESTRATO_INGRESO','CRITERIO_ACEPTACION','TOTAL_PERIODOS_APROBADOS',
                 'PERIODOS_FINANCIADOS','PERIODOS_A_FINANCIAR','RESULTADO_VALIDACION','RESULTAD_APROBACION_RENOVACION','VAL_NETO_DER_MAT_A_CARGO_EST',
                 'VALOR_MATRICULADO','VALOR_A_CUBRIR','ESTADO_GIRO']
  orden_columnas = [
        'ID-SNIES', 'codigo', 'Tercero', 'RECIBO', 'Documento', 'Id  factura', 'BRUTA', 'BRUTAORD', 'NETAORD',
        'MERITO', 'MTRNETA', 'NETAAPL', 'DERECHOS_MATRICULA', 'BIBLIOTECA_DEPORTES', 'LABORATORIOS',
        'RECURSOS_COMPUTACIONALES', 'SEGURO_ESTUDIANTIL', 'VRES_COMPLEMENTARIOS', 'RESIDENCIAS',
        'REPETICIONES', 'VOTO', 'CONVENIO_DESCENTRALIZACION', 'BECA', 'MATRICULA_HONOR', 'MEDIA_MATRICULA_HONOR',
        'TRABAJO_GRADO', 'DOS_PROGRAMAS', 'DESCUENTO_HERMANO', 'ESTIMULO_EMP_DTE_PLANTA', 'ESTIMULO_CONYUGE',
        'EXEN_HIJOS_CONYUGE_CATEDRA', 'HIJOS_TRABAJADORES_OFICIALES', 'ACTIVIDAES_LUDICAS_DEPOR', 'DESCUENTOS',
        'SERVICIOS_RELIQUIDACION', 'DESCUENTO_LEY_1171', 'PROGRAMA', 'TELEFONO', 'CELULAR', 'EMAILINSTITUCIONAL',
        'Relación de Giro', 'VALOR DEL GIRO ICETEX', 'VALOR PAGO FACTURA APLICADO', 'SALDO A FAVOR', 'MERITO UNICAUCA',
        'Sublínea Crédito', 'PERIODO_APROBACION', 'FONDO_ORIGEN', 'CRITERIO_NO_ACEPTACION', 'CRITERIO_NO_RENOVACION',
        'GRADO_PREVIO', 'PUNTAJE_SISBEN', 'ES_VICTIMA', 'ES_INDIGENA', 'ESTRATO_INGRESO', 'CRITERIO_ACEPTACION',
        'RESULTADO_VALIDACION', 'RESULTAD_APROBACION_RENOVACION', 'VAL_NETO_DER_MAT_A_CARGO_EST', 'VALOR_MATRICULADO',
        'VALOR_A_CUBRIR', '_merge', 'ObservacionEstado', 'ESTADO_GIRO_df2', 'ESTADO_GIRO_df1',
        'TOTAL_PERIODOS_APROBADOS_df2', 'PERIODOS_FINANCIADOS_df2', 'PERIODOS_A_FINANCIAR_df2',
        'TOTAL_PERIODOS_APROBADOS_df1', 'PERIODOS_FINANCIADOS_df1', 'PERIODOS_A_FINANCIAR_df1', 'Estado Beneficio 1',
        'ESTADO BENEFICIO', 'FONDO', 'FSE apl', 'Merito 1', 'Pago1', 'Pago2', 'ComEjecucionFSE', 'Saldo_Pago2']
  df = pd.merge(
      df1[columnasDf1],
      df2[columnasDf2],
      on='ID-SNIES',
      how='left',
      suffixes=('_df1', '_df2'),
      indicator=True)
  df['ObservacionEstado'] = np.where(
        df['ESTADO_GIRO_df1'] == df['ESTADO_GIRO_df2'],
        "Estado igual",
        'Cambio estado: ' + df['ESTADO_GIRO_df2'].fillna('') + ' - ' + df['ESTADO_GIRO_df1'].fillna('')
    )
  df = df[orden_columnas]
  return df

def interpoladorPiamConciliacionFacturacionSq(df1, df2):
  df = pd.merge(
      df1,
      df2,
      left_on='RECIBO',
      right_on='Documento',
      how='left')
  return df

def tramitadorPagos(df):
  condicion = (
        (df['ObservacionEstado'] == "Estado igual") &
        (df['Estado Actual'] == "ac") &
        (df['ComEjecucionFSE'].isin(["Pago parcial", "Sin evaluación"])) &
        (df['ESTADO_GIRO_df2'].isin(["Renovado con giro", "Aprobado con giro"]))
    )
  registros_condicion = df[condicion]
  print(f"Registros que cumplen la condición: {registros_condicion.shape[0]}")
  df['Pago1'].fillna(0, inplace=True)
  df['Pago2'].fillna(0, inplace=True)
  df['Pago3'] = 0
  df.loc[condicion, 'Pago3'] = df.loc[condicion].apply(
        lambda row: row['NETAAPL'] - row['Pago1'] - row['Pago2']
        if row['Valor Factura'] == row['MTRNETA'] else 0,
        axis=1)
  df['ValidacionPago'] = df.apply(
        lambda row: (row['Pago1'] + row['Pago2'] + row['Pago3'] == row['NETAAPL'])
        if (row['Pago1'] != 0 or row['Pago2'] != 0 or row['Pago3'] != 0) else None,
        axis=1)
  df['EstadoBeneficio'] = df.apply(
        lambda row: "BAJA BENEFICIO"
        if row['ValidacionPago'] == True and row['ESTADO_GIRO_df2'] == "No aprobado"
        and row['ComEjecucionFSE'] == "Pago total" and row['Pago1'] != 0
        else None,
        axis=1)
  return df

def generadorPago3(df):
    dFiltrado = df[df['Pago3'] != 0]
    dfSalida = dFiltrado[['Tercero_x', 'RECIBO', 'Id  factura_x', 'MTRNETA', 'Pago3']].copy()
    dfSalida.columns = [
        'ID TERCERO',
        'NUMERO RECIBO',
        'ID FACTURA',
        'VALOR FACTURA',
        'VALOR A CANCELAR APROBADO POR EL FSE'
    ]
    dfSalida['NUMERO DE LA CUOTA AFECTAR'] = None
    dfSalida['FECHA DE PAGO'] = None
    return dfSalida

# Extraccion
conciliacionMen = cargarArchivosDataframes(file_path_men261224,'plantilla_gratuidad_ies')
sq = cargarArchivosDataframes(file_path_sq131224,'sq')
piam242 = cargarArchivosDataframes(file_path_piam,'PlantillaCiaFinalPiam2024_2')

# Manipulación
## DataFrame  Conciliacion
dfConciliacionMen = ajustadorConciliacion(conciliacionMen)
df_sq_ajustado = ajustadorSq(sq)

## DataFrame Piam ejecutado vs conciliacion
piamEC = interpoladorPiamConciliacion(piam242,dfConciliacionMen)
piamECS = interpoladorPiamConciliacionFacturacionSq(piamEC,df_sq_ajustado)

### Dataframe Pago3
piamECSP = tramitadorPagos(piamECS)
print(piamECSP.columns)
dfPago3 = generadorPago3(piamECSP)

# Carga
with pd.ExcelWriter(output_pathXlsx, engine='xlsxwriter') as writer:
    """if piam242 is not None:
        piam242.to_excel(writer, sheet_name='Piam242', index=False)
    if dfConciliacionMen is not None:
        dfConciliacionMen.to_excel(writer, sheet_name='conciliacion2024_2', index=False)
    if df_sq_ajustado is not None:
        df_sq_ajustado.to_excel(writer, sheet_name='Sq2024_2', index=False)
    if piamEC is not None:
        piamEC.to_excel(writer, sheet_name='Piam242EC', index=False)
    if piamECS is not None:
        piamECS.to_excel(writer, sheet_name='Piam242ECS', index=False)
    if piamECSP is not None:
        piamECSP.to_excel(writer, sheet_name='Piam242ECSP', index=False)"""
    if dfPago3 is not None:
        dfPago3.to_excel(writer, sheet_name='Pago3', index=False)

print("Los resultados han sido guardados en el documento y archivo Excel.")