import os
import pandas as pd
from datetime import datetime

from config import (
    EXCEL_FAP, EXCEL_FAO, EXCEL_EXCEPCIONES, EXCEL_NOMINAS,
    COL_DNI, COL_NOMBRE_APELLIDO, COL_NUM_PERMISO, COL_VENCE, COL_LOCAL, COL_TAREA, COL_TIPO_PERMISO,
    COL_DNI_FAP_ORIGINAL, COL_NOMBRE_FAP_ORIGINAL, COL_APELLIDO_FAP_ORIGINAL, COL_NUM_PERMISO_FAP_ORIGINAL, COL_VENCE_FAP_ORIGINAL, COL_LOCAL_FAP_ORIGINAL,
    COL_DNI_FAO_ORIGINAL, COL_NOMBRE_FAO_ORIGINAL, COL_APELLIDO_FAO_ORIGINAL, COL_NUM_PERMISO_FAO_ORIGINAL, COL_VENCE_FAO_ORIGINAL, COL_LOCAL_FAO_ORIGINAL, COL_TAREA_FAO_ORIGINAL
)

# --- VARIABLES GLOBALES PARA DATAFRAMES Y TIEMPOS DE MODIFICACIÓN ---
df_fap, df_fao, df_excepciones, df_nominas = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
ult_mod_fap, ult_mod_fao, ult_mod_excepciones, ult_mod_nominas = 0, 0, 0, 0

# --- FUNCIÓN AUXILIAR PARA LIMPIAR DNI/CUIL ---
def extraer_dni_de_cuil(valor):
    valor_str = str(valor).strip().replace('-', '')
    if len(valor_str) == 11 and valor_str.isdigit():
        return valor_str[2:10]
    elif len(valor_str) >= 7 and valor_str.isdigit():
        return valor_str
    else:
        return valor_str

def cargar_y_procesar_excel(archivo_excel, ultima_modificacion, tipo_permiso, mapa_columnas, df_actual, header=0, dni_col_original=None):
    try:
        if not os.path.exists(archivo_excel):
            return pd.DataFrame(), 0
        mod_time = os.path.getmtime(archivo_excel)
        if mod_time == ultima_modificacion and not df_actual.empty:
            return df_actual, ultima_modificacion
        print(f"INFO: Detectado cambio en '{os.path.basename(archivo_excel)}'. Recargando...")
        dtype_map = {}
        if dni_col_original is not None:
            dtype_map[dni_col_original] = str
        df = pd.read_excel(archivo_excel, header=header, dtype=dtype_map)
        df.columns = [str(c).strip() for c in df.columns]
        df.rename(columns=mapa_columnas, inplace=True)
        if COL_DNI in df.columns:
            df[COL_DNI] = df[COL_DNI].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'[\.\s-]', '', regex=True).str.strip()
            df.loc[df[COL_DNI].str.lower() == 'nan', COL_DNI] = ''
            df = df[df[COL_DNI] != ''].copy()
            df[COL_DNI] = df[COL_DNI].apply(extraer_dni_de_cuil)
        if 'Nombre' in df.columns and 'Apellido' in df.columns and COL_NOMBRE_APELLIDO not in df.columns:
            df[COL_NOMBRE_APELLIDO] = df['Nombre'].fillna('') + ' ' + df['Apellido'].fillna('')
        df[COL_TIPO_PERMISO] = tipo_permiso
        print(f"-> Recargado y procesado: {os.path.basename(archivo_excel)}")
        return df, mod_time
    except FileNotFoundError:
        print(f"ADVERTENCIA: No se encontró el archivo {archivo_excel}. Se continuará sin él.")
        return pd.DataFrame(), 0
    except Exception as e:
        print(f"Error Crítico al procesar {archivo_excel}: {e}")
        return df_actual, ultima_modificacion

def cargar_autorizaciones():
    global df_fap, ult_mod_fap, df_fao, ult_mod_fao, df_excepciones, ult_mod_excepciones, df_nominas, ult_mod_nominas
    
    mapa_cols_fap = {COL_DNI_FAP_ORIGINAL: COL_DNI, COL_NOMBRE_FAP_ORIGINAL: 'Nombre', COL_APELLIDO_FAP_ORIGINAL: 'Apellido', COL_NUM_PERMISO_FAP_ORIGINAL: COL_NUM_PERMISO, COL_VENCE_FAP_ORIGINAL: COL_VENCE, COL_LOCAL_FAP_ORIGINAL: COL_LOCAL}
    df_fap, ult_mod_fap = cargar_y_procesar_excel(EXCEL_FAP, ult_mod_fap, 'FAP', mapa_cols_fap, df_fap, header=1, dni_col_original=COL_DNI_FAP_ORIGINAL)
    if not df_fap.empty and COL_VENCE in df_fap.columns:
        df_fap[COL_VENCE] = pd.to_datetime(df_fap[COL_VENCE], format='%d/%m/%Y', errors='coerce')

    mapa_cols_fao = {COL_DNI_FAO_ORIGINAL: COL_DNI, COL_NOMBRE_FAO_ORIGINAL: 'Nombre', COL_APELLIDO_FAO_ORIGINAL: 'Apellido', COL_NUM_PERMISO_FAO_ORIGINAL: COL_NUM_PERMISO, COL_VENCE_FAO_ORIGINAL: COL_VENCE, COL_LOCAL_FAO_ORIGINAL: COL_LOCAL, COL_TAREA_FAO_ORIGINAL: COL_TAREA}
    df_fao, ult_mod_fao = cargar_y_procesar_excel(EXCEL_FAO, ult_mod_fao, 'FAO', mapa_cols_fao, df_fao, header=1, dni_col_original=COL_DNI_FAO_ORIGINAL)
    if not df_fao.empty and COL_VENCE in df_fao.columns:
        df_fao[COL_VENCE] = pd.to_datetime(df_fao[COL_VENCE], format='%d/%m/%Y', errors='coerce')
    
    # --- ESTA ES LA CORRECCIÓN ---
    mapa_cols_excepciones = {'Numero': COL_DNI, 'Nombre Completo': COL_NOMBRE_APELLIDO, 'Local': COL_LOCAL, 'Quien Autoriza': 'Quien_Autoriza', 'Fecha de Alta': COL_VENCE}
    df_excepciones, ult_mod_excepciones = cargar_y_procesar_excel(EXCEL_EXCEPCIONES, ult_mod_excepciones, 'Excepcion', mapa_cols_excepciones, df_excepciones, header=0, dni_col_original='Numero')
    # --- FIN DE LA CORRECCIÓN ---
    
    if not df_excepciones.empty and COL_VENCE in df_excepciones.columns:
        df_excepciones[COL_VENCE] = pd.to_datetime(df_excepciones[COL_VENCE], errors='coerce')

    mapa_cols_nominas = {'DNI': COL_DNI, 'Apellido': 'Apellido', 'Nombre': 'Nombre', 'Empresa': COL_LOCAL, 'Vigencia Desde': 'Vigencia Desde', 'Vigencia Hasta': 'Vigencia Hasta'}
    df_nominas, ult_mod_nominas = cargar_y_procesar_excel(EXCEL_NOMINAS, ult_mod_nominas, 'Nomina', mapa_cols_nominas, df_nominas, header=0, dni_col_original='DNI')
    if not df_nominas.empty:
        if 'Nombre' in df_nominas.columns and 'Apellido' in df_nominas.columns:
            df_nominas[COL_NOMBRE_APELLIDO] = df_nominas['Nombre'].fillna('') + ' ' + df_nominas['Apellido'].fillna('')
        if 'Vigencia Desde' in df_nominas.columns:
            df_nominas['Vigencia Desde'] = pd.to_datetime(df_nominas['Vigencia Desde'], errors='coerce', dayfirst=True)
        if 'Vigencia Hasta' in df_nominas.columns:
            df_nominas['Vigencia Hasta'] = pd.to_datetime(df_nominas['Vigencia Hasta'], errors='coerce', dayfirst=True)