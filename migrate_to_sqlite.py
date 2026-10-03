import os
import pandas as pd
import logging
from database import init_db, get_db_connection
from config import EXCEL_FAO, EXCEL_FAP, EXCEL_NOMINAS

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def parse_date(date_series):
    # Trata de convertir fechas de pandas a formato YYYY-MM-DD para SQLite
    return pd.to_datetime(date_series, errors='coerce').dt.strftime('%Y-%m-%d')

def migrate_faos():
    if not os.path.exists(EXCEL_FAO):
        logger.warning(f"No existe {EXCEL_FAO}")
        return
        
    try:
        df = pd.read_excel(EXCEL_FAO, skiprows=1)
        if df.empty: return
        
        # Limpiar NaNs
        df = df.fillna('')
        
        with get_db_connection() as conn:
            cursor = conn.cursor()
            count = 0
            for _, row in df.iterrows():
                dni = str(row.get('Numero', '')).strip()
                if not dni: continue
                
                cursor.execute('''
                    INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, num_permiso, local, tarea, fecha_inicio, fecha_fin)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    dni,
                    str(row.get('Nombre', '')),
                    str(row.get('Apellido', '')),
                    'FAO',
                    str(row.get('FAO', '')),
                    str(row.get('Marca', '')),
                    str(row.get('Tarea/s', '')),
                    str(row.get('Fecha Inicio', '')),
                    str(row.get('Fecha Fin', ''))
                ))
                count += 1
            logger.info(f"Migrados {count} registros de FAO.")
    except Exception as e:
        logger.error(f"Error migrando FAOs: {e}")

def migrate_faps():
    if not os.path.exists(EXCEL_FAP):
        logger.warning(f"No existe {EXCEL_FAP}")
        return
        
    try:
        df = pd.read_excel(EXCEL_FAP, skiprows=1)
        if df.empty: return
        df = df.fillna('')
        
        with get_db_connection() as conn:
            cursor = conn.cursor()
            count = 0
            for _, row in df.iterrows():
                dni = str(row.get('Numero', '')).strip()
                if not dni: continue
                
                cursor.execute('''
                    INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, num_permiso, local, fecha_inicio, fecha_fin)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    dni,
                    str(row.get('Nombre', '')),
                    str(row.get('Apellido', '')),
                    'FAP',
                    str(row.get('FAP', '')),
                    str(row.get('Marca', '')),
                    str(row.get('Fecha Inicio', '')),
                    str(row.get('Fecha Fin', ''))
                ))
                count += 1
            logger.info(f"Migrados {count} registros de FAP.")
    except Exception as e:
        logger.error(f"Error migrando FAPs: {e}")

def migrate_nominas():
    if not os.path.exists(EXCEL_NOMINAS):
        logger.warning(f"No existe {EXCEL_NOMINAS}")
        return
        
    try:
        df = pd.read_excel(EXCEL_NOMINAS)
        if df.empty: return
        df = df.fillna('')
        
        with get_db_connection() as conn:
            cursor = conn.cursor()
            count = 0
            for _, row in df.iterrows():
                dni = str(row.get('DNI', '')).strip()
                if not dni: continue
                
                cursor.execute('''
                    INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, local, fecha_inicio, fecha_fin)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    dni,
                    str(row.get('Nombre', '')),
                    str(row.get('Apellido', '')),
                    'NOMINA',
                    str(row.get('Empresa', '')),
                    str(row.get('Vigencia Desde', '')).split(' ')[0], # Truncar a YYYY-MM-DD si trae hora
                    str(row.get('Vigencia Hasta', '')).split(' ')[0]
                ))
                count += 1
            logger.info(f"Migradas {count} personas en Nóminas.")
    except Exception as e:
        logger.error(f"Error migrando Nóminas: {e}")

def run_migration():
    logger.info("Iniciando migración a SQLite...")
    
    # 1. Crear base de datos y tablas
    init_db()
    
    # 2. Limpiar la tabla de autorizaciones antes de migrar (por si se corre varias veces)
    with get_db_connection() as conn:
        conn.execute('DELETE FROM autorizaciones')
        
    # 3. Migrar Excels a SQLite
    migrate_faos()
    migrate_faps()
    migrate_nominas()
    
    logger.info("Migración completada con éxito.")

if __name__ == '__main__':
    run_migration()
