"""
Módulo de Utilidades para Control de Acceso AOS

Funciones auxiliares para:
- Validación de archivos Excel importados
- Sistema de backup automático
- Validación de datos
"""

import os
import shutil
from datetime import datetime
import pandas as pd
from logger_config import get_logger
from config import BASE_DIR

logger = get_logger(__name__)

# Directorio de backups
BACKUPS_DIR = os.path.join(BASE_DIR, 'backups')
os.makedirs(BACKUPS_DIR, exist_ok=True)


def validar_archivo_fap(filepath):
    """
    Valida que el archivo FAP tenga el formato esperado.
    
    Args:
        filepath (str): Ruta al archivo Excel FAP
        
    Returns:
        tuple: (bool, str) - (es_válido, mensaje)
    """
    try:
        # Leer archivo con header en fila 1 (índice 1)
        df = pd.read_excel(filepath, header=1)
        
        # Columnas requeridas según config.py
        columnas_requeridas = ['Numero', 'Nombre', 'Apellido', 'FAP', 'Fecha Fin', 'Marca']
        
        # Verificar que existan todas las columnas
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
        
        if columnas_faltantes:
            mensaje = f"Columnas faltantes en FAP: {', '.join(columnas_faltantes)}"
            logger.error(mensaje)
            return False, mensaje
        
        # Verificar que haya al menos una fila de datos
        if len(df) == 0:
            mensaje = "El archivo FAP está vacío (no tiene datos)"
            logger.warning(mensaje)
            return False, mensaje
        
        # Verificar que la columna DNI tenga datos válidos
        df['Numero'] = df['Numero'].astype(str)
        dnis_validos = df['Numero'].str.match(r'^\d{7,8}$').sum()
        
        if dnis_validos == 0:
            mensaje = "El archivo FAP no contiene DNIs válidos"
            logger.error(mensaje)
            return False, mensaje
        
        logger.info(f"Archivo FAP validado correctamente: {len(df)} registros, {dnis_validos} DNIs válidos")
        return True, f"Archivo válido: {len(df)} registros"
        
    except Exception as e:
        mensaje = f"Error al validar archivo FAP: {str(e)}"
        logger.error(mensaje, exc_info=True)
        return False, mensaje


def validar_archivo_fao(filepath):
    """
    Valida que el archivo FAO tenga el formato esperado.
    
    Args:
        filepath (str): Ruta al archivo Excel FAO
        
    Returns:
        tuple: (bool, str) - (es_válido, mensaje)
    """
    try:
        # Leer archivo con header en fila 1 (índice 1)
        df = pd.read_excel(filepath, header=1)
        
        # Columnas requeridas según config.py
        columnas_requeridas = ['Numero', 'Nombre', 'Apellido', 'FAO', 'Fecha Fin', 'Marca', 'Tarea']
        
        # Verificar que existan todas las columnas
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
        
        if columnas_faltantes:
            mensaje = f"Columnas faltantes en FAO: {', '.join(columnas_faltantes)}"
            logger.error(mensaje)
            return False, mensaje
        
        # Verificar que haya al menos una fila de datos
        if len(df) == 0:
            mensaje = "El archivo FAO está vacío (no tiene datos)"
            logger.warning(mensaje)
            return False, mensaje
        
        # Verificar que la columna DNI tenga datos válidos
        df['Numero'] = df['Numero'].astype(str)
        dnis_validos = df['Numero'].str.match(r'^\d{7,8}$').sum()
        
        if dnis_validos == 0:
            mensaje = "El archivo FAO no contiene DNIs válidos"
            logger.error(mensaje)
            return False, mensaje
        
        logger.info(f"Archivo FAO validado correctamente: {len(df)} registros, {dnis_validos} DNIs válidos")
        return True, f"Archivo válido: {len(df)} registros"
        
    except Exception as e:
        mensaje = f"Error al validar archivo FAO: {str(e)}"
        logger.error(mensaje, exc_info=True)
        return False, mensaje


def crear_backup(filepath, tipo='manual'):
    """
    Crea un backup de un archivo Excel con timestamp.
    
    Args:
        filepath (str): Ruta al archivo a respaldar
        tipo (str): Tipo de backup ('manual', 'auto', 'pre-update')
        
    Returns:
        tuple: (bool, str) - (éxito, ruta_backup o mensaje_error)
    """
    try:
        if not os.path.exists(filepath):
            mensaje = f"Archivo no existe: {filepath}"
            logger.warning(mensaje)
            return False, mensaje
        
        # Obtener nombre base del archivo
        filename = os.path.basename(filepath)
        nombre_sin_ext, ext = os.path.splitext(filename)
        
        # Crear timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Crear subdirectorio por tipo de backup
        backup_subdir = os.path.join(BACKUPS_DIR, tipo)
        os.makedirs(backup_subdir, exist_ok=True)
        
        # Nombre del backup
        backup_filename = f"{nombre_sin_ext}_{timestamp}{ext}"
        backup_path = os.path.join(backup_subdir, backup_filename)
        
        # Copiar archivo
        shutil.copy2(filepath, backup_path)
        
        logger.info(f"Backup creado: {backup_filename} ({tipo})")
        return True, backup_path
        
    except Exception as e:
        mensaje = f"Error al crear backup: {str(e)}"
        logger.error(mensaje, exc_info=True)
        return False, mensaje


def limpiar_backups_antiguos(dias=30, tipo='auto'):
    """
    Elimina backups automáticos más antiguos que X días.
    
    Args:
        dias (int): Días de antigüedad para eliminar
        tipo (str): Tipo de backup a limpiar
        
    Returns:
        int: Cantidad de archivos eliminados
    """
    try:
        backup_subdir = os.path.join(BACKUPS_DIR, tipo)
        
        if not os.path.exists(backup_subdir):
            return 0
        
        ahora = datetime.now()
        eliminados = 0
        
        for filename in os.listdir(backup_subdir):
            filepath = os.path.join(backup_subdir, filename)
            
            # Verificar que sea archivo
            if not os.path.isfile(filepath):
                continue
            
            # Obtener fecha de modificación
            mod_time = datetime.fromtimestamp(os.path.getmtime(filepath))
            
            # Calcular antigüedad
            antiguedad = (ahora - mod_time).days
            
            if antiguedad > dias:
                os.remove(filepath)
                eliminados += 1
                logger.debug(f"Backup eliminado (antigüedad {antiguedad} días): {filename}")
        
        if eliminados > 0:
            logger.info(f"Limpieza de backups: {eliminados} archivos eliminados (>{dias} días)")
        
        return eliminados
        
    except Exception as e:
        logger.error(f"Error al limpiar backups: {str(e)}", exc_info=True)
        return 0


def validar_dni(dni_str):
    """
    Valida formato de DNI.
    
    Args:
        dni_str (str): DNI a validar
        
    Returns:
        bool: True si es válido
    """
    import re
    
    if not dni_str:
        return False
    
    # Limpiar posibles espacios o puntos
    dni_limpio = str(dni_str).strip().replace('.', '')
    
    # Verificar longitud (7 u 8 dígitos) y que sean estrictamente numéricos
    return bool(re.match(r'^\d{7,8}$', dni_limpio))


def validar_fecha(fecha_str):
    """
    Valida formato de fecha.
    
    Args:
        fecha_str (str): Fecha a validar
        
    Returns:
        tuple: (bool, datetime o None)
    """
    if not fecha_str:
        return False, None
    
    formatos = ['%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']
    
    for formato in formatos:
        try:
            fecha = datetime.strptime(str(fecha_str), formato)
            return True, fecha
        except ValueError:
            continue
    
    return False, None


def obtener_estadisticas_backups():
    """
    Obtiene estadísticas de los backups existentes.
    
    Returns:
        dict: Estadísticas por tipo de backup
    """
    stats = {}
    
    try:
        if not os.path.exists(BACKUPS_DIR):
            return stats
        
        for tipo in os.listdir(BACKUPS_DIR):
            tipo_path = os.path.join(BACKUPS_DIR, tipo)
            
            if not os.path.isdir(tipo_path):
                continue
            
            archivos = [f for f in os.listdir(tipo_path) if os.path.isfile(os.path.join(tipo_path, f))]
            
            if archivos:
                tamaño_total = sum(os.path.getsize(os.path.join(tipo_path, f)) for f in archivos)
                
                stats[tipo] = {
                    'cantidad': len(archivos),
                    'tamaño_mb': round(tamaño_total / (1024 * 1024), 2),
                    'ultimo': max(archivos, key=lambda f: os.path.getmtime(os.path.join(tipo_path, f)))
                }
        
        return stats
        
    except Exception as e:
        logger.error(f"Error al obtener estadísticas de backups: {str(e)}")
        return stats
