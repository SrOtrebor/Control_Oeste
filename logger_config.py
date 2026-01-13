"""
Sistema de Logging Centralizado para Control de Acceso AOS

Este módulo configura el sistema de logging para toda la aplicación.
Características:
- Logs rotativos por tamaño (máximo 10MB por archivo)
- Mantiene hasta 5 archivos de backup
- Diferentes niveles: DEBUG, INFO, WARNING, ERROR, CRITICAL
- Formato consistente con timestamps
- Logs separados por módulo
- Salida a archivo y consola
"""

import logging
import os
from logging.handlers import RotatingFileHandler
from datetime import datetime

# Directorio de logs
LOGS_DIR = os.path.join(os.path.dirname(__file__), 'logs')
os.makedirs(LOGS_DIR, exist_ok=True)

# Formato de logs
LOG_FORMAT = '%(asctime)s | %(name)s | %(levelname)s | %(message)s'
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

# Configuración de rotación
MAX_BYTES = 10 * 1024 * 1024  # 10 MB
BACKUP_COUNT = 5  # Mantener 5 archivos de backup


def get_logger(name, level=logging.INFO):
    """
    Obtiene un logger configurado para un módulo específico.
    
    Args:
        name (str): Nombre del módulo (ej: 'access_manager', 'data_manager')
        level (int): Nivel de logging (default: INFO)
    
    Returns:
        logging.Logger: Logger configurado
    
    Ejemplo:
        logger = get_logger(__name__)
        logger.info("Mensaje informativo")
        logger.error("Error crítico")
    """
    logger = logging.getLogger(name)
    
    # Evitar duplicar handlers si ya está configurado
    if logger.handlers:
        return logger
    
    logger.setLevel(level)
    
    # Formatter común
    formatter = logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT)
    
    # --- Handler 1: Archivo específico del módulo ---
    # Cada módulo tiene su propio archivo de log
    module_log_file = os.path.join(LOGS_DIR, f'{name}.log')
    file_handler = RotatingFileHandler(
        module_log_file,
        maxBytes=MAX_BYTES,
        backupCount=BACKUP_COUNT,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)  # Guardar todo en archivo
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # --- Handler 2: Archivo general de la aplicación ---
    # Todos los logs también van a un archivo general
    app_log_file = os.path.join(LOGS_DIR, 'app.log')
    app_handler = RotatingFileHandler(
        app_log_file,
        maxBytes=MAX_BYTES,
        backupCount=BACKUP_COUNT,
        encoding='utf-8'
    )
    app_handler.setLevel(logging.INFO)  # Solo INFO y superiores en el log general
    app_handler.setFormatter(formatter)
    logger.addHandler(app_handler)
    
    # --- Handler 3: Consola (solo para desarrollo) ---
    # Mostrar logs en consola solo para WARNING y superiores
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # --- Handler 4: Archivo de errores críticos ---
    # Todos los errores van a un archivo especial
    error_log_file = os.path.join(LOGS_DIR, 'errors.log')
    error_handler = RotatingFileHandler(
        error_log_file,
        maxBytes=MAX_BYTES,
        backupCount=BACKUP_COUNT,
        encoding='utf-8'
    )
    error_handler.setLevel(logging.ERROR)
    error_handler.setFormatter(formatter)
    logger.addHandler(error_handler)
    
    return logger


def get_access_logger():
    """
    Logger específico para eventos de control de acceso.
    Registra todos los intentos de acceso (permitidos y denegados).
    """
    logger = logging.getLogger('access_events')
    
    if logger.handlers:
        return logger
    
    logger.setLevel(logging.INFO)
    
    # Formato especial para eventos de acceso
    access_format = '%(asctime)s | %(levelname)s | %(message)s'
    formatter = logging.Formatter(access_format, datefmt=DATE_FORMAT)
    
    # Archivo específico para eventos de acceso
    access_log_file = os.path.join(LOGS_DIR, 'access_events.log')
    access_handler = RotatingFileHandler(
        access_log_file,
        maxBytes=MAX_BYTES,
        backupCount=BACKUP_COUNT,
        encoding='utf-8'
    )
    access_handler.setLevel(logging.INFO)
    access_handler.setFormatter(formatter)
    logger.addHandler(access_handler)
    
    return logger


def log_access_event(dni, nombre, resultado, tipo_permiso='N/A', mensaje=''):
    """
    Registra un evento de acceso de forma estructurada.
    
    Args:
        dni (str): DNI de la persona
        nombre (str): Nombre completo
        resultado (str): 'PERMITIDO' o 'DENEGADO'
        tipo_permiso (str): Tipo de permiso (FAP, FAO, Nómina, etc.)
        mensaje (str): Mensaje adicional
    """
    access_logger = get_access_logger()
    
    log_msg = f"DNI: {dni} | Nombre: {nombre} | Resultado: {resultado} | Permiso: {tipo_permiso}"
    if mensaje:
        log_msg += f" | {mensaje}"
    
    if resultado == 'PERMITIDO':
        access_logger.info(log_msg)
    else:
        access_logger.warning(log_msg)


# Configuración inicial al importar el módulo
if __name__ != '__main__':
    # Crear logger principal de la aplicación
    app_logger = get_logger('app')
    app_logger.info('=' * 80)
    app_logger.info('Sistema de Control de Acceso AOS - Iniciado')
    app_logger.info(f'Fecha: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    app_logger.info('=' * 80)
