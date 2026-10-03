import os
import shutil
import zipfile
from datetime import datetime
import time
import glob
from logger_config import get_logger

logger = get_logger(__name__)

BACKUP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'backups')
DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'control_acceso.db')
MAX_BACKUP_DAYS = 15

def ensure_backup_dir():
    os.makedirs(BACKUP_DIR, exist_ok=True)

def cleanup_old_backups():
    """Mantiene solo los backups de los últimos MAX_BACKUP_DAYS días."""
    try:
        now = time.time()
        for f in glob.glob(os.path.join(BACKUP_DIR, '*.zip')):
            if os.stat(f).st_mtime < now - (MAX_BACKUP_DAYS * 86400):
                os.remove(f)
                logger.info(f"Backup antiguo eliminado: {f}")
    except Exception as e:
        logger.error(f"Error limpiando backups antiguos: {e}")

def create_backup(manual=False):
    """
    Comprime la base de datos SQLite en un archivo ZIP.
    Retorna la ruta absoluta del ZIP generado.
    """
    ensure_backup_dir()
    cleanup_old_backups()
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    prefix = "manual_" if manual else "auto_"
    zip_filename = f"backup_{prefix}{timestamp}.zip"
    zip_path = os.path.join(BACKUP_DIR, zip_filename)
    
    try:
        if not os.path.exists(DB_FILE):
            logger.warning(f"No se encontró {DB_FILE} para hacer backup.")
            return None
            
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(DB_FILE, os.path.basename(DB_FILE))
            
        logger.info(f"Backup creado exitosamente: {zip_path}")
        return zip_path
    except Exception as e:
        logger.error(f"Error creando backup: {e}")
        return None
