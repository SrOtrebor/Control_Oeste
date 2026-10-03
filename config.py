import os
import sys
import shutil
import secrets
from dotenv import load_dotenv

# Prevenir que fallos de encoding en Windows (cp1252) detengan la ejecución
if hasattr(sys, 'stdout') and sys.stdout is not None:
    try:
        sys.stdout.reconfigure(errors='replace')
    except Exception:
        pass
if hasattr(sys, 'stderr') and sys.stderr is not None:
    try:
        sys.stderr.reconfigure(errors='replace')
    except Exception:
        pass

# --- RUTAS DE ACCESO ---
def get_bundle_path():
    """Obtiene la ruta a los recursos empaquetados (templates, static)."""
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.abspath(os.path.dirname(__file__))

def get_base_path():
    """Obtiene la ruta base donde se almacenan los archivos de datos permanentes (Excels, registros, backups).
    En modo compilado (.exe) es el directorio donde está el ejecutable.
    En modo desarrollo es el directorio del código fuente."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.dirname(__file__))

BUNDLE_DIR = get_bundle_path()
BASE_DIR = get_base_path()

# Copiar archivos Excel por defecto si no existen en BASE_DIR (para ejecutable compilado)
if getattr(sys, 'frozen', False):
    for f in ['ListadoFAPs.xlsx', 'ListadoFAOs.xlsx', 'excepciones.xlsx', 'nominas_persistentes.xlsx']:
        dst_path = os.path.join(BASE_DIR, f)
        src_path = os.path.join(BUNDLE_DIR, f)
        if not os.path.exists(dst_path) and os.path.exists(src_path):
            try:
                shutil.copy2(src_path, dst_path)
            except Exception:
                pass

# --- CARGAR VARIABLES DE ENTORNO ---
env_path = os.path.join(BASE_DIR, '.env')
if os.path.exists(env_path):
    load_dotenv(env_path)
else:
    load_dotenv()

# --- CONFIGURACIÓN DE LA APLICACIÓN FLASK ---
# Generar SECRET_KEY segura si no existe en variables de entorno
SECRET_KEY = os.getenv('SECRET_KEY')
if not SECRET_KEY:
    # Generar una clave aleatoria segura de 32 bytes (64 caracteres hex)
    SECRET_KEY = secrets.token_hex(32)
    print("[ADVERTENCIA] Se genero una SECRET_KEY temporal.")
    print("   Para produccion, agregue SECRET_KEY a su archivo .env")

# Credenciales de administrador (desde variables de entorno)
ADMIN_USERNAME = os.getenv('ADMIN_USERNAME', 'Seguridad')
ADMIN_PASSWORD = os.getenv('ADMIN_PASSWORD', 'ControldeAcceso1.0')

# Advertencia si se usan credenciales por defecto
if ADMIN_PASSWORD == 'ControldeAcceso1.0':
    print("[ADVERTENCIA] Usando contrasena de admin por defecto.")
    print("   Cambie ADMIN_PASSWORD en su archivo .env para mayor seguridad.")

# --- CONFIGURACIÓN DE EMAIL ---
EMAIL_SENDER = os.getenv('EMAIL_SENDER', 'acceso.alcorta@gmail.com')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD', '')
EMAIL_RECEIVER = os.getenv('EMAIL_RECEIVER', 'rlaforcada@irsa.com.ar')
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))

# Advertencia si no hay contraseña de email configurada
if not EMAIL_PASSWORD:
    print("[ADVERTENCIA] EMAIL_PASSWORD no configurado.")
    print("   El envio de reportes por email no funcionara.")
    print("   Agregue EMAIL_PASSWORD a su archivo .env")

# --- RUTAS DE ARCHIVOS ---
REGISTROS_DIARIOS_DIR = os.path.join(BASE_DIR, 'registros_diarios')
REGISTROS_FICHAJES_DIR = os.path.join(BASE_DIR, 'registros_fichajes')
REGISTROS_VISITAS_DIR = os.path.join(BASE_DIR, 'registros_visitas')
EXCEL_FAP = os.path.join(BASE_DIR, 'ListadoFAPs.xlsx')
EXCEL_FAO = os.path.join(BASE_DIR, 'ListadoFAOs.xlsx')
EXCEL_EXCEPCIONES = os.path.join(BASE_DIR, 'excepciones.xlsx')
EXCEL_NOMINAS = os.path.join(BASE_DIR, 'nominas_persistentes.xlsx')

# --- NOMBRES DE COLUMNAS ESTANDARIZADOS ---
COL_DNI = 'DNI'
COL_NOMBRE_APELLIDO = 'Nombre y Apellido'
COL_NUM_PERMISO = 'Num_Permiso'
COL_VENCE = 'Vence'
COL_LOCAL = 'Local'
COL_TAREA = 'Tarea'
COL_TIPO_PERMISO = 'Tipo de Permiso'

# --- NOMBRES DE COLUMNAS ORIGINALES (PARA MAPEO) ---
COL_DNI_FAP_ORIGINAL = 'Numero'
COL_NOMBRE_FAP_ORIGINAL = 'Nombre'
COL_APELLIDO_FAP_ORIGINAL = 'Apellido'
COL_NUM_PERMISO_FAP_ORIGINAL = 'FAP'
COL_VENCE_FAP_ORIGINAL = 'Fecha Fin'
COL_LOCAL_FAP_ORIGINAL = 'Marca'

COL_DNI_FAO_ORIGINAL = 'Numero'
COL_NOMBRE_FAO_ORIGINAL = 'Nombre'
COL_APELLIDO_FAO_ORIGINAL = 'Apellido'
COL_NUM_PERMISO_FAO_ORIGINAL = 'FAO'
COL_VENCE_FAO_ORIGINAL = 'Fecha Fin'
COL_LOCAL_FAO_ORIGINAL = 'Marca'
COL_TAREA_FAO_ORIGINAL = 'Tarea'