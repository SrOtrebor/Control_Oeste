import os
import sys
import secrets
from dotenv import load_dotenv

# --- CARGAR VARIABLES DE ENTORNO ---
# Buscar archivo .env en el directorio del proyecto
load_dotenv()

# --- FUNCIÓN PARA OBTENER LA RUTA CORRECTA ---
def get_base_path():
    """ Obtiene la ruta base correcta, funciona para desarrollo y para PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        # Si está compilado, la base es la carpeta temporal de PyInstaller
        return sys._MEIPASS
    # Si está en desarrollo, la base es la carpeta donde está este archivo (config.py)
    return os.path.abspath(os.path.dirname(__file__))

# --- RUTAS DE ARCHIVOS Y CARPETAS ---
BASE_DIR = get_base_path()

# --- CONFIGURACIÓN DE LA APLICACIÓN FLASK ---
# Generar SECRET_KEY segura si no existe en variables de entorno
SECRET_KEY = os.getenv('SECRET_KEY')
if not SECRET_KEY:
    # Generar una clave aleatoria segura de 32 bytes (64 caracteres hex)
    SECRET_KEY = secrets.token_hex(32)
    print("⚠️  ADVERTENCIA: Se generó una SECRET_KEY temporal.")
    print("   Para producción, agrega SECRET_KEY a tu archivo .env")

# Credenciales de administrador (desde variables de entorno)
ADMIN_USERNAME = os.getenv('ADMIN_USERNAME', 'Seguridad')
ADMIN_PASSWORD = os.getenv('ADMIN_PASSWORD', 'ControldeAcceso1.0')

# Advertencia si se usan credenciales por defecto
if ADMIN_PASSWORD == 'ControldeAcceso1.0':
    print("⚠️  ADVERTENCIA: Usando contraseña de admin por defecto.")
    print("   Cambia ADMIN_PASSWORD en tu archivo .env para mayor seguridad.")

# --- CONFIGURACIÓN DE EMAIL ---
EMAIL_SENDER = os.getenv('EMAIL_SENDER', 'acceso.alcorta@gmail.com')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD', '')
EMAIL_RECEIVER = os.getenv('EMAIL_RECEIVER', 'rlaforcada@irsa.com.ar')
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))

# Advertencia si no hay contraseña de email configurada
if not EMAIL_PASSWORD:
    print("⚠️  ADVERTENCIA: EMAIL_PASSWORD no configurado.")
    print("   El envío de reportes por email no funcionará.")
    print("   Agrega EMAIL_PASSWORD a tu archivo .env")

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