import os
import sys

# --- CONFIGURACIÓN GENERAL ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- CONFIGURACIÓN DE LA APLICACIÓN FLASK ---
SECRET_KEY = 'tu_clave_secreta_aqui_CAMBIALA'
ADMIN_USERNAME = 'Seguridad'
ADMIN_PASSWORD = 'ControldeAcceso1.0'

# --- CONFIGURACIÓN DE EMAIL ---
EMAIL_SENDER = 'acceso.alcorta@gmail.com'
EMAIL_PASSWORD = 'vmcg dqel hhdx zzqj'
EMAIL_RECEIVER = 'rlaforcada@irsa.com.ar'
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587

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
