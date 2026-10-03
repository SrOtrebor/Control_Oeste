import sqlite3
import os
import logging
from datetime import datetime

# Configuracion del logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger('SatelliteDB')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'satellite.db')

def get_db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=20)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Inicializa la base de datos local del satélite (100% offline-first)."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # Registros locales de la garita
        # sync: 0=Pendiente de subir al Master, 1=Sincronizado
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS registros_locales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                dni TEXT NOT NULL,
                puerta TEXT NOT NULL,
                estado TEXT NOT NULL,
                tipo_visita TEXT DEFAULT '',
                sync INTEGER DEFAULT 0,
                sync_timestamp TEXT
            )
        ''')
        
        # Caché de Nóminas para consultar offline
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cache_nominas (
                id INTEGER PRIMARY KEY,
                dni TEXT UNIQUE,
                nombre TEXT,
                categoria TEXT,
                puesto TEXT,
                activo INTEGER
            )
        ''')
        
        # Caché de Lista Negra para bloquear ingresos offline
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cache_lista_negra (
                id INTEGER PRIMARY KEY,
                dni TEXT UNIQUE,
                motivo TEXT
            )
        ''')
        
        # Caché de Puestos (por si la garita necesita seleccionar un destino para la visita)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cache_puestos (
                id INTEGER PRIMARY KEY,
                nombre TEXT UNIQUE
            )
        ''')
        
        # Configuración local (ej: última fecha de sincronización, puerta asignada a este satélite)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS config_local (
                clave TEXT PRIMARY KEY,
                valor TEXT
            )
        ''')
        
        # Insertar valores por defecto si no existen
        cursor.execute("INSERT OR IGNORE INTO config_local (clave, valor) VALUES ('puerta_actual', 'Satelite 1')")
        cursor.execute("INSERT OR IGNORE INTO config_local (clave, valor) VALUES ('master_ip', 'http://127.0.0.1:5000')")
        cursor.execute("INSERT OR IGNORE INTO config_local (clave, valor) VALUES ('last_sync', 'Nunca')")
        
        conn.commit()
        logger.info("Base de datos Satélite inicializada correctamente.")

def registrar_acceso(dni, estado, tipo_visita=''):
    """Registra una entrada/salida/rechazo/visita en la BD local marcándola como pendiente de sync"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT valor FROM config_local WHERE clave='puerta_actual'")
        puerta = cursor.fetchone()['valor']
        
        timestamp = datetime.now().isoformat()
        
        cursor.execute("""
            INSERT INTO registros_locales (timestamp, dni, puerta, estado, tipo_visita, sync)
            VALUES (?, ?, ?, ?, ?, 0)
        """, (timestamp, dni, puerta, estado, tipo_visita))
        return True

def obtener_registros_pendientes():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM registros_locales WHERE sync = 0")
        return [dict(row) for row in cursor.fetchall()]

def marcar_registros_como_sincronizados(ids):
    if not ids: return
    with get_db_connection() as conn:
        placeholders = ','.join('?' for _ in ids)
        conn.execute(f"UPDATE registros_locales SET sync = 1, sync_timestamp = ? WHERE id IN ({placeholders})",
                     [datetime.now().isoformat()] + ids)

def actualizar_cache(tabla, datos, pk='id'):
    """Reemplaza la caché local con los datos frescos del master"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(f"DELETE FROM {tabla}")
        if not datos: return
        
        columnas = list(datos[0].keys())
        placeholders = ','.join(['?'] * len(columnas))
        query = f"INSERT INTO {tabla} ({','.join(columnas)}) VALUES ({placeholders})"
        
        for fila in datos:
            valores = [fila[col] for col in columnas]
            cursor.execute(query, valores)

def actualizar_ultima_sincronizacion():
    with get_db_connection() as conn:
        ahora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn.execute("UPDATE config_local SET valor = ? WHERE clave = 'last_sync'", (ahora,))
        return ahora

def obtener_config():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT clave, valor FROM config_local")
        return {row['clave']: row['valor'] for row in cursor.fetchall()}

# Inicializar al importar
if not os.path.exists(DB_PATH):
    init_db()
else:
    # Ensure tables exist just in case
    init_db()
