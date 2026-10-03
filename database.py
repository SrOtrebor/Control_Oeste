import sqlite3
import os
import logging
from contextlib import contextmanager

logger = logging.getLogger(__name__)
from config import BASE_DIR

DB_PATH = os.path.join(BASE_DIR, 'control_acceso.db')

@contextmanager
def get_db_connection():
    """Context manager para manejar conexiones seguras a SQLite"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row  # Para poder acceder a las columnas por nombre
    try:
        yield conn
        conn.commit()
    except Exception as e:
        conn.rollback()
        logger.error(f"Error en base de datos: {e}")
        raise
    finally:
        conn.close()

def init_db():
    """Crea las tablas iniciales si no existen"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # 1. Tabla de Autorizaciones (FAPs, FAOs, Nóminas)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS autorizaciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                dni TEXT NOT NULL,
                nombre TEXT,
                apellido TEXT,
                tipo_permiso TEXT,      -- 'FAP', 'FAO', 'NOMINA'
                num_permiso TEXT,       -- ID del FAO/FAP o identificador de nómina
                local TEXT,             -- Marca/Empresa
                tarea TEXT,
                fecha_inicio DATE,
                fecha_fin DATE,
                huella_template BLOB,   -- Para la futura integración biométrica
                foto_path TEXT,         -- Ruta o b64 de la foto
                id_empresa INTEGER,
                tipo_empleado TEXT,
                activo BOOLEAN DEFAULT 1
            )
        ''')
        
        # Índice para búsqueda rápida por DNI
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_autorizaciones_dni ON autorizaciones(dni)')
        
        # 2. Tabla de Registros Diarios (Accesos)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS registros_accesos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha DATE NOT NULL,
                hora TIME NOT NULL,
                dni TEXT NOT NULL,
                nombre TEXT,
                evento TEXT,            -- 'ENTRADA', 'SALIDA', 'RECHAZO'
                tipo_permiso TEXT,
                local TEXT,
                resultado TEXT,         -- 'AUTORIZADO', 'DENEGADO'
                motivo_rechazo TEXT,
                registrado_por TEXT,    -- Usuario/Garita
                puerta TEXT DEFAULT 'Master',
                tipo_visita TEXT
            )
        ''')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_accesos_fecha ON registros_accesos(fecha)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_accesos_dni ON registros_accesos(dni)')

        # 3. Tabla de Fichajes Diarios (Control de Asistencia)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS registros_fichajes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha DATE NOT NULL,
                dni TEXT NOT NULL,
                nombre TEXT,
                hora_entrada TIME,
                hora_salida TIME,
                estado TEXT,             -- 'EN CURSO', 'COMPLETADO'
                categoria TEXT,
                puesto_historico TEXT,
                horas_diurnas REAL,
                horas_nocturnas REAL,
                puerta TEXT DEFAULT 'Master',
                id_puesto_asignado INTEGER
            )
        ''')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_fichajes_fecha ON registros_fichajes(fecha)')
        
        # 4. Tabla de Estado Adentro (Gente que está actualmente en el predio)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS estado_adentro (
                dni TEXT PRIMARY KEY,
                nombre TEXT,
                hora_ingreso TIME,
                tipo TEXT,
                local TEXT,
                puerta TEXT DEFAULT 'Master'
            )
        ''')

        # 5. Tabla de Puestos Operativos
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS puestos_operativos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL UNIQUE,
                activo BOOLEAN DEFAULT 1
            )
        ''')

        # 6. Tabla de Puestos Fisicos
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS puestos_fisicos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL UNIQUE,
                sector TEXT,
                hora_inicio TIME,
                hora_fin TIME,
                activo BOOLEAN DEFAULT 1
            )
        ''')

        # 7. Tabla de Empresas Contratistas
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS empresas_contratistas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL UNIQUE,
                activo BOOLEAN DEFAULT 1
            )
        ''')

    logger.info("Base de datos SQLite inicializada correctamente.")

if __name__ == '__main__':
    init_db()
