import sqlite3
from database import get_db_connection

def apply_migrations():
    """Crea la tabla lista_negra para Personas de Interés."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS lista_negra (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                dni TEXT UNIQUE NOT NULL,
                nombre TEXT,
                motivo TEXT,
                fecha_agregado TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        print("Tabla 'lista_negra' asegurada.")

if __name__ == "__main__":
    apply_migrations()
