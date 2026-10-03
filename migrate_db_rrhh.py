import sqlite3
import os
from database import get_db_connection, init_db

def apply_migrations():
    """Agrega las nuevas columnas de RRHH si no existen."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # 1. Agregar columnas a autorizaciones
        try:
            cursor.execute("ALTER TABLE autorizaciones ADD COLUMN categoria TEXT DEFAULT 'N/A'")
            print("Columna 'categoria' agregada a autorizaciones.")
        except sqlite3.OperationalError:
            pass # Ya existe
            
        try:
            cursor.execute("ALTER TABLE autorizaciones ADD COLUMN puesto_especifico TEXT DEFAULT ''")
            print("Columna 'puesto_especifico' agregada a autorizaciones.")
        except sqlite3.OperationalError:
            pass

        # 2. Agregar columnas a registros_fichajes
        try:
            cursor.execute("ALTER TABLE registros_fichajes ADD COLUMN categoria TEXT DEFAULT 'N/A'")
            print("Columna 'categoria' agregada a registros_fichajes.")
        except sqlite3.OperationalError:
            pass
            
        try:
            cursor.execute("ALTER TABLE registros_fichajes ADD COLUMN puesto_historico TEXT DEFAULT ''")
            print("Columna 'puesto_historico' agregada a registros_fichajes.")
        except sqlite3.OperationalError:
            pass

        # 3. Agregar columnas de cálculo de horas a registros_fichajes
        try:
            cursor.execute("ALTER TABLE registros_fichajes ADD COLUMN horas_diurnas REAL DEFAULT 0.0")
            print("Columna 'horas_diurnas' agregada a registros_fichajes.")
        except sqlite3.OperationalError:
            pass

        try:
            cursor.execute("ALTER TABLE registros_fichajes ADD COLUMN horas_nocturnas REAL DEFAULT 0.0")
            print("Columna 'horas_nocturnas' agregada a registros_fichajes.")
        except sqlite3.OperationalError:
            pass
            
    print("Migraciones aplicadas con éxito.")

if __name__ == "__main__":
    apply_migrations()
