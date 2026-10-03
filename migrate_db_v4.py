import sqlite3
from database import DB_PATH

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

def try_add_col(table, col, dtype):
    try:
        cursor.execute(f'ALTER TABLE {table} ADD COLUMN {col} {dtype}')
        print(f'Added {col} to {table}')
    except sqlite3.OperationalError as e:
        print(f'Skipped {col} in {table}: {e}')

try_add_col('registros_accesos', 'puerta', 'TEXT DEFAULT "Master"')
try_add_col('registros_accesos', 'tipo_visita', 'TEXT')
try_add_col('estado_adentro', 'puerta', 'TEXT DEFAULT "Master"')
try_add_col('registros_fichajes', 'puerta', 'TEXT DEFAULT "Master"')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS puestos_operativos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL UNIQUE,
        activo BOOLEAN DEFAULT 1
    )
''')

try:
    cursor.execute("INSERT OR IGNORE INTO puestos_operativos (nombre) VALUES ('Portón 1'), ('Portón 2'), ('Portón 3'), ('Rondín')")
except Exception as e:
    pass

conn.commit()
conn.close()
print('Migration complete')
