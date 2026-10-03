import sqlite3
import json
import os

DB_PATH = 'control_acceso.db'
OUTPUT_PATH = 'public/migracion.json'

def export_data():
    if not os.path.exists(DB_PATH):
        print(f"Error: No se encontró la base de datos en {DB_PATH}")
        return

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    data = {
        'lista_negra': [],
        'nominas': []
    }

    # Exportar Lista Negra
    try:
        cursor.execute("SELECT * FROM lista_negra")
        for row in cursor.fetchall():
            data['lista_negra'].append(dict(row))
        print(f"Exportados {len(data['lista_negra'])} registros de Lista Negra.")
    except Exception as e:
        print(f"Error leyendo lista_negra: {e}")

    # Exportar Nóminas (FAO/FAP) desde BD
    try:
        cursor.execute("SELECT * FROM autorizaciones WHERE tipo_permiso='NOMINA'")
        for row in cursor.fetchall():
            data['nominas'].append({
                'dni': row['dni'],
                'nombre': row['nombre'],
                'categoria': row['categoria'] if 'categoria' in row.keys() else 'N/A',
                'puesto_especifico': row['puesto_especifico'] if 'puesto_especifico' in row.keys() else '',
                'activo': True
            })
    except Exception as e:
        print(f"Error leyendo autorizaciones: {e}")

    # Exportar desde JSONs (FAOs y FAPs)
    try:
        if os.path.exists('MetadatosFAOs.json'):
            with open('MetadatosFAOs.json', 'r', encoding='utf-8') as f:
                faos = json.load(f)
                for f_data in faos:
                    for p in f_data.get('personal', []):
                        if p.get('activo', True):
                            data['nominas'].append({
                                'dni': str(p.get('numeroDocumento', '')),
                                'nombre': f"{p.get('apellido', '')} {p.get('nombre', '')}".strip(),
                                'categoria': 'FAO',
                                'puesto_especifico': f_data.get('marca', ''),
                                'activo': True
                            })
                            
        if os.path.exists('MetadatosFAPs.json'):
            with open('MetadatosFAPs.json', 'r', encoding='utf-8') as f:
                faps = json.load(f)
                for f_data in faps:
                    for p in f_data.get('personal', []):
                        if p.get('activo', True):
                            data['nominas'].append({
                                'dni': str(p.get('numeroDocumento', '')),
                                'nombre': f"{p.get('apellido', '')} {p.get('nombre', '')}".strip(),
                                'categoria': 'FAP',
                                'puesto_especifico': f_data.get('marca', ''),
                                'activo': True
                            })
                            
        print(f"Exportadas {len(data['nominas'])} personas en total (FAOs, FAPs, Nóminas).")
    except Exception as e:
        print(f"Error leyendo JSONs: {e}")

    conn.close()

    # Guardar en public/
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nDatos exportados exitosamente a {OUTPUT_PATH}")

if __name__ == '__main__':
    export_data()
