import sys
import os
import json
sys.path.append('c:/PC laburo muleto/Control de Acceso Oeste')
from irsa_client import IRSAClient
from irsa_config import get_irsa_credentials

def main():
    c = get_irsa_credentials()
    client = IRSAClient()
    client.login(c['username'], c['password'])
    
    # 1. Fetch
    detalle = client._obtener_detalle('/api/faos/fao', 1266421)
    if not detalle:
        print("No se pudo obtener el FAO 1266421")
        return
        
    print(f"Obtenido FAO: {detalle.get('id')}")
    
    # 2. Append to JSON
    p = 'c:/PC laburo muleto/Control de Acceso Oeste/MetadatosFAOs.json'
    faos_data = []
    if os.path.exists(p):
        with open(p, 'r', encoding='utf-8') as f:
            faos_data = json.load(f)
            
    # Remove if exists to avoid duplicates
    faos_data = [f for f in faos_data if str(f.get('id')) != '1266421']
    faos_data.append(detalle)
    
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(faos_data, f, indent=2, ensure_ascii=False)
        
    # 3. Regenerate Excel
    validos = [f for f in faos_data if str(f.get('estado')) in ['4', '6']]
    client._guardar_excel_faos(validos)
    
    print("Inyectado exitosamente. Sincronizando con Firebase...")
    
    # 4. Sincronizar a firebase
    import satellite_db
    satellite_db.init_firebase()
    satellite_db.sincronizar_bd_completa(forzar=True)
    
if __name__ == '__main__':
    main()
