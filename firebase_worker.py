import firebase_admin
from firebase_admin import credentials, firestore
import time
import subprocess
import os
import json
import logging
from google.cloud.firestore_v1.base_query import FieldFilter

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuración de Firebase Admin usando credenciales locales
try:
    cred = credentials.Certificate('serviceAccountKey.json')
    firebase_admin.initialize_app(cred)
except Exception as e:
    logger.error(f"Error inicializando Firebase: {e}")
    os._exit(1)

db = firestore.client()

def process_sync_irsa(doc_id):
    logger.info(f"[{doc_id}] Iniciando sincronización IRSA...")
    try:
        # 1. Obtener credenciales desde Firestore
        config_doc = db.collection('config').document('irsa').get()
        if config_doc.exists:
            data = config_doc.to_dict()
            user = data.get('username')
            pwd = data.get('password')
            if user and pwd:
                # Escribir usando la encriptación oficial del sistema
                from irsa_config import save_irsa_credentials
                save_irsa_credentials(user, pwd)
                    
        from irsa_client import IRSAClient
        client = IRSAClient()
        res = client.sincronizar_todo()
        
        batch = db.batch()
        nominas_ref = db.collection('nominas')
        
        count = 0
        
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        faos_path = os.path.join(BASE_DIR, 'MetadatosFAOs.json')
        faps_path = os.path.join(BASE_DIR, 'MetadatosFAPs.json')
        
        logger.info(f"[{doc_id}] Buscando {faos_path} ... existe: {os.path.exists(faos_path)}")
        
        tramites_ref = db.collection('tramites_irsa')
        
        if os.path.exists(faos_path):
            with open(faos_path, 'r', encoding='utf-8') as f:
                faos = json.load(f)
                logger.info(f"[{doc_id}] Encontrados {len(faos)} nodos de FAOs")
                for f_data in faos:
                    # Guardar el tramite en Firestore para el dashboard
                    tramites_ref.document(str(f_data.get('id'))).set({
                        'id': str(f_data.get('id')),
                        'tipo': 'FAO',
                        'empresa': f_data.get('marca', ''),
                        'fecha_inicio': str(f_data.get('fechaInicio', '')).split('T')[0],
                        'fecha_fin': str(f_data.get('fechaFin', '')).split('T')[0],
                        'estado_codigo': str(f_data.get('estado', '')),
                        'personal': [{'nombre': p.get('nombre', ''), 'apellido': p.get('apellido', ''), 'dni': p.get('numeroDocumento', '')} for p in f_data.get('personal', []) if p.get('activo', True)]
                    })
                    # Solo autorizar en puerta si está aprobado
                    if str(f_data.get('estado', '')) not in ['4', '6']:
                        continue
                        
                    for p in f_data.get('personal', []):
                        if p.get('activo', True):
                            dni = str(p.get('numeroDocumento', ''))
                            if dni:
                                doc_ref = nominas_ref.document(dni)
                                batch.set(doc_ref, {
                                    'dni': dni,
                                    'nombre': f"{p.get('apellido', '')} {p.get('nombre', '')}".strip(),
                                    'categoria': 'FAO',
                                    'puesto_especifico': f_data.get('marca', ''),
                                    'nro_tramite': str(f_data.get('id', '')),
                                    'fecha_vencimiento': str(f_data.get('fechaFin', '')).split('T')[0],
                                    'activo': True
                                }, merge=True)
                                count += 1
                                
        logger.info(f"[{doc_id}] Conteo despues de FAOs: {count}")

        if os.path.exists(faps_path):
            with open(faps_path, 'r', encoding='utf-8') as f:
                faps = json.load(f)
                for f_data in faps:
                    # Guardar el tramite en Firestore para el dashboard
                    tramites_ref.document(str(f_data.get('id'))).set({
                        'id': str(f_data.get('id')),
                        'tipo': 'FAP',
                        'empresa': f_data.get('local', ''),
                        'fecha_inicio': str(f_data.get('fechaInicio', '')).split('T')[0],
                        'fecha_fin': str(f_data.get('fechaFin', '')).split('T')[0],
                        'estado_codigo': str(f_data.get('estado', '')),
                        'personal': [{'nombre': p.get('nombre', ''), 'apellido': p.get('apellido', ''), 'dni': p.get('numeroDocumento', '')} for p in f_data.get('personal', []) if p.get('activo', True)]
                    })
                    
                    if str(f_data.get('estado', '')) not in ['4', '6']:
                        continue
                        
                    for p in f_data.get('personal', []):
                        if p.get('activo', True):
                            dni = str(p.get('numeroDocumento', ''))
                            if dni:
                                doc_ref = nominas_ref.document(dni)
                                batch.set(doc_ref, {
                                    'dni': dni,
                                    'nombre': f"{p.get('apellido', '')} {p.get('nombre', '')}".strip(),
                                    'categoria': 'FAP',
                                    'puesto_especifico': f_data.get('marca', ''),
                                    'nro_tramite': str(f_data.get('id', '')),
                                    'fecha_vencimiento': str(f_data.get('fechaFin', '')).split('T')[0],
                                    'activo': True
                                }, merge=True)
                                count += 1

        # Enviar batch solo si hay datos (límite 500 ops por batch, pero como son pocos...)
        if count > 0:
            batch.commit()
            
        db.collection('system_commands').document(doc_id).update({
            'status': 'DONE',
            'result': f'Sincronizados {count} registros desde IRSA.'
        })
        logger.info(f"[{doc_id}] IRSA Sync completado exitosamente.")
    except Exception as e:
        logger.error(f"[{doc_id}] Error en Sincronización IRSA: {e}")
        db.collection('system_commands').document(doc_id).update({
            'status': 'ERROR',
            'result': str(e)
        })

def process_force_sync_fao(doc_id, fao_id):
    logger.info(f"[{doc_id}] Iniciando inyección manual del FAO {fao_id}...")
    try:
        from irsa_client import IRSAClient
        from irsa_config import get_irsa_credentials
        
        c = get_irsa_credentials()
        if not c:
            raise Exception("No hay credenciales configuradas")
            
        client = IRSAClient()
        client.login(c['username'], c['password'])
        
        detalle = client._obtener_detalle('/api/faos/fao', fao_id)
        if not detalle:
            raise Exception(f"FAO {fao_id} no encontrado en IRSA")
            
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        p = os.path.join(BASE_DIR, 'MetadatosFAOs.json')
        faos_data = []
        if os.path.exists(p):
            with open(p, 'r', encoding='utf-8') as f:
                faos_data = json.load(f)
                
        faos_data = [f for f in faos_data if str(f.get('id')) != str(fao_id)]
        faos_data.append(detalle)
        
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(faos_data, f, indent=2, ensure_ascii=False)
            
        validos = [f for f in faos_data if str(f.get('estado')) in ['4', '6']]
        client._guardar_excel_faos(validos)
        
        # Subir a firebase
        db.collection('tramites_irsa').document(str(detalle.get('id'))).set({
            'id': str(detalle.get('id')),
            'marca': detalle.get('marca', ''),
            'estado': detalle.get('estado', ''),
            'fechaFin': detalle.get('fechaFin', ''),
            'personal': [p.get('numeroDocumento', '') for p in detalle.get('personal', []) if p.get('activo', True)]
        })
        
        db.collection('system_commands').document(doc_id).update({
            'status': 'DONE',
            'result': f'FAO {fao_id} inyectado exitosamente.'
        })
    except Exception as e:
        logger.error(f"[{doc_id}] Error inyectando FAO: {e}")
        db.collection('system_commands').document(doc_id).update({
            'status': 'ERROR',
            'result': str(e)
        })

def process_save_excepciones(doc_id, nomina_nueva, empresa, vigencia_desde, vigencia_hasta):
    logger.info(f"[{doc_id}] Guardando nómina de excepciones: {empresa} con {len(nomina_nueva)} personas")
    try:
        from database import get_db_connection
        import data_manager
        
        if nomina_nueva:
            with get_db_connection() as conn:
                cursor = conn.cursor()
                for persona in nomina_nueva:
                    dni = persona.get('dni')
                    apellido = persona.get('apellido', '')
                    nombre = persona.get('nombre', '')
                    
                    # Eliminar registro viejo de este DNI en esta misma empresa si existe
                    cursor.execute("DELETE FROM autorizaciones WHERE dni=? AND tipo_permiso='NOMINA' AND local=?", (dni, empresa))
                    
                    # Insertar
                    cursor.execute('''
                        INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, local, fecha_inicio, fecha_fin, activo)
                        VALUES (?, ?, ?, 'NOMINA', ?, ?, ?, 1)
                    ''', (dni, nombre, apellido, empresa, vigencia_desde, vigencia_hasta))
            
            # Recargar cache
            try:
                data_manager.cargar_autorizaciones()
            except Exception as cache_e:
                logger.warning(f"No se pudo recargar caché al instante: {cache_e}")

        db.collection('system_commands').document(doc_id).update({
            'status': 'DONE',
            'result': f'Excepción guardada en SQLite. {len(nomina_nueva)} personas autorizadas.'
        })
    except Exception as e:
        logger.error(f"[{doc_id}] Error guardando excepciones: {e}")
        db.collection('system_commands').document(doc_id).update({
            'status': 'ERROR',
            'result': str(e)
        })

def process_ota_update(doc_id):
    logger.info(f"[{doc_id}] Iniciando actualización remota (OTA)...")
    try:
        subprocess.run(["git", "pull"], check=True)
        
        db.collection('system_commands').document(doc_id).update({
            'status': 'DONE',
            'result': 'Git Pull exitoso. Reiniciando obrero local...'
        })
        logger.info(f"[{doc_id}] OTA completado. Reiniciando servicio...")
        time.sleep(2)
        os._exit(0)  # Cierra para que el .bat o servicio lo vuelva a levantar
    except Exception as e:
        logger.error(f"[{doc_id}] Error en actualización OTA: {e}")
        db.collection('system_commands').document(doc_id).update({
            'status': 'ERROR',
            'result': str(e)
        })

def on_command_snapshot(doc_snapshot, changes, read_time):
    for change in changes:
        if change.type.name == 'ADDED' or change.type.name == 'MODIFIED':
            doc = change.document
            data = doc.to_dict()
            if data.get('status') == 'PENDING':
                action = data.get('action')
                doc_id = doc.id
                
                logger.info(f"Comando detectado: {action} (ID: {doc_id})")
                
                # Marcar en progreso para que nadie más lo toque
                db.collection('system_commands').document(doc_id).update({'status': 'IN_PROGRESS'})
                
                if action == 'SYNC_IRSA':
                    process_sync_irsa(doc_id)
                elif action == 'UPDATE_OTA':
                    process_ota_update(doc_id)
                elif action == 'FORCE_SYNC_FAO':
                    fao_id = data.get('fao_id')
                    if fao_id:
                        process_force_sync_fao(doc_id, fao_id)
                    else:
                        db.collection('system_commands').document(doc_id).update({'status': 'ERROR', 'result': 'fao_id no especificado'})
                elif action == 'SAVE_EXCEPCIONES':
                    process_save_excepciones(doc_id, data.get('nomina', []), data.get('empresa'), data.get('vigencia_desde'), data.get('vigencia_hasta'))
                else:
                    db.collection('system_commands').document(doc_id).update({
                        'status': 'ERROR',
                        'result': f'Comando desconocido: {action}'
                    })

if __name__ == '__main__':
    commands_ref = db.collection('system_commands')
    # Filtramos comandos pendientes
    query = commands_ref.where(filter=FieldFilter('status', '==', 'PENDING'))
    commands_watch = query.on_snapshot(on_command_snapshot)
    
    logger.info("👷‍♂️ Obrero Python iniciado y conectado a Firebase.")
    logger.info("Escuchando órdenes de la nube (SYNC_IRSA, UPDATE_OTA)...")
    
    try:
        last_sync = 0
        while True:
            time.sleep(1)
            current_time = time.time()
            if current_time - last_sync > 10:
                try:
                    from database import get_db_connection
                    from datetime import datetime
                    import pandas as pd
                    
                    hoy = datetime.now().strftime('%Y-%m-%d')
                    total_adentro = 0
                    
                    # Calcular total adentro
                    with get_db_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("SELECT COUNT(DISTINCT dni) as count FROM estado_adentro")
                        row = cursor.fetchone()
                        if row:
                            total_adentro = row['count']
                    
                    db.collection('config').document('live_stats').set({
                        'total_adentro': total_adentro,
                        'last_updated': current_time
                    }, merge=True)
                    last_sync = current_time
                except Exception as e:
                    logger.error(f"Error syncing live stats: {e}")
                    
    except KeyboardInterrupt:
        logger.info("Obrero detenido.")
