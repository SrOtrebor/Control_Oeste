import time
import subprocess
import os
import sys
import json
import socket
import hashlib
import logging
import queue
import threading
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)  # Asegura rutas relativas correctas (serviceAccountKey, git, etc.)

# ============================================================
# INSTANCIA ÚNICA: evita dos obreros procesando lo mismo
# (app.py lo relanza periódicamente; si ya hay uno vivo, salimos)
# ============================================================
LOCK_PORT = 50555
_lock_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
try:
    _lock_socket.bind(('127.0.0.1', LOCK_PORT))
    _lock_socket.listen(1)
except OSError:
    print("Ya hay un obrero Firebase ejecutándose. Saliendo.")
    sys.exit(0)

import firebase_admin
from firebase_admin import credentials, firestore
from google.cloud.firestore_v1.base_query import FieldFilter

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuración de Firebase Admin usando credenciales locales
try:
    cred = credentials.Certificate(os.path.join(BASE_DIR, 'serviceAccountKey.json'))
    firebase_admin.initialize_app(cred)
except Exception as e:
    logger.error(f"Error inicializando Firebase: {e}")
    os._exit(1)

db = firestore.client()

# Tiempos (segundos)
POLL_INTERVAL = 120        # Respaldo por si el listener se cae sin avisar
WATCHDOG_INTERVAL = 30     # Revisión de salud del listener
HEARTBEAT_INTERVAL = 600   # Latido en Firestore (144 escrituras/día)
GIT_TIMEOUT = 120

SYNC_CACHE_PATH = os.path.join(BASE_DIR, 'firestore_sync_cache.json')

# Archivos que NUNCA se tocan en una actualización OTA (datos locales de la garita)
PROTECTED_EXT = ('.db', '.xlsx', '.json', '.sqlite', '.log')
PROTECTED_DIRS = ('registros_diarios/', 'registros_fichajes/', 'registros_visitas/',
                  'backups/', 'logs/', 'temp/', '.firebase/')
CODE_JSON_ALLOWED = ('firebase.json', 'firestore.indexes.json', 'public/manifest.json')

command_queue = queue.Queue()
_enqueued_ids = set()
_enqueued_lock = threading.Lock()
_last_ok_contact = time.time()


# ============================================================
# Helpers resilientes
# ============================================================
def safe_update(doc_id, data, retries=5):
    """Actualiza un comando tolerando cortes de red (reintenta con backoff)."""
    delay = 5
    for intento in range(retries):
        try:
            db.collection('system_commands').document(doc_id).update(data)
            return True
        except Exception as e:
            logger.warning(f"[{doc_id}] No se pudo actualizar estado ({intento + 1}/{retries}): {e}")
            time.sleep(delay)
            delay = min(delay * 2, 120)
    return False


def claim_command(doc_id):
    """Pasa PENDING -> IN_PROGRESS de forma atómica. Devuelve el dict o None."""
    ref = db.collection('system_commands').document(doc_id)
    transaction = db.transaction()

    @firestore.transactional
    def _claim(tx):
        snap = ref.get(transaction=tx)
        if not snap.exists:
            return None
        data = snap.to_dict()
        if data.get('status') != 'PENDING':
            return None
        tx.update(ref, {'status': 'IN_PROGRESS', 'started_at': firestore.SERVER_TIMESTAMP})
        return data

    return _claim(transaction)


def _hash(payload):
    return hashlib.sha1(json.dumps(payload, sort_keys=True, ensure_ascii=False, default=str).encode('utf-8')).hexdigest()


def _load_sync_cache():
    try:
        with open(SYNC_CACHE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_sync_cache(cache):
    tmp = SYNC_CACHE_PATH + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(cache, f)
    os.replace(tmp, SYNC_CACHE_PATH)


def _fecha(v):
    return str(v or '').split('T')[0]


# ============================================================
# SYNC IRSA (escritura diferencial para cuidar la cuota de Firestore)
# ============================================================
def _build_firestore_docs(faos, faps):
    """Arma {ruta_doc: payload} para tramites_irsa y nominas."""
    docs = {}
    nominas = {}

    def procesar(lista, tipo):
        for t in lista:
            tid = str(t.get('id', ''))
            if not tid:
                continue
            empresa = t.get('marca') or t.get('local') or t.get('lugarTrabajo') or ''
            estado = str(t.get('estado', ''))
            personal_activo = [p for p in t.get('personal', []) if p.get('activo', True)]
            docs[f"tramites_irsa/{tid}"] = {
                'id': tid,
                'tipo': tipo,
                'empresa': empresa,
                'fecha_inicio': _fecha(t.get('fechaInicio')),
                'fecha_fin': _fecha(t.get('fechaFin')),
                'estado_codigo': estado,
                'personal': [{'nombre': p.get('nombre', ''), 'apellido': p.get('apellido', ''),
                              'dni': str(p.get('numeroDocumento', ''))} for p in personal_activo],
            }
            # Solo autorizar en puerta si está aprobado (6) o aprobado parcial (4)
            if estado not in ('4', '6'):
                continue
            for p in personal_activo:
                dni = str(p.get('numeroDocumento', '')).strip()
                if not dni:
                    continue
                venc = _fecha(t.get('fechaFin'))
                previo = nominas.get(dni)
                # Si el DNI está en varios trámites, nos quedamos con el de vencimiento más lejano
                if previo and previo['fecha_vencimiento'] >= venc:
                    continue
                nominas[dni] = {
                    'dni': dni,
                    'nombre': f"{p.get('apellido', '')} {p.get('nombre', '')}".strip(),
                    'categoria': tipo,
                    'puesto_especifico': empresa,
                    'nro_tramite': tid,
                    'fecha_vencimiento': venc,
                    'activo': True,
                }

    procesar(faos, 'FAO')
    procesar(faps, 'FAP')
    for dni, payload in nominas.items():
        docs[f"nominas/{dni}"] = payload
    return docs


def _seed_cache_from_firestore(docs):
    """Primer sync sin caché local: lee lo que ya hay en Firestore (lecturas, baratas)
    para no reescribir todo (escrituras, lo que agota la cuota diaria)."""
    cache = {}
    try:
        for col in ('tramites_irsa', 'nominas'):
            for snap in db.collection(col).stream(timeout=120):
                path = f"{col}/{snap.id}"
                if path in docs:
                    actual = snap.to_dict() or {}
                    subset = {k: actual.get(k) for k in docs[path].keys()}
                    cache[path] = _hash(subset)
        logger.info(f"Caché de sincronización inicializado desde Firestore ({len(cache)} docs).")
    except Exception as e:
        logger.warning(f"No se pudo sembrar el caché desde Firestore: {e}")
    return cache


def process_sync_irsa(doc_id, data):
    logger.info(f"[{doc_id}] Iniciando sincronización IRSA...")
    try:
        # 1. Obtener credenciales desde Firestore (si las cargaron desde la web)
        try:
            config_doc = db.collection('config').document('irsa').get()
            if config_doc.exists:
                cfg = config_doc.to_dict()
                user, pwd = cfg.get('username'), cfg.get('password')
                if user and pwd:
                    from irsa_config import save_irsa_credentials
                    save_irsa_credentials(user, pwd)
        except Exception as e:
            logger.warning(f"[{doc_id}] No se pudieron leer credenciales de la nube, uso las locales: {e}")

        # 2. Descargar desde el portal (actualiza JSON + Excel locales => offline-first)
        from irsa_client import IRSAClient
        res = IRSAClient().sincronizar_todo()

        faos, faps = [], []
        p_fao = os.path.join(BASE_DIR, 'MetadatosFAOs.json')
        p_fap = os.path.join(BASE_DIR, 'MetadatosFAPs.json')
        if os.path.exists(p_fao):
            with open(p_fao, 'r', encoding='utf-8') as f:
                faos = json.load(f)
        if os.path.exists(p_fap):
            with open(p_fap, 'r', encoding='utf-8') as f:
                faps = json.load(f)

        # 3. Subir a Firestore SOLO lo nuevo o modificado
        force = bool(data.get('force'))
        cache = {} if force else _load_sync_cache()
        docs = _build_firestore_docs(faos, faps)
        if not cache and not force:
            cache = _seed_cache_from_firestore(docs)

        pendientes = [(path, payload, _hash(payload)) for path, payload in docs.items()
                      if cache.get(path) != _hash(payload)]

        escritos = 0
        for i in range(0, len(pendientes), 400):
            chunk = pendientes[i:i + 400]
            batch = db.batch()
            for path, payload, _ in chunk:
                col, did = path.split('/', 1)
                batch.set(db.collection(col).document(did), payload, merge=True)
            batch.commit()
            for path, _, h in chunk:
                cache[path] = h
            escritos += len(chunk)
            _save_sync_cache(cache)  # guardado incremental por si se corta

        resumen = (f"FAOs aprobados: {res.get('faos_procesados', 0)}, FAPs aprobados: {res.get('faps_procesados', 0)}. "
                   f"Firestore: {escritos} cambios escritos ({len(docs) - escritos} sin cambios).")
        safe_update(doc_id, {'status': 'DONE', 'result': resumen, 'finished_at': firestore.SERVER_TIMESTAMP})
        logger.info(f"[{doc_id}] IRSA Sync OK. {resumen}")
    except Exception as e:
        logger.error(f"[{doc_id}] Error en Sincronización IRSA: {e}", exc_info=True)
        safe_update(doc_id, {'status': 'ERROR', 'result': str(e)})


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

        # Subir a firebase solo este FAO (mismo formato que el sync general)
        docs = _build_firestore_docs([detalle], [])
        cache = _load_sync_cache()
        batch = db.batch()
        for path, payload in docs.items():
            col, did = path.split('/', 1)
            batch.set(db.collection(col).document(did), payload, merge=True)
            cache[path] = _hash(payload)
        batch.commit()
        _save_sync_cache(cache)

        safe_update(doc_id, {'status': 'DONE', 'result': f'FAO {fao_id} inyectado exitosamente.'})
    except Exception as e:
        logger.error(f"[{doc_id}] Error inyectando FAO: {e}")
        safe_update(doc_id, {'status': 'ERROR', 'result': str(e)})


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

        safe_update(doc_id, {
            'status': 'DONE',
            'result': f'Excepción guardada en SQLite. {len(nomina_nueva)} personas autorizadas.'
        })
    except Exception as e:
        logger.error(f"[{doc_id}] Error guardando excepciones: {e}")
        safe_update(doc_id, {'status': 'ERROR', 'result': str(e)})


# ============================================================
# OTA: actualiza SOLO código, nunca los datos locales
# ============================================================
def _git(*args):
    r = subprocess.run(['git', *args], cwd=BASE_DIR, capture_output=True, text=True,
                       timeout=GIT_TIMEOUT, encoding='utf-8', errors='replace')
    if r.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} falló: {(r.stderr or r.stdout).strip()}")
    return r.stdout.strip()


def _is_protected(path):
    p = path.replace('\\', '/')
    if p in CODE_JSON_ALLOWED:
        return False
    if p == 'serviceAccountKey.json' or p.startswith(PROTECTED_DIRS):
        return True
    return p.lower().endswith(PROTECTED_EXT)


def process_ota_update(doc_id):
    logger.info(f"[{doc_id}] Iniciando actualización remota (OTA)...")
    try:
        branch = _git('rev-parse', '--abbrev-ref', 'HEAD')
        _git('fetch', 'origin', branch)
        remote = f'origin/{branch}'

        old_head = _git('rev-parse', 'HEAD')
        new_head = _git('rev-parse', remote)
        if old_head == new_head:
            safe_update(doc_id, {'status': 'DONE', 'result': f'Ya está actualizado ({branch} @ {new_head[:7]}).'})
            return

        cambios = []
        for line in _git('diff', '--name-status', '--no-renames', 'HEAD', remote).splitlines():
            parts = line.split('\t')
            if len(parts) >= 2:
                cambios.append((parts[0][0], parts[-1]))

        codigo = [(s, p) for s, p in cambios if not _is_protected(p)]
        omitidos = [p for s, p in cambios if _is_protected(p)]

        a_traer = [p for s, p in codigo if s != 'D']
        a_borrar = [p for s, p in codigo if s == 'D']

        # Traer archivos de código en tandas (límite de largo de línea en Windows)
        for i in range(0, len(a_traer), 50):
            _git('checkout', remote, '--', *a_traer[i:i + 50])
        for p in a_borrar:
            full = os.path.join(BASE_DIR, p)
            if os.path.exists(full):
                os.remove(full)

        # Mueve HEAD e índice sin tocar el árbol de trabajo (los datos locales quedan intactos)
        _git('reset', '--mixed', '-q', remote)

        resultado = (f"Actualizado {old_head[:7]} -> {new_head[:7]}. "
                     f"{len(a_traer) + len(a_borrar)} archivos de código aplicados"
                     + (f", {len(omitidos)} archivos de datos preservados" if omitidos else "") + ".")
        logger.info(f"[{doc_id}] {resultado}")

        safe_update(doc_id, {'status': 'DONE', 'result': resultado + ' Reiniciando obrero...',
                             'finished_at': firestore.SERVER_TIMESTAMP})
        time.sleep(2)
        restart_self()
    except Exception as e:
        logger.error(f"[{doc_id}] Error en actualización OTA: {e}", exc_info=True)
        safe_update(doc_id, {'status': 'ERROR', 'result': str(e)})


def restart_self():
    logger.info("Reiniciando proceso del obrero...")
    try:
        _lock_socket.close()
    except Exception:
        pass
    if os.environ.get('FIREBASE_WORKER_SUPERVISED') == '1':
        os._exit(0)  # app.py lo vuelve a levantar en <30s con el código nuevo
    # Ejecutado a mano: lanzamos una copia nueva y terminamos esta
    subprocess.Popen([sys.executable, os.path.join(BASE_DIR, 'firebase_worker.py')], cwd=BASE_DIR)
    os._exit(0)


# ============================================================
# Cola de comandos
# ============================================================
def enqueue(doc_id):
    with _enqueued_lock:
        if doc_id in _enqueued_ids:
            return
        _enqueued_ids.add(doc_id)
    command_queue.put(doc_id)


def command_processor():
    while True:
        doc_id = command_queue.get()
        try:
            data = claim_command(doc_id)
            if data is None:
                continue  # Ya lo tomó otro o no está PENDING
            action = data.get('action')
            logger.info(f"Comando: {action} (ID: {doc_id})")

            if action == 'SYNC_IRSA':
                process_sync_irsa(doc_id, data)
            elif action == 'UPDATE_OTA':
                process_ota_update(doc_id)
            elif action == 'FORCE_SYNC_FAO':
                fao_id = data.get('fao_id')
                if fao_id:
                    process_force_sync_fao(doc_id, fao_id)
                else:
                    safe_update(doc_id, {'status': 'ERROR', 'result': 'fao_id no especificado'})
            elif action == 'SAVE_EXCEPCIONES':
                process_save_excepciones(doc_id, data.get('nomina', []), data.get('empresa'),
                                         data.get('vigencia_desde'), data.get('vigencia_hasta'))
            else:
                safe_update(doc_id, {'status': 'ERROR', 'result': f'Comando desconocido: {action}'})
        except Exception as e:
            # Probablemente sin internet: liberamos para reintentar en el próximo polling
            logger.error(f"[{doc_id}] Error procesando comando (se reintentará): {e}")
        finally:
            with _enqueued_lock:
                _enqueued_ids.discard(doc_id)
            command_queue.task_done()


def on_command_snapshot(doc_snapshot, changes, read_time):
    global _last_ok_contact
    _last_ok_contact = time.time()
    # El callback solo encola: nunca bloquear el hilo del listener con trabajo largo
    for change in changes:
        if change.type.name in ('ADDED', 'MODIFIED'):
            if (change.document.to_dict() or {}).get('status') == 'PENDING':
                enqueue(change.document.id)


def pending_query():
    return db.collection('system_commands').where(filter=FieldFilter('status', '==', 'PENDING'))


def start_listener():
    return pending_query().on_snapshot(on_command_snapshot)


def watch_is_dead(watch):
    if watch is None:
        return True
    # Atributos internos de google-cloud-firestore Watch
    if getattr(watch, '_closed', False):
        return True
    rpc = getattr(watch, '_rpc', None)
    if rpc is not None and hasattr(rpc, 'is_active') and not rpc.is_active:
        return True
    return False


def poll_pending():
    """Respaldo: si el listener se perdió algo (corte de red), lo levantamos acá."""
    global _last_ok_contact
    docs = list(pending_query().limit(20).stream(timeout=30))
    _last_ok_contact = time.time()
    for d in docs:
        enqueue(d.id)
    return len(docs)


def heartbeat():
    db.collection('system_status').document('worker').set({
        'last_seen': firestore.SERVER_TIMESTAMP,
        'host': socket.gethostname(),
        'pid': os.getpid(),
    }, merge=True)


def main():
    threading.Thread(target=command_processor, daemon=True, name='command-processor').start()

    watch = None
    last_poll = 0
    last_heartbeat = 0

    logger.info("👷‍♂️ Obrero Python iniciado.")
    logger.info("Escuchando órdenes de la nube (SYNC_IRSA, UPDATE_OTA, FORCE_SYNC_FAO, SAVE_EXCEPCIONES)...")

    while True:
        now = time.time()
        try:
            if watch_is_dead(watch):
                if watch is not None:
                    logger.warning("Listener de Firestore caído. Reconectando...")
                    try:
                        watch.unsubscribe()
                    except Exception:
                        pass
                watch = start_listener()
                logger.info("Listener de Firestore conectado.")
                last_poll = 0  # forzar un polling inmediato tras reconectar

            if now - last_poll >= POLL_INTERVAL:
                n = poll_pending()
                if n:
                    logger.info(f"Polling: {n} comando(s) pendiente(s) encontrados.")
                last_poll = now

            if now - last_heartbeat >= HEARTBEAT_INTERVAL:
                heartbeat()
                last_heartbeat = now

        except Exception as e:
            # Sin internet / cuota excedida: seguimos vivos y reintentamos
            logger.warning(f"Sin conexión con Firebase ({type(e).__name__}: {e}). Reintento en {WATCHDOG_INTERVAL}s.")
            # Si llevamos mucho sin contacto, recreamos el listener desde cero
            if time.time() - _last_ok_contact > 600 and watch is not None:
                try:
                    watch.unsubscribe()
                except Exception:
                    pass
                watch = None

        time.sleep(WATCHDOG_INTERVAL)


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        logger.info("Obrero detenido.")
