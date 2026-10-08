import os
from datetime import datetime
import smtplib
import shutil
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import base64
import threading
import time
import webview
import pandas as pd
from flask import (
    Flask,
    jsonify,
    request,
    render_template,
    redirect,
    url_for,
    session,
    send_from_directory,
    send_file,
    after_this_request,
)
from io import BytesIO

# Local application imports
import data_manager
from config import (
    SECRET_KEY,
    ADMIN_USERNAME,
    ADMIN_PASSWORD,
    EMAIL_SENDER,
    EMAIL_PASSWORD,
    EMAIL_RECEIVER,
    SMTP_SERVER,
    SMTP_PORT,
    REGISTROS_DIARIOS_DIR,
    REGISTROS_FICHAJES_DIR,
    EXCEL_FAP,
    EXCEL_FAO,
    EXCEL_EXCEPCIONES,
    EXCEL_NOMINAS,
)
from data_manager import (
    cargar_autorizaciones,
    get_nominas_agrupadas,
    delete_nomina_by_criteria,
    get_nomina_detalle_by_criteria,
    recargar_cache_nominas_persistentes,
    procesar_nomina_texto,
    generar_reporte_consolidado,
)
from access_manager import (
    registrar_fichaje,
    verificar_dni,
    personas_adentro,
)
from logger_config import get_logger
from utils import (
    validar_archivo_fap,
    validar_archivo_fao,
    crear_backup,
    limpiar_backups_antiguos,
)


from flask_cors import CORS

app = Flask(__name__)
CORS(app) # Permitir llamadas locales desde Firebase Hosting
app.secret_key = SECRET_KEY

# Logger para este módulo
logger = get_logger(__name__)


# ================= SUPERVISOR DEL OBRERO FIREBASE =================
# Mantiene vivo firebase_worker.py (comandos de la web: SYNC_IRSA, UPDATE_OTA, etc.).
# El obrero tiene bloqueo de instancia única, así que lanzarlo de más es inofensivo.
def _supervisar_obrero_firebase():
    import subprocess
    import sys as _sys
    base = os.path.dirname(os.path.abspath(__file__))
    worker_path = os.path.join(base, 'firebase_worker.py')
    proc = None
    while True:
        try:
            if proc is None or proc.poll() is not None:
                proc = subprocess.Popen(
                    [_sys.executable, worker_path], cwd=base,
                    env={**os.environ, 'FIREBASE_WORKER_SUPERVISED': '1'},
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
                )
        except Exception as e:
            logger.error(f"Supervisor del obrero Firebase: {e}")
        time.sleep(30)


if not getattr(__import__('sys'), 'frozen', False) and os.path.exists(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'firebase_worker.py')):
    threading.Thread(target=_supervisar_obrero_firebase, daemon=True, name='firebase-worker-supervisor').start()

class Api:
    def __init__(self):
        self._window = None

    def set_window(self, window):
        self._window = window

    def save_file_dialog(self, data):
        if not self._window:
            return {'success': False, 'message': 'Window not set'}

        try:
            filename = data['filename']
            content_base64 = data['content']
            
            content_bytes = base64.b64decode(content_base64)

            result = self._window.create_file_dialog(webview.SAVE_DIALOG, directory='/', save_filename=filename)

            if result:
                filepath = result[0] if isinstance(result, (list, tuple)) else result
                with open(filepath, 'wb') as f:
                    f.write(content_bytes)
                return {'success': True, 'message': f'Archivo guardado en {filepath}'}
            else:
                return {'success': False, 'message': 'Guardado cancelado por el usuario.'}

        except Exception as e:
            return {'success': False, 'message': f'Error al guardar el archivo: {str(e)}'}

api = Api()

os.makedirs(REGISTROS_DIARIOS_DIR, exist_ok=True)
os.makedirs(REGISTROS_FICHAJES_DIR, exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/login')
def login_page():
    return render_template('login.html')

@app.route('/verificar_dni', methods=['POST'])
def api_verificar_dni():
    data = request.get_json()
    scanner_data = data.get('scanner_data')
    mode = data.get('mode')
    
    if not scanner_data or not mode:
        return jsonify({'acceso': 'DENEGADO', 'mensaje': 'Faltan datos en la solicitud.'}), 400
        
    resultado = verificar_dni(scanner_data, mode)
    return jsonify(resultado)

@app.route('/registrar_fichaje', methods=['POST'])
def api_registrar_fichaje():
    data = request.get_json()
    scanner_data = data.get('scanner_data')
    mode = data.get('mode')

    if not scanner_data or not mode:
        return jsonify({'acceso': 'DENEGADO', 'mensaje': 'Faltan datos.'}), 400

    resultado = registrar_fichaje(scanner_data, mode)
    return jsonify(resultado)

@app.route('/get_daily_records')
def get_daily_records():
    fecha_actual_str = datetime.now().strftime('%Y-%m-%d')
    nombre_archivo = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_actual_str}.xlsx')

    if not os.path.exists(nombre_archivo):
        return jsonify({'success': True, 'records': [], 'message': 'No hay registros para hoy.'})

    try:
        df = pd.read_excel(nombre_archivo).fillna('')
        records = df.to_dict('records')
        return jsonify({'success': True, 'records': records})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error al leer registros: {e}'}), 500

@app.route('/get_dynamic_stats')
def get_dynamic_stats():
    fecha_actual_str = datetime.now().strftime('%Y-%m-%d')
    nombre_archivo = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_actual_str}.xlsx')
    
    permitidos = 0
    rechazados = 0

    if os.path.exists(nombre_archivo):
        try:
            df = pd.read_excel(nombre_archivo)
            permitidos = df[df['Resultado'] == 'VERDE'].shape[0]
            rechazados = df[df['Resultado'] == 'ROJO'].shape[0]
        except Exception:
            pass 

    return jsonify({
        'total_adentro': len(personas_adentro),
        'permitidos': permitidos,
        'rechazados': rechazados
    })

@app.route('/admin')
def admin_page():
    if 'logged_in' in session:
        return render_template('admin.html')
    return redirect(url_for('login_page'))

@app.route('/perform_login', methods=['POST'])
def perform_login():
    data = request.get_json()
    if data.get('username') == ADMIN_USERNAME and data.get('password') == ADMIN_PASSWORD:
        session['logged_in'] = True
        return jsonify({'success': True})
    return jsonify({'success': False, 'message': 'Credenciales incorrectas.'})

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('home'))

@app.route('/api/admin/create_user', methods=['POST'])
def api_create_user():
    data = request.get_json()
    email = data.get('email')
    password = data.get('password')
    rol = data.get('rol')
    
    if not email or not password or not rol:
        return jsonify({'success': False, 'message': 'Faltan datos'})
        
    try:
        from firebase_admin import auth, firestore
        import firebase_admin
        from firebase_admin import credentials
        
        # Initialize if not already initialized
        if not firebase_admin._apps:
            cred = credentials.Certificate('serviceAccountKey.json')
            firebase_admin.initialize_app(cred)
            
        db = firestore.client()
        
        user = auth.create_user(email=email, password=password)
        db.collection('usuarios').document(user.uid).set({
            'email': email,
            'rol': rol,
            'activo': True,
            'centro_id': 'al_oeste'
        })
        return jsonify({'success': True, 'message': f'Usuario {email} creado con rol {rol}'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403

    fap_file = request.files.get('fapFile')
    fao_file = request.files.get('faoFile')
    
    mensajes = []
    errores = []
    
    # Procesar archivo FAP
    if fap_file:
        try:
            # Guardar temporalmente
            temp_fap = EXCEL_FAP + '.temp'
            fap_file.save(temp_fap)
            
            # Validar archivo
            valido, mensaje = validar_archivo_fap(temp_fap)
            
            if valido:
                # Crear backup del archivo actual si existe
                if os.path.exists(EXCEL_FAP):
                    exito_backup, ruta_backup = crear_backup(EXCEL_FAP, tipo='pre-update')
                    if exito_backup:
                        logger.info(f"Backup de FAP creado antes de actualizar: {os.path.basename(ruta_backup)}")
                
                # Reemplazar con el nuevo
                shutil.move(temp_fap, EXCEL_FAP)
                mensajes.append(f"FAP: {mensaje}")
                logger.info(f"Archivo FAP actualizado correctamente")
            else:
                # Eliminar temporal si no es válido
                if os.path.exists(temp_fap):
                    os.remove(temp_fap)
                errores.append(f"FAP: {mensaje}")
                logger.warning(f"Archivo FAP rechazado: {mensaje}")
                
        except Exception as e:
            errores.append(f"FAP: Error al procesar - {str(e)}")
            logger.error(f"Error al procesar archivo FAP: {e}", exc_info=True)
    
    # Procesar archivo FAO
    if fao_file:
        try:
            # Guardar temporalmente
            temp_fao = EXCEL_FAO + '.temp'
            fao_file.save(temp_fao)
            
            # Validar archivo
            valido, mensaje = validar_archivo_fao(temp_fao)
            
            if valido:
                # Crear backup del archivo actual si existe
                if os.path.exists(EXCEL_FAO):
                    exito_backup, ruta_backup = crear_backup(EXCEL_FAO, tipo='pre-update')
                    if exito_backup:
                        logger.info(f"Backup de FAO creado antes de actualizar: {os.path.basename(ruta_backup)}")
                
                # Reemplazar con el nuevo
                shutil.move(temp_fao, EXCEL_FAO)
                mensajes.append(f"FAO: {mensaje}")
                logger.info(f"Archivo FAO actualizado correctamente")
            else:
                # Eliminar temporal si no es válido
                if os.path.exists(temp_fao):
                    os.remove(temp_fao)
                errores.append(f"FAO: {mensaje}")
                logger.warning(f"Archivo FAO rechazado: {mensaje}")
                
        except Exception as e:
            errores.append(f"FAO: Error al procesar - {str(e)}")
            logger.error(f"Error al procesar archivo FAO: {e}", exc_info=True)
    
    # Limpiar backups antiguos (DESACTIVADO - los backups se mantienen permanentemente)
    # limpiar_backups_antiguos(dias=30, tipo='pre-update')
    
    # Recargar datos si hubo al menos un archivo válido
    if mensajes:
        cargar_autorizaciones()
    
    # Preparar respuesta
    if errores and not mensajes:
        return jsonify({
            'success': False,
            'message': 'Errores en validación:\n' + '\n'.join(errores)
        })
    elif errores and mensajes:
        return jsonify({
            'success': True,
            'message': '\n'.join(mensajes) + '\n\nAdvertencias:\n' + '\n'.join(errores)
        })
    else:
        return jsonify({
            'success': True,
            'message': '\n'.join(mensajes) if mensajes else 'Archivos procesados correctamente'
        })
@app.route('/agregar_excepcion', methods=['POST'])
def agregar_excepcion():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    
    data = request.get_json()
    
    columnas_excepcion = ['Numero', 'Nombre Completo', 'Local', 'Quien Autoriza', 'Fecha de Alta', 'Vigencia']
    
    df_actual = pd.DataFrame(columns=columnas_excepcion)
    
    try:
        if os.path.exists(EXCEL_EXCEPCIONES):
            df_leido = pd.read_excel(EXCEL_EXCEPCIONES)
            for col in columnas_excepcion:
                if col not in df_leido.columns:
                    df_leido[col] = pd.NA
            df_actual = df_leido[columnas_excepcion].copy()
            
    except Exception as e:
        logger.info(f"No se pudo leer {EXCEL_EXCEPCIONES} (puede ser la primera vez). Se creará uno nuevo. Error: {e}")
        df_actual = pd.DataFrame(columns=columnas_excepcion)

    if 'Numero' in df_actual.columns:
        df_actual['Numero'] = df_actual['Numero'].astype(str)
    
    dni_nuevo = str(data['dni'])
    nombre_nuevo = f"{data['nombre']} {data['apellido']}"
    local_nuevo = data['local']
    autoriza_nuevo = data['autoriza']
    vigencia_nueva = data.get('vigencia')
    fecha_alta_nueva = datetime.now().strftime('%Y-%m-%d')

    if vigencia_nueva:
        vigencia_dt = pd.to_datetime(vigencia_nueva, errors='coerce')
    else:
        vigencia_dt = pd.NaT

    idx = df_actual[df_actual['Numero'] == dni_nuevo].index

    if not idx.empty:
        df_actual.loc[idx[0], 'Nombre Completo'] = nombre_nuevo
        df_actual.loc[idx[0], 'Local'] = local_nuevo
        df_actual.loc[idx[0], 'Quien Autoriza'] = autoriza_nuevo
        df_actual.loc[idx[0], 'Fecha de Alta'] = fecha_alta_nueva
        df_actual.loc[idx[0], 'Vigencia'] = vigencia_dt
        mensaje_exito = 'Excepción actualizada.'
    else:
        nuevo_registro = pd.DataFrame([{
            'Numero': dni_nuevo,
            'Nombre Completo': nombre_nuevo,
            'Local': local_nuevo,
            'Quien Autoriza': autoriza_nuevo,
            'Fecha de Alta': fecha_alta_nueva,
            'Vigencia': vigencia_dt
        }], columns=columnas_excepcion)
        df_actual = pd.concat([df_actual, nuevo_registro], ignore_index=True)
        mensaje_exito = 'Excepción agregada.'
    
    try:
        df_actual.to_excel(EXCEL_EXCEPCIONES, index=False, columns=columnas_excepcion)
        
        # También guardar en SQLite para que el escáner y el dashboard lo vean inmediatamente
        try:
            from database import get_db_connection
            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM autorizaciones WHERE dni=? AND tipo_permiso='EXCEPCION'", (dni_nuevo,))
                
                apellido_split = ""
                nombre_split = nombre_nuevo
                if " " in nombre_nuevo:
                    partes = nombre_nuevo.split(" ", 1)
                    nombre_split = partes[1]
                    apellido_split = partes[0]
                    
                vigencia_str = vigencia_dt.strftime('%Y-%m-%d') if pd.notna(vigencia_dt) else ""
                
                cursor.execute('''
                    INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, local, fecha_inicio, fecha_fin, activo)
                    VALUES (?, ?, ?, 'EXCEPCION', ?, ?, ?, 1)
                ''', (dni_nuevo, nombre_split, apellido_split, local_nuevo, fecha_alta_nueva, vigencia_str))
        except Exception as sqlite_e:
            logger.error(f"Error al guardar excepción en SQLite: {sqlite_e}")
            
        cargar_autorizaciones()
        return jsonify({'success': True, 'message': mensaje_exito})
        
    except Exception as e:
        logger.error(f"Error crítico al guardar excepción en Excel: {e}", exc_info=True)
        return jsonify({'success': False, 'message': f'Error al guardar: {e}'})

@app.route('/parse_nomina', methods=['POST'])
def parse_nomina():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403

    data = request.get_json()
    texto_pegado = data.get('texto_pegado', '')

    if not texto_pegado:
        return jsonify({'success': False, 'message': 'El texto de la nómina no puede estar vacío.'})

    personas_procesadas = procesar_nomina_texto(texto_pegado)

    if not personas_procesadas:
        return jsonify({'success': False, 'message': 'No se pudo interpretar ninguna persona. Revise el formato del texto.'})

    nomina_para_frontend = [
        {'dni': p['DNI'], 'apellido': p['Apellido'], 'nombre': p['Nombre']}
        for p in personas_procesadas
    ]

    return jsonify({'success': True, 'nomina': nomina_para_frontend})

@app.route('/save_nomina', methods=['POST'])
def save_nomina():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    
    data = request.get_json()
    nomina_nueva = data.get('nomina', [])
    empresa = data.get('empresa')
    vigencia_desde = data.get('vigencia_desde')
    vigencia_hasta = data.get('vigencia_hasta')
    is_update = data.get('is_update', False)
    original_empresa = data.get('original_empresa')
    original_vigencia_desde = data.get('original_vigencia_desde')

    if not all([nomina_nueva, empresa, vigencia_desde, vigencia_hasta]):
        return jsonify({'success': False, 'message': 'Faltan datos para guardar la nómina.'}), 400

    try:
        # Leemos directamente del archivo para asegurar que tenemos la última versión
        df_total = pd.DataFrame()
        if os.path.exists(EXCEL_NOMINAS):
            df_total = pd.read_excel(EXCEL_NOMINAS)
        
        # Si es una actualización, eliminamos la nómina anterior completa.
        if is_update and original_empresa and original_vigencia_desde:
            try:
                # Normalizamos las fechas para una comparación segura
                fecha_desde_obj = pd.to_datetime(original_vigencia_desde, dayfirst=True, errors='coerce').normalize()
                df_total['Vigencia Desde'] = pd.to_datetime(df_total['Vigencia Desde'], errors='coerce').dt.normalize()
                
                # Condición para MANTENER todo lo que NO coincida con la nómina a actualizar
                condition = ~((df_total['Empresa'] == original_empresa) & (df_total['Vigencia Desde'] == fecha_desde_obj))
                df_total = df_total[condition].copy()
            except Exception as e:
                # Si hay un error en la conversión de fechas, es mejor no continuar para no corromper los datos
                logger.error(f"Error al procesar fechas durante la actualización: {e}")
                return jsonify({'success': False, 'message': f'Error al procesar fechas: {e}'}), 500

        # Preparamos el nuevo DataFrame
        df_nueva = pd.DataFrame(nomina_nueva)
        # Nos aseguramos de que no haya duplicados DENTRO de la nueva nómina
        df_nueva = df_nueva.drop_duplicates(subset=['dni'])
        
        df_nueva['Empresa'] = empresa
        df_nueva['Vigencia Desde'] = pd.to_datetime(vigencia_desde, dayfirst=True, errors='coerce')
        df_nueva['Vigencia Hasta'] = pd.to_datetime(vigencia_hasta, dayfirst=True, errors='coerce')
        
        df_nueva.rename(columns={'dni': 'DNI', 'apellido': 'Apellido', 'nombre': 'Nombre'}, inplace=True)
        
        # Juntamos los datos viejos (ya filtrados) con los nuevos
        # Usar ignore_index=True es CRUCIAL para evitar el error de "Reindexing"
        df_final = pd.concat([df_total, df_nueva], ignore_index=True)

        # Guardamos el resultado final
        df_final.to_excel(EXCEL_NOMINAS, index=False)
        
        recargar_cache_nominas_persistentes()
        
        mensaje = 'Nómina actualizada correctamente.' if is_update else 'Nómina guardada correctamente.'
        return jsonify({'success': True, 'message': mensaje})

    except Exception as e:
        logger.error(f"Error crítico al guardar la nómina: {e}", exc_info=True)
        return jsonify({'success': False, 'message': f'Error del servidor: {e}'}), 500

@app.route('/get_nominas_guardadas', methods=['GET'])
def get_nominas_guardadas_route():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    try:
        nominas = get_nominas_agrupadas()
        return jsonify({'success': True, 'nominas': nominas})
    except Exception as e:
        logger.error(f"Error al obtener nóminas guardadas: {e}")
        return jsonify({'success': False, 'message': f'Error del servidor: {e}'}), 500
@app.route('/delete_nomina', methods=['POST'])
def delete_nomina_route():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403

    data = request.get_json()
    empresa = data.get('empresa')
    vigencia = data.get('vigencia')

    if not empresa or not vigencia:
        return jsonify({'success': False, 'message': 'Faltan datos para eliminar la nómina.'}), 400

    try:
        eliminado = delete_nomina_by_criteria(empresa, vigencia)
        if not eliminado:
            return jsonify({'success': False, 'message': 'No se encontró la nómina especificada.'}), 404
        
        return jsonify({'success': True, 'message': 'Nómina eliminada correctamente.'})
    except Exception as e:
        logger.error(f"Error crítico al eliminar la nómina: {e}", exc_info=True)
        return jsonify({'success': False, 'message': f'Error del servidor: {e}'}), 500

@app.route('/get_nomina_detalle', methods=['POST'])
def get_nomina_detalle_route():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403

    data = request.get_json()
    empresa = data.get('empresa')
    vigencia = data.get('vigencia')
    filtro_dni = data.get('filtro_dni')
    filtro_nombre = data.get('filtro_nombre')
    filtro_apellido = data.get('filtro_apellido')

    if not empresa or not vigencia:
        return jsonify({'success': False, 'message': 'Faltan datos para obtener el detalle.'}), 400

    try:
        detalle = get_nomina_detalle_by_criteria(empresa, vigencia, filtro_dni, filtro_nombre, filtro_apellido)
        if detalle is None:
            return jsonify({'success': False, 'message': 'No se encontró la nómina para editar.'}), 404
        return jsonify({'success': True, 'detalle': detalle})
    except Exception as e:
        logger.error(f"Error al obtener el detalle de la nómina: {e}")
        return jsonify({'success': False, 'message': f'Error del servidor: {e}'}), 500

@app.route('/send_report_email', methods=['POST'])
def send_report_email():

    if 'logged_in' not in session:

        return jsonify({'success': False, 'message': 'No autorizado'}), 403

    fecha_hoy = datetime.now().strftime('%Y-%m-%d')

    archivo_accesos = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_hoy}.xlsx')

    archivo_fichajes = os.path.join(REGISTROS_FICHAJES_DIR, f'registros_fichaje_{fecha_hoy}.xlsx')

    if not os.path.exists(archivo_accesos) and not os.path.exists(archivo_fichajes):

        return jsonify({'success': False, 'message': 'No hay reportes para enviar hoy.'})

    try:

        msg = MIMEMultipart()

        msg['From'] = EMAIL_SENDER

        msg['To'] = EMAIL_RECEIVER

        msg['Subject'] = f"Reportes de Control de Acceso y Fichajes - {fecha_hoy}"

        msg.attach(MIMEText("Se adjuntan los reportes del día.", 'plain'))

        for archivo in [archivo_accesos, archivo_fichajes]:

            if os.path.exists(archivo):

                with open(archivo, "rb") as f:

                    part = MIMEBase('application', 'octet-stream')

                    part.set_payload(f.read())

                encoders.encode_base64(part)

                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(archivo)}')

                msg.attach(part)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)

        server.starttls()

        server.login(EMAIL_SENDER, EMAIL_PASSWORD)

        server.send_message(msg)

        server.quit()

        return jsonify({'success': True, 'message': 'Reporte enviado exitosamente'})

    except Exception as e:

        return jsonify({'success': False, 'message': f'Error al enviar email: {e}'})

@app.route('/descargar_reporte_fechas', methods=['GET'])
def descargar_reporte_fechas():
    """Descarga reporte de accesos para un período específico"""
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    
    try:
        fecha_desde = request.args.get('desde', '')
        fecha_hasta = request.args.get('hasta', '')
        
        # Si no hay fechas, usar hoy
        if not fecha_desde and not fecha_hasta:
            hoy = datetime.now().strftime('%Y-%m-%d')
            fecha_desde = hoy
            fecha_hasta = hoy
        elif not fecha_hasta:
            fecha_hasta = fecha_desde
        elif not fecha_desde:
            fecha_desde = fecha_hasta
        
        # Convertir a datetime
        desde = datetime.strptime(fecha_desde, '%Y-%m-%d')
        hasta = datetime.strptime(fecha_hasta, '%Y-%m-%d')
        
        # Generar reporte consolidado para el período
        registros_consolidados = []
        
        fecha_actual = desde
        while fecha_actual <= hasta:
            fecha_str = fecha_actual.strftime('%Y-%m-%d')
            archivo_dia = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_str}.xlsx') # Corrected filename
            
            if os.path.exists(archivo_dia):
                try:
                    df_dia = pd.read_excel(archivo_dia)
                    df_dia['Fecha'] = fecha_str
                    registros_consolidados.append(df_dia)
                except Exception as e:
                    logger.warning(f"Error al leer {archivo_dia}: {e}")
            
            fecha_actual += pd.Timedelta(days=1)
        
        if not registros_consolidados:
            return jsonify({
                'success': False,
                'message': f'No hay registros para el período {fecha_desde} - {fecha_hasta}'
            }), 404
        
        # Consolidar todos los registros
        df_consolidado = pd.concat(registros_consolidados, ignore_index=True)
        
        # Reordenar columnas para poner Fecha primero
        cols = df_consolidado.columns.tolist()
        if 'Fecha' in cols:
            cols = ['Fecha'] + [c for c in cols if c != 'Fecha']
            df_consolidado = df_consolidado[cols]
        
        # Crear archivo temporal
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_consolidado.to_excel(writer, index=False, sheet_name='Registros')
        
        output.seek(0)
        
        # Nombre del archivo
        if fecha_desde == fecha_hasta:
            filename = f'reporte_accesos_{fecha_desde}.xlsx'
        else:
            filename = f'reporte_accesos_{fecha_desde}_a_{fecha_hasta}.xlsx'
        
        logger.info(f"Descargando reporte de accesos: {filename}")
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        logger.error(f"Error al generar reporte de accesos: {e}", exc_info=True)
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500


@app.route('/descargar_fichajes_fechas', methods=['GET'])
def descargar_fichajes_fechas():
    """Descarga reporte de fichajes para un período específico"""
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    
    try:
        fecha_desde = request.args.get('desde', '')
        fecha_hasta = request.args.get('hasta', '')
        
        # Si no hay fechas, usar hoy
        if not fecha_desde and not fecha_hasta:
            hoy = datetime.now().strftime('%Y-%m-%d')
            fecha_desde = hoy
            fecha_hasta = hoy
        elif not fecha_hasta:
            fecha_hasta = fecha_desde
        elif not fecha_desde:
            fecha_desde = fecha_hasta
        
        # Convertir a datetime
        desde = datetime.strptime(fecha_desde, '%Y-%m-%d')
        hasta = datetime.strptime(fecha_hasta, '%Y-%m-%d')
        
        # Generar reporte consolidado para el período
        fichajes_consolidados = []
        
        fecha_actual = desde
        while fecha_actual <= hasta:
            fecha_str = fecha_actual.strftime('%Y-%m-%d')
            archivo_dia = os.path.join(REGISTROS_FICHAJES_DIR, f'registros_fichaje_{fecha_str}.xlsx') # Corrected filename
            
            if os.path.exists(archivo_dia):
                try:
                    df_dia = pd.read_excel(archivo_dia)
                    fichajes_consolidados.append(df_dia)
                except Exception as e:
                    logger.warning(f"Error al leer {archivo_dia}: {e}")
            
            fecha_actual += pd.Timedelta(days=1)
        
        if not fichajes_consolidados:
            return jsonify({
                'success': False,
                'message': f'No hay fichajes para el período {fecha_desde} - {fecha_hasta}'
            }), 404
        
        # Consolidar todos los fichajes
        df_consolidado = pd.concat(fichajes_consolidados, ignore_index=True)
        
        # Crear archivo temporal
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_consolidado.to_excel(writer, index=False, sheet_name='Fichajes')
        
        output.seek(0)
        
        # Nombre del archivo
        if fecha_desde == fecha_hasta:
            filename = f'fichajes_{fecha_desde}.xlsx'
        else:
            filename = f'fichajes_{fecha_desde}_a_{fecha_hasta}.xlsx'
        
        logger.info(f"Descargando reporte de fichajes: {filename}")
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        logger.error(f"Error al generar reporte de fichajes: {e}", exc_info=True)
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500


@app.route('/descargar_reporte_diario')
def descargar_reporte_diario():
    if 'logged_in' not in session:
        return redirect(url_for('login_page'))

    # Llama a la nueva función para generar el reporte consolidado
    temp_path, _ = generar_reporte_consolidado()
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    download_filename = f'Reporte de Accesos diarios {fecha_hoy}.xlsx'

    if not temp_path:
        return "No hay reporte de accesos para hoy.", 404

    try:
        with open(temp_path, 'rb') as f:
            file_data = f.read()
    finally:
        try:
            os.remove(temp_path)
            logger.info(f"Archivo temporal '{temp_path}' eliminado.")
        except Exception as e:
            logger.error(f"Error al eliminar archivo temporal: {e}")

    return send_file(
        BytesIO(file_data),
        download_name=download_filename,
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )



@app.route('/descargar_reporte_fichajes')
def descargar_reporte_fichajes():
    if 'logged_in' not in session:
        return redirect(url_for('login_page'))
    
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    filename = f'Reporte de fichaje diario {fecha_hoy}.xlsx'
    filepath = os.path.join(REGISTROS_FICHAJES_DIR, f'registros_fichaje_{fecha_hoy}.xlsx')
    
    output = BytesIO()

    if os.path.exists(filepath):
        try:
            with open(filepath, 'rb') as f:
                output.write(f.read())
        except Exception as e:
            logger.error(f"Error al leer el archivo de fichajes existente: {e}")
            # Opcional: podrías querer enviar un archivo de error o un 500
            return "Error al procesar el reporte de fichajes.", 500
    else:
        # Si el archivo no existe, crea uno vacío con las cabeceras correctas.
        columnas = ['DNI', 'Nombre y Apellido', 'Fecha', 'Hora_Entrada', 'Hora_Salida']
        df_vacio = pd.DataFrame(columns=columnas)
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_vacio.to_excel(writer, index=False, sheet_name='Fichajes')

    output.seek(0)
    
    return send_file(
        output,
        download_name=filename,
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

import threading
import time

# ... (all the existing imports and code, including Api class and routes) ...

def run_flask():
    # use_reloader=False is important to prevent the app from running twice
    app.run(port=5050, use_reloader=False) 



# ================= RECOVERED ENDPOINTS =================

@app.route('/dashboard')
def dashboard():
    if 'logged_in' not in session:
        return redirect(url_for('login_page'))
    return render_template('dashboard.html')

@app.route('/api/dashboard/stats')
def api_dashboard_stats():
    # Solo un dummy o la lógica que use la bd SQLite
    try:
        from db_queries import get_db_connection
        hoy = datetime.now().strftime('%Y-%m-%d')
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT COUNT(*) as c FROM registros_accesos WHERE fecha = ? AND evento = 'ENTRADA'", (hoy,))
            entradas = c.fetchone()['c']
            c.execute("SELECT COUNT(*) as c FROM registros_accesos WHERE fecha = ? AND evento = 'SALIDA'", (hoy,))
            salidas = c.fetchone()['c']
            c.execute("SELECT COUNT(*) as c FROM registros_accesos WHERE fecha = ? AND evento = 'ENTRADA_RECHAZADA'", (hoy,))
            rechazos = c.fetchone()['c']
            c.execute("SELECT COUNT(*) as c FROM estado_adentro")
            adentro = c.fetchone()['c']
            c.execute("SELECT COUNT(*) as c FROM registros_accesos WHERE fecha = ? AND tipo_permiso = 'Visita'", (hoy,))
            visitas = c.fetchone()['c']
            c.execute("SELECT COUNT(*) as c FROM registros_fichajes WHERE fecha = ?", (hoy,))
            fichajes_hoy = c.fetchone()['c']
            return jsonify({
                'success': True, 
                'total_adentro': adentro,
                'entradas': entradas,
                'salidas': salidas,
                'rechazos': rechazos,
                'visitas': visitas,
                'fichajes_hoy': fichajes_hoy,
                'ultima_actualizacion': datetime.now().strftime('%H:%M:%S')
            })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/dashboard/accesos_hoy')
def api_dashboard_accesos_hoy():
    try:
        from db_queries import get_db_connection
        hoy = datetime.now().strftime('%Y-%m-%d')
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT * FROM registros_accesos WHERE fecha = ? ORDER BY hora DESC", (hoy,))
            return jsonify({'success': True, 'records': [dict(r) for r in c.fetchall()]})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/dashboard/fichajes_hoy')
def api_dashboard_fichajes_hoy():
    try:
        from db_queries import get_db_connection
        hoy = datetime.now().strftime('%Y-%m-%d')
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT * FROM registros_fichajes WHERE fecha = ? ORDER BY hora_entrada DESC", (hoy,))
            return jsonify({'success': True, 'records': [dict(r) for r in c.fetchall()]})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/dashboard/directorio_irsa')
def api_dashboard_directorio_irsa():
    try:
        directorio = data_manager.obtener_directorio_irsa()
        return jsonify({'success': True, 'data': directorio})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/dashboard/tramites_irsa')
def api_dashboard_tramites_irsa():
    try:
        tramites = data_manager.obtener_tramites_irsa_raw()
        return jsonify({'success': True, 'data': tramites})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/dashboard/auditoria')
def api_dashboard_auditoria():
    return jsonify({'success': True, 'records': []})

@app.route('/api/config/irsa', methods=['GET', 'POST'])
def api_config_irsa():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    
    from irsa_config import save_irsa_credentials, get_irsa_credentials, has_irsa_credentials
    
    if request.method == 'POST':
        data = request.get_json()
        username = data.get('username', '').strip()
        password = data.get('password', '').strip()
        dias_alerta = int(data.get('dias_alerta_vencimiento', 0) or 0)
        
        if not username:
            return jsonify({'success': False, 'message': 'El usuario no puede estar vacío.'})
        
        # Si no mandaron contraseña, mantenemos la existente
        if not password:
            creds = get_irsa_credentials()
            if creds:
                password = creds.get('password', '')
            else:
                return jsonify({'success': False, 'message': 'Ingresá la contraseña por primera vez.'})
        
        try:
            save_irsa_credentials(username, password, dias_alerta_vencimiento=dias_alerta)
            return jsonify({'success': True, 'message': '✅ Credenciales guardadas correctamente.'})
        except Exception as e:
            logger.error(f"Error guardando credenciales IRSA: {e}")
            return jsonify({'success': False, 'message': f'Error al guardar: {str(e)}'})
    
    else:  # GET
        try:
            has_creds = has_irsa_credentials()
            creds = get_irsa_credentials() if has_creds else None
            return jsonify({
                'success': True,
                'has_credentials': has_creds,
                'username': creds.get('username', '') if creds else '',
                'dias_alerta_vencimiento': creds.get('dias_alerta_vencimiento', 0) if creds else 0,
                'last_sync': None
            })
        except Exception as e:
            return jsonify({'success': True, 'has_credentials': False, 'username': '', 'dias_alerta_vencimiento': 0, 'last_sync': None})

@app.route('/api/dashboard/nominas', methods=['GET'])
def api_dashboard_nominas():
    try:
        data = data_manager.get_nominas_agrupadas()
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/dashboard/nomina', methods=['POST'])
def api_dashboard_nomina():
    if 'logged_in' not in session: return jsonify({'success': False}), 403
    try:
        data = request.json
        # Check if it's the old bulk text parse
        if 'texto' in data:
            res = data_manager.procesar_nomina_texto(data['texto'])
            return jsonify(res)
        
        # New individual ABM
        from db_queries import guardar_empleado_nomina
        s, m = guardar_empleado_nomina(data)
        return jsonify({'success': s, 'message': m})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/dashboard/nomina/<int:id>', methods=['DELETE'])
def api_dashboard_nomina_delete_id(id):
    if 'logged_in' not in session: return jsonify({'success': False}), 403
    try:
        from db_queries import eliminar_empleado_nomina
        s, m = eliminar_empleado_nomina(id)
        return jsonify({'success': s, 'message': m})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/dashboard/lista_negra', methods=['GET', 'POST'])
def api_dashboard_lista_negra():
    return jsonify({'success': True, 'data': []})

@app.route('/api/dashboard/lista_negra/<int:id>', methods=['DELETE'])
def api_dashboard_lista_negra_del(id):
    return jsonify({'success': True})

@app.route('/api/sync/irsa', methods=['POST'])
def api_sync_irsa():
    try:
        from irsa_client import IRSAClient
        client = IRSAClient()
        stats = client.sincronizar_todo()
        mensaje = f'Éxito: {stats["faos_procesados"]} FAOs, {stats["faps_procesados"]} FAPs'
        return jsonify({'success': True, 'message': mensaje})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/admin/backup/download', methods=['GET'])
def api_admin_backup_download():
    return jsonify({'success': False, 'message': 'Not implemented'})

@app.route('/stream/alerts')
def stream_alerts():
    from alert_manager import stream_alerts as stream_alerts_gen
    from flask import Response
    return Response(stream_alerts_gen(), mimetype='text/event-stream')

# PUESTOS OPERATIVOS
@app.route('/api/config/puestos', methods=['GET', 'POST'])
def api_config_puestos():
    if request.method == 'GET':
        from db_queries import obtener_puestos_operativos
        return jsonify({'success': True, 'data': obtener_puestos_operativos()})
    else:
        from db_queries import agregar_puesto_operativo
        data = request.json
        s, m = agregar_puesto_operativo(data.get('nombre', ''))
        return jsonify({'success': s, 'message': m})

@app.route('/api/config/puestos/<int:id>', methods=['DELETE'])
def api_config_puestos_del(id):
    from db_queries import eliminar_puesto_operativo
    s, m = eliminar_puesto_operativo(id)
    return jsonify({'success': s, 'message': m})



# ABM PUESTOS FISICOS
@app.route('/api/config/puestos_fisicos', methods=['GET', 'POST'])
def api_config_puestos_fisicos():
    if request.method == 'GET':
        from db_queries import obtener_puestos_fisicos
        return jsonify({'success': True, 'data': obtener_puestos_fisicos()})
    else:
        from db_queries import agregar_puesto_fisico
        data = request.json
        s, m = agregar_puesto_fisico(data.get('nombre', ''), data.get('sector', ''), data.get('hora_inicio', ''), data.get('hora_fin', ''))
        return jsonify({'success': s, 'message': m})

@app.route('/api/config/puestos_fisicos/<int:id>', methods=['DELETE'])
def api_config_puestos_fisicos_del(id):
    from db_queries import eliminar_puesto_fisico
    s, m = eliminar_puesto_fisico(id)
    return jsonify({'success': s, 'message': m})

# ABM EMPRESAS CONTRATISTAS
@app.route('/api/config/empresas', methods=['GET', 'POST'])
def api_config_empresas():
    if request.method == 'GET':
        from db_queries import obtener_empresas
        return jsonify({'success': True, 'data': obtener_empresas()})
    else:
        from db_queries import agregar_empresa
        data = request.json
        s, m = agregar_empresa(data.get('nombre', ''))
        return jsonify({'success': s, 'message': m})

@app.route('/api/config/empresas/<int:id>', methods=['DELETE'])
def api_config_empresas_del(id):
    from db_queries import eliminar_empresa
    s, m = eliminar_empresa(id)
    return jsonify({'success': s, 'message': m})


# RUTAS DE FICHAJE BIOMÉTRICO (SIMULADAS)
@app.route('/api/biometria/identificar', methods=['POST'])
def api_biometria_identificar():
    # Simula la lectura biométrica devolviendo los datos si el DNI existe y está activo.
    dni = request.json.get('dni')
    from db_queries import get_db_connection
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT a.*, e.nombre as empresa_nombre 
            FROM autorizaciones a
            LEFT JOIN empresas_contratistas e ON a.id_empresa = e.id
            WHERE a.dni = ? AND a.tipo_permiso = 'NOMINA' AND a.activo = 1
        ''', (dni,))
        row = cursor.fetchone()
        
        if row:
            empleado = dict(row)
            empleado['puesto_habitual'] = empleado.get('tarea') # En NOMINA, 'tarea' es la categoria/puesto fisico base
            return jsonify({'success': True, 'empleado': empleado})
        else:
            return jsonify({'success': False, 'message': 'Empleado no encontrado o inactivo.'})

@app.route('/api/fichaje/registrar', methods=['POST'])
def api_fichaje_registrar():
    # Endpoint para registrar el fichaje real de entrada
    dni = request.json.get('dni')
    id_puesto_asignado = request.json.get('id_puesto_asignado')
    puesto_texto = request.json.get('puesto_texto', '')
    
    from db_queries import get_db_connection
    from datetime import datetime
    
    ahora = datetime.now()
    fecha_hoy = ahora.strftime('%Y-%m-%d')
    hora_actual = ahora.strftime('%H:%M:%S')
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        # Obtener datos base
        cursor.execute("SELECT nombre, tarea as categoria FROM autorizaciones WHERE dni=? AND tipo_permiso='NOMINA'", (dni,))
        row = cursor.fetchone()
        if not row:
            return jsonify({'success': False, 'message': 'DNI no encontrado'})
        
        nombre = row['nombre']
        categoria = row['categoria']
        
        cursor.execute('''
            INSERT INTO registros_fichajes 
            (fecha, dni, nombre, hora_entrada, estado, categoria, puesto_historico, id_puesto_asignado, puerta)
            VALUES (?, ?, ?, ?, 'EN CURSO', ?, ?, ?, 'Master')
        ''', (fecha_hoy, dni, nombre, hora_actual, categoria, puesto_texto, id_puesto_asignado))
        
    return jsonify({'success': True})

# === RUTAS DE DEBUG Y DIAGNÓSTICO (ACCESO REMOTO DESDE MULETO) ===
@app.route('/api/debug/ping', methods=['GET'])
def api_debug_ping():
    return jsonify({'success': True, 'status': 'ONLINE', 'time': datetime.now().strftime("%Y-%m-%d %H:%M:%S")})

@app.route('/api/debug/logs', methods=['GET'])
def api_debug_logs():
    lines = request.args.get('lines', 50, type=int)
    log_file = request.args.get('file', 'app.log')
    from config import BASE_DIR
    import os
    
    file_path = os.path.join(BASE_DIR, 'logs', log_file)
    if not os.path.exists(file_path):
        return jsonify({'success': False, 'message': 'Archivo de log no encontrado'})
        
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.readlines()
            
        last_lines = content[-lines:] if len(content) > lines else content
        return jsonify({
            'success': True,
            'file': log_file,
            'lines_returned': len(last_lines),
            'logs': last_lines
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

# SINCRO SATELITE
@app.route('/api/sync/receive_satellite_data', methods=['POST'])
def api_sync_receive_satellite_data():
    data = request.json
    registros = data.get('registros', [])
    if not registros: return jsonify({'success': True})
    
    from db_queries import get_db_connection
    with get_db_connection() as conn:
        cursor = conn.cursor()
        for r in registros:
            cursor.execute("""
                INSERT INTO accesos (timestamp, dni, puerta, estado, tipo_visita)
                VALUES (?, ?, ?, ?, ?)
            """, (r['timestamp'], r['dni'], r['puerta'], r['estado'], r.get('tipo_visita', '')))
        conn.commit()
    return jsonify({'success': True, 'message': 'Registros guardados en Master'})


# --- MÓDULO OTA: ACTUALIZACIONES REMOTAS Y REINICIO ---
import sys
import zipfile
import shutil
from werkzeug.utils import secure_filename

@app.route('/api/admin/restart', methods=['POST'])
def api_admin_restart():
    """Reinicia la aplicación en caliente."""
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    
    # Extraer contraseña enviada en el request para extra seguridad
    data = request.json or {}
    password = data.get('password', '')
    if password != ADMIN_PASSWORD:
        return jsonify({'success': False, 'message': 'Contraseña de administrador incorrecta'}), 401

    logger.warning("Recibida solicitud de reinicio remoto. Reiniciando en 1 segundo...")
    
    def restart_server():
        time.sleep(1.5) # Esperar a que se envíe la respuesta HTTP
        os.execv(sys.executable, ['python'] + sys.argv)
        
    threading.Thread(target=restart_server, daemon=True).start()
    return jsonify({'success': True, 'message': 'Reiniciando aplicación...'})

@app.route('/api/admin/update_code', methods=['POST'])
def api_admin_update_code():
    """Sube un archivo .py o .zip para parchear el código."""
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403

    password = request.form.get('password', '')
    if password != ADMIN_PASSWORD:
        return jsonify({'success': False, 'message': 'Contraseña de administrador incorrecta'}), 401

    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No se envió ningún archivo'}), 400
        
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'Nombre de archivo vacío'}), 400

    filename = secure_filename(file.filename)
    if not (filename.endswith('.py') or filename.endswith('.zip')):
        return jsonify({'success': False, 'message': 'Formato no soportado. Use .py o .zip'}), 400

    try:
        temp_dir = os.path.join(BASE_DIR, 'temp_updates')
        os.makedirs(temp_dir, exist_ok=True)
        filepath = os.path.join(temp_dir, filename)
        file.save(filepath)

        if filename.endswith('.py'):
            # Reemplazar archivo py directo en el root
            dest = os.path.join(BASE_DIR, filename)
            shutil.copy2(filepath, dest)
            logger.info(f"Parcheado archivo: {filename}")
            
        elif filename.endswith('.zip'):
            # Descomprimir en el root (sobrescribe archivos existentes)
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                zip_ref.extractall(BASE_DIR)
            logger.info(f"Parche ZIP aplicado: {filename}")

        # Limpieza
        shutil.rmtree(temp_dir, ignore_errors=True)
        
        return jsonify({'success': True, 'message': 'Código actualizado exitosamente. El sistema se reiniciará para aplicar los cambios.'})
    except Exception as e:
        logger.error(f"Error aplicando actualización: {e}", exc_info=True)
        return jsonify({'success': False, 'message': str(e)}), 500

if __name__ == '__main__':
    print("Iniciando aplicación en modo de depuración...")
    app.run(host='0.0.0.0', port=5050, debug=True)
