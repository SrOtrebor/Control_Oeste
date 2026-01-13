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


app = Flask(__name__)
app.secret_key = SECRET_KEY

# Logger para este módulo
logger = get_logger(__name__)

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
    app.run(port=5000, use_reloader=False) 

if __name__ == '__main__':
    with app.app_context():
        cargar_autorizaciones()

    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()
    time.sleep(1) 

    window = webview.create_window('Control de Acceso', 'http://127.0.0.1:5000', js_api=api)
    api.set_window(window)
    webview.start()
