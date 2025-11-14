import os
from datetime import datetime
import pandas as pd
from flask import (Flask, jsonify, request, render_template, redirect, url_for, session, send_from_directory, after_this_request)
from werkzeug.utils import secure_filename
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- INICIO DE CORRECCIÓN ---

import data_manager

from config import (

    BASE_DIR, SECRET_KEY, ADMIN_USERNAME, ADMIN_PASSWORD,

    EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER, SMTP_SERVER, SMTP_PORT,

    REGISTROS_DIARIOS_DIR, REGISTROS_FICHAJES_DIR,

    EXCEL_FAP, EXCEL_FAO, EXCEL_EXCEPCIONES, EXCEL_NOMINAS

)

import re

from data_manager import (
    cargar_autorizaciones,
    extraer_dni_de_cuil,
    get_nominas_agrupadas,  # <--- ASEGÚRATE DE AGREGAR ESTA LÍNEA
    delete_nomina_by_criteria,
    get_nomina_detalle_by_criteria,
    get_df_nominas_persistentes,
    recargar_cache_nominas_persistentes,
    procesar_nomina_texto,
    generar_reporte_consolidado
)

from access_manager import (

    registrar_fichaje,

    verificar_dni,

    personas_adentro

)



# --- CONFIGURACIÓN ---

app = Flask(__name__, template_folder=os.path.join(BASE_DIR, 'templates'), static_folder=os.path.join(BASE_DIR, 'static'))

app.secret_key = SECRET_KEY



os.makedirs(REGISTROS_DIARIOS_DIR, exist_ok=True)

os.makedirs(REGISTROS_FICHAJES_DIR, exist_ok=True)



# --- RUTAS PÚBLICAS ---

@app.route('/')

def home():

    return render_template('index.html')



@app.route('/login')

def login_page():

    return render_template('login.html')



# --- RUTAS DE API (para el frontend) ---

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



# --- RUTAS DE ADMINISTRACIÓN ---

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

    

    if fap_file:

        fap_file.save(EXCEL_FAP)

    if fao_file:

        fao_file.save(EXCEL_FAO)

        

    cargar_autorizaciones() # Forzar recarga de datos

    return jsonify({'success': True, 'message': 'Archivos subidos y procesados.'})











@app.route('/agregar_excepcion', methods=['POST'])
def agregar_excepcion():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    
    data = request.get_json()
    
    # --- Columnas ESTÁNDAR que SIEMPRE usaremos ---
    columnas_excepcion = ['Numero', 'Nombre Completo', 'Local', 'Quien Autoriza', 'Fecha de Alta', 'Vigencia']
    
    df_actual = pd.DataFrame(columns=columnas_excepcion)
    
    try:
        # 1. Intentamos leer el archivo existente
        if os.path.exists(EXCEL_EXCEPCIONES):
            df_leido = pd.read_excel(EXCEL_EXCEPCIONES)
            # 2. Nos aseguramos de que solo tenga las columnas que nos importan
            # Si faltan columnas nuevas, se añadirán con valores vacíos (NaN)
            for col in columnas_excepcion:
                if col not in df_leido.columns:
                    df_leido[col] = pd.NA
            df_actual = df_leido[columnas_excepcion].copy()
            
    except Exception as e:
        # Si falla (porque no existe o está corrupto), simplemente usamos el df vacío que creamos
        print(f"INFO: No se pudo leer {EXCEL_EXCEPCIONES} (puede ser la primera vez). Se creará uno nuevo. Error: {e}")
        df_actual = pd.DataFrame(columns=columnas_excepcion)

    # 3. Preparamos los datos nuevos
    if 'Numero' in df_actual.columns:
        df_actual['Numero'] = df_actual['Numero'].astype(str)
    
    dni_nuevo = str(data['dni'])
    nombre_nuevo = f"{data['nombre']} {data['apellido']}"
    local_nuevo = data['local']
    autoriza_nuevo = data['autoriza']
    vigencia_nueva = data.get('vigencia') # Capturamos la vigencia
    fecha_alta_nueva = datetime.now().strftime('%Y-%m-%d')

    # Convertir la vigencia a datetime para asegurar el formato correcto en Excel
    if vigencia_nueva:
        vigencia_dt = pd.to_datetime(vigencia_nueva, errors='coerce')
    else:
        vigencia_dt = pd.NaT


    idx = df_actual[df_actual['Numero'] == dni_nuevo].index

    if not idx.empty:
        # SI EXISTE: Actualizamos la fila
        print(f"DEBUG: DNI {dni_nuevo} encontrado. Actualizando excepción.")
        df_actual.loc[idx[0], 'Nombre Completo'] = nombre_nuevo
        df_actual.loc[idx[0], 'Local'] = local_nuevo
        df_actual.loc[idx[0], 'Quien Autoriza'] = autoriza_nuevo
        df_actual.loc[idx[0], 'Fecha de Alta'] = fecha_alta_nueva
        df_actual.loc[idx[0], 'Vigencia'] = vigencia_dt
        mensaje_exito = 'Excepción actualizada.'
    else:
        # NO EXISTE: Agregamos una nueva fila
        print(f"DEBUG: DNI {dni_nuevo} no encontrado. Agregando nueva excepción.")
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
        # 4. Guardamos SIEMPRE un DataFrame limpio con las columnas correctas
        df_actual.to_excel(EXCEL_EXCEPCIONES, index=False, columns=columnas_excepcion)
        
        # 5. Forzamos la recarga
        cargar_autorizaciones()
        return jsonify({'success': True, 'message': mensaje_exito})
        
    except Exception as e:
        print(f"Error Crítico al guardar excepción en Excel: {e}")
        return jsonify({'success': False, 'message': f'Error al guardar: {e}'})

@app.route('/parse_nomina', methods=['POST'])
def parse_nomina():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403

    data = request.get_json()
    texto_pegado = data.get('texto_pegado', '')

    if not texto_pegado:
        return jsonify({'success': False, 'message': 'El texto de la nómina no puede estar vacío.'})

    # Usamos la función centralizada de data_manager
    personas_procesadas = procesar_nomina_texto(texto_pegado)

    if not personas_procesadas:
        return jsonify({'success': False, 'message': 'No se pudo interpretar ninguna persona. Revise el formato del texto.'})

    # Adaptamos el formato de salida para el frontend (claves en minúscula)
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
                print(f"Error al procesar fechas durante la actualización: {e}")
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
        print(f"Error Crítico al guardar la nómina: {e}")
        return jsonify({'success': False, 'message': f'Error del servidor: {e}'}), 500

@app.route('/get_nominas_guardadas', methods=['GET'])

def get_nominas_guardadas_route():

    if 'logged_in' not in session:

        return jsonify({'success': False, 'message': 'No autorizado'}), 403

    try:

        nominas = get_nominas_agrupadas()

        return jsonify({'success': True, 'nominas': nominas})

    except Exception as e:

        print(f"Error al obtener nóminas guardadas: {e}")

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





        print(f"Error Crítico al eliminar la nómina: {e}")





        return jsonify({'success': False, 'message': f'Error del servidor: {e}'}), 500



@app.route('/get_nomina_detalle', methods=['POST'])



def get_nomina_detalle_route():



    if 'logged_in' not in session:



        return jsonify({'success': False, 'message': 'No autorizado'}), 403







    data = request.get_json()



    empresa = data.get('empresa')



    vigencia = data.get('vigencia') # Cambiado de vigencia_desde a vigencia



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



        print(f"Error al obtener el detalle de la nómina: {e}")



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

        return jsonify({'success': True, 'message': 'Reportes enviados por email.'})

    except Exception as e:

        return jsonify({'success': False, 'message': f'Error al enviar email: {e}'})



@app.route('/descargar_reporte_diario')
def descargar_reporte_diario():
    if 'logged_in' not in session:
        return redirect(url_for('login_page'))

    # Llama a la nueva función para generar el reporte consolidado
    temp_path, filename = generar_reporte_consolidado()

    if not temp_path:
        return "No hay reporte de accesos para hoy.", 404

    temp_dir = os.path.dirname(temp_path)

    @after_this_request
    def cleanup(response):
        try:
            os.remove(temp_path)
            print(f"INFO: Archivo temporal '{filename}' eliminado.")
        except Exception as e:
            print(f"Error al eliminar archivo temporal: {e}")
        return response

    return send_from_directory(directory=temp_dir, path=filename, as_attachment=True)



@app.route('/descargar_reporte_fichajes')

def descargar_reporte_fichajes():

    if 'logged_in' not in session:

        return redirect(url_for('login_page'))

    fecha_hoy = datetime.now().strftime('%Y-%m-%d')

    archivo_fichajes = os.path.join(REGISTROS_FICHAJES_DIR, f'registros_fichaje_{fecha_hoy}.xlsx')

    if not os.path.exists(archivo_fichajes):

        return "No hay reporte de fichajes para hoy.", 404

    return send_from_directory(directory=REGISTROS_FICHAJES_DIR, path=f'registros_fichaje_{fecha_hoy}.xlsx', as_attachment=True)



# --- EJECUCIÓN DE LA APP ---

if __name__ == '__main__':

    with app.app_context():

        cargar_autorizaciones()

    app.run(debug=True, host='0.0.0.0', port=5000)
