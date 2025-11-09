import os
from datetime import datetime
import pandas as pd
from flask import (Flask, jsonify, request, render_template, redirect, url_for, session, send_from_directory)
from werkzeug.utils import secure_filename
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- INICIO DE CORRECCIÓN ---
import data_manager # Importamos el módulo completo
# --- FIN DE CORRECCIÓN ---

from config import (
    BASE_DIR, SECRET_KEY, ADMIN_USERNAME, ADMIN_PASSWORD,
    EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER, SMTP_SERVER, SMTP_PORT,
    REGISTROS_DIARIOS_DIR, REGISTROS_FICHAJES_DIR,
    EXCEL_FAP, EXCEL_FAO, EXCEL_EXCEPCIONES, EXCEL_NOMINAS
)
from data_manager import (
    cargar_autorizaciones,
    # Ya no importamos los dataframes desde aquí
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

@app.route('/upload_logo', methods=['POST'])
def upload_logo():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    logo_file = request.files.get('logoFile')
    if logo_file:
        filename = secure_filename(logo_file.filename)
        logo_path = os.path.join(app.static_folder, 'logo.png')
        logo_file.save(logo_path)
        return jsonify({'success': True, 'message': 'Logo actualizado correctamente.'})
    return jsonify({'success': False, 'message': 'No se proporcionó ningún archivo.'})

@app.route('/delete_logo', methods=['POST'])
def delete_logo_route():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    try:
        logo_path = os.path.join(app.static_folder, 'logo.png')
        if os.path.exists(logo_path):
            os.remove(logo_path)
            return jsonify({'success': True, 'message': 'Logo eliminado correctamente.'})
        else:
            return jsonify({'success': False, 'message': 'No se encontró ningún logo para eliminar.'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error al eliminar el logo: {e}'}), 500

@app.route('/agregar_excepcion', methods=['POST'])
def agregar_excepcion():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    
    data = request.get_json()
    
    try:
        # Usamos el DataFrame que está en la memoria (en data_manager)
        df_actual = data_manager.df_excepciones.copy()
    except AttributeError: 
        df_actual = pd.DataFrame(columns=['Numero', 'Nombre Completo', 'Local', 'Quien Autoriza', 'Fecha de Alta'])

    if 'Numero' in df_actual.columns:
        df_actual['Numero'] = df_actual['Numero'].astype(str)
    else:
        df_actual['Numero'] = ''

    dni_nuevo = str(data['dni'])
    nombre_nuevo = f"{data['nombre']} {data['apellido']}"
    local_nuevo = data['local']
    autoriza_nuevo = data['autoriza']
    fecha_nueva = datetime.now().strftime('%Y-%m-%d') # Usamos formato estándar

    idx = df_actual[df_actual['Numero'] == dni_nuevo].index

    if not idx.empty:
        print(f"DEBUG: DNI {dni_nuevo} encontrado. Actualizando excepción.")
        df_actual.loc[idx[0], 'Nombre Completo'] = nombre_nuevo
        df_actual.loc[idx[0], 'Local'] = local_nuevo
        df_actual.loc[idx[0], 'Quien Autoriza'] = autoriza_nuevo
        df_actual.loc[idx[0], 'Fecha de Alta'] = fecha_nueva
        mensaje_exito = 'Excepción actualizada.'
    else:
        print(f"DEBUG: DNI {dni_nuevo} no encontrado. Agregando nueva excepción.")
        nuevo_registro = pd.DataFrame([{
            'Numero': dni_nuevo,
            'Nombre Completo': nombre_nuevo,
            'Local': local_nuevo,
            'Quien Autoriza': autoriza_nuevo,
            'Fecha de Alta': fecha_nueva
        }])
        df_actual = pd.concat([df_actual, nuevo_registro], ignore_index=True)
        mensaje_exito = 'Excepción agregada.'
    
    try:
        df_actual.to_excel(EXCEL_EXCEPCIONES, index=False)
        cargar_autorizaciones() # Forzamos la recarga
        return jsonify({'success': True, 'message': mensaje_exito})
    except Exception as e:
        print(f"Error Crítico al guardar excepción en Excel: {e}")
        return jsonify({'success': False, 'message': f'Error al guardar: {e}'})

@app.route('/parse_nomina', methods=['POST'])
def parse_nomina():
    data = request.get_json()
    texto_pegado = data.get('texto_pegado', '')
    personas = []
    for linea in texto_pegado.splitlines():
        partes = linea.split('\t') 
        if len(partes) >= 3:
            personas.append({'dni': partes[0], 'apellido': partes[1], 'nombre': partes[2]})
    return jsonify({'success': True, 'nomina': personas})

@app.route('/save_nomina', methods=['POST'])
def save_nomina():
    if 'logged_in' not in session:
        return jsonify({'success': False, 'message': 'No autorizado'}), 403
    data = request.get_json()
    nomina = data.get('nomina', [])
    try:
        df = pd.read_excel(EXCEL_NOMINAS)
    except FileNotFoundError:
        df = pd.DataFrame()
    df_nueva = pd.DataFrame(nomina)
    df_nueva['Empresa'] = data.get('empresa')
    df_nueva['Vigencia Desde'] = data.get('vigencia_desde')
    df_nueva['Vigencia Hasta'] = data.get('vigencia_hasta')
    df_final = pd.concat([df, df_nueva], ignore_index=True)
    df_final.to_excel(EXCEL_NOMINAS, index=False)
    cargar_autorizaciones()
    return jsonify({'success': True, 'message': 'Nómina guardada correctamente.'})

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
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    archivo_accesos = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_hoy}.xlsx')
    if not os.path.exists(archivo_accesos):
        return "No hay reporte de accesos para hoy.", 404
    return send_from_directory(directory=REGISTROS_DIARIOS_DIR, path=f'registros_ingreso_{fecha_hoy}.xlsx', as_attachment=True)

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