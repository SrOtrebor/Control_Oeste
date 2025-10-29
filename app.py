import os
import smtplib
import sys
import webbrowser
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd
from flask import (Flask, jsonify, request, render_template, redirect, url_for, session)

# --- CONFIGURACIÓN ---
if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    static_folder = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
    BASE_DIR = os.path.dirname(sys.executable)
else:
    app = Flask(__name__)
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app.secret_key = 'tu_clave_secreta_aqui_CAMBIALA'
ADMIN_USERNAME = 'Seguridad'
ADMIN_PASSWORD = 'ControldeAcceso1.0'

EMAIL_SENDER = 'acceso.alcorta@gmail.com'
EMAIL_PASSWORD = 'vmcg dqel hhdx zzqj'
EMAIL_RECEIVER = 'rlaforcada@irsacorp.com.ar'
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587

REGISTROS_DIARIOS_DIR = os.path.join(BASE_DIR, 'registros_diarios')
EXCEL_FAP = os.path.join(BASE_DIR, 'ListadoFAPs.xlsx')
EXCEL_FAO = os.path.join(BASE_DIR, 'ListadoFAOs.xlsx')
EXCEL_EXCEPCIONES = os.path.join(BASE_DIR, 'excepciones.xlsx')

COL_DNI, COL_NOMBRE_APELLIDO, COL_NUM_PERMISO, COL_VENCE, COL_LOCAL, COL_TAREA, COL_TIPO_PERMISO = 'DNI', 'Nombre y Apellido', 'Num_Permiso', 'Vence', 'Local', 'Tarea', 'Tipo de Permiso'
COL_DNI_FAP_ORIGINAL, COL_NOMBRE_FAP_ORIGINAL, COL_APELLIDO_FAP_ORIGINAL, COL_NUM_PERMISO_FAP_ORIGINAL, COL_VENCE_FAP_ORIGINAL, COL_LOCAL_FAP_ORIGINAL = 'Numero', 'Nombre', 'Apellido', 'FAP', 'Fecha Fin', 'Marca'
COL_DNI_FAO_ORIGINAL, COL_NOMBRE_FAO_ORIGINAL, COL_APELLIDO_FAO_ORIGINAL, COL_NUM_PERMISO_FAO_ORIGINAL, COL_VENCE_FAO_ORIGINAL, COL_LOCAL_FAO_ORIGINAL, COL_TAREA_FAO_ORIGINAL = 'Numero', 'Nombre', 'Apellido', 'FAO', 'Fecha Fin', 'Marca', 'Tarea'

df_fap, df_fao, df_excepciones = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
ult_mod_fap, ult_mod_fao, ult_mod_excepciones = 0, 0, 0

os.makedirs(REGISTROS_DIARIOS_DIR, exist_ok=True)

# --- FUNCIONES AUXILIARES ---

def extraer_dni_de_cuil(valor):
    valor_str = str(valor).strip()
    if len(valor_str) == 11 and valor_str.isdigit():
        return valor_str[2:10]
    return valor_str

def cargar_y_procesar_excel(filepath, ult_mod_global, tipo_permiso, cols_rename_map, df_global):
    try:
        mod_time = os.path.getmtime(filepath)
        if mod_time <= ult_mod_global and not df_global.empty:
            return df_global, ult_mod_global
        header_row = 1 if tipo_permiso in ['FAP', 'FAO'] else 0
        df_temp = pd.read_excel(filepath, header=header_row)
        df_temp = df_temp.rename(columns=cols_rename_map)
        if COL_DNI in df_temp.columns:
            df_temp[COL_DNI] = pd.to_numeric(df_temp[COL_DNI], errors='coerce')
            df_temp.dropna(subset=[COL_DNI], inplace=True)
            df_temp[COL_DNI] = df_temp[COL_DNI].astype('Int64').astype(str).apply(extraer_dni_de_cuil)
        if COL_VENCE in df_temp.columns:
            df_temp[COL_VENCE] = pd.to_datetime(df_temp[COL_VENCE], dayfirst=True, errors='coerce')
        if 'Nombre Completo' in df_temp.columns and COL_NOMBRE_APELLIDO not in df_temp.columns:
            df_temp[COL_NOMBRE_APELLIDO] = df_temp['Nombre Completo']
        elif 'Nombre' in df_temp.columns and 'Apellido' in df_temp.columns:
             df_temp[COL_NOMBRE_APELLIDO] = df_temp['Nombre'].fillna('') + ' ' + df_temp['Apellido'].fillna('')
             df_temp[COL_NOMBRE_APELLIDO] = df_temp[COL_NOMBRE_APELLIDO].str.strip()
        df_temp[COL_TIPO_PERMISO] = tipo_permiso
        for col in [COL_LOCAL, COL_TAREA, COL_NUM_PERMISO, COL_NOMBRE_APELLIDO, 'Quien_Autoriza']:
            if col not in df_temp.columns:
                df_temp[col] = 'N/A'
        return df_temp, mod_time
    except FileNotFoundError:
        return pd.DataFrame(), 0
    except Exception as e:
        print(f"Error al cargar '{filepath}': {e}")
        return pd.DataFrame(), ult_mod_global

def cargar_autorizaciones():
    global df_fap, ult_mod_fap, df_fao, ult_mod_fao, df_excepciones, ult_mod_excepciones
    mapa_cols_fap = {COL_DNI_FAP_ORIGINAL: COL_DNI, COL_NOMBRE_FAP_ORIGINAL: 'Nombre', COL_APELLIDO_FAP_ORIGINAL: 'Apellido', COL_NUM_PERMISO_FAP_ORIGINAL: COL_NUM_PERMISO, COL_VENCE_FAP_ORIGINAL: COL_VENCE, COL_LOCAL_FAP_ORIGINAL: COL_LOCAL}
    df_fap, ult_mod_fap = cargar_y_procesar_excel(EXCEL_FAP, ult_mod_fap, 'FAP', mapa_cols_fap, df_fap)
    mapa_cols_fao = {COL_DNI_FAO_ORIGINAL: COL_DNI, COL_NOMBRE_FAO_ORIGINAL: 'Nombre', COL_APELLIDO_FAO_ORIGINAL: 'Apellido', COL_NUM_PERMISO_FAO_ORIGINAL: COL_NUM_PERMISO, COL_VENCE_FAO_ORIGINAL: COL_VENCE, COL_LOCAL_FAO_ORIGINAL: COL_LOCAL, COL_TAREA_FAO_ORIGINAL: COL_TAREA}
    df_fao, ult_mod_fao = cargar_y_procesar_excel(EXCEL_FAO, ult_mod_fao, 'FAO', mapa_cols_fao, df_fao)
    mapa_cols_excepciones = {'Numero': COL_DNI, 'Nombre Completo': 'Nombre Completo', 'Fecha de Alta': COL_VENCE, 'Local': COL_LOCAL, 'Quien Autoriza': 'Quien_Autoriza'}
    df_excepciones, ult_mod_excepciones = cargar_y_procesar_excel(EXCEL_EXCEPCIONES, ult_mod_excepciones, 'Excepcion', mapa_cols_excepciones, df_excepciones)

def guardar_registro(dni, nombre, hora_ingreso, tipo_permiso, num_permiso, local, tarea, resultado):
    fecha_actual_str = datetime.now().strftime('%Y-%m-%d')
    nombre_archivo = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_actual_str}.xlsx')
    columnas_registro = ['DNI', 'Nombre y Apellido', 'Hora_Ingreso', 'Tipo_Permiso', 'Num_Permiso', 'Local', 'Tarea', 'Resultado']
    nuevo_registro_df = pd.DataFrame([{'DNI': dni, 'Nombre y Apellido': nombre, 'Hora_Ingreso': hora_ingreso, 'Tipo_Permiso': tipo_permiso, 'Num_Permiso': num_permiso, 'Local': local, 'Tarea': tarea, 'Resultado': resultado}])
    try:
        if not os.path.exists(nombre_archivo):
            df_registros = pd.DataFrame(columns=columnas_registro)
        else:
            df_registros = pd.read_excel(nombre_archivo)
        df_final = pd.concat([df_registros, nuevo_registro_df], ignore_index=True)
        df_final.to_excel(nombre_archivo, index=False)
    except Exception as e:
        print(f"Error Crítico al guardar registro en Excel: {e}")

def crear_reporte_formateado(ruta_original):
    try:
        df = pd.read_excel(ruta_original)
        if df.empty: return None
        df = df.iloc[::-1].reset_index(drop=True)
        columna_resultado = df['Resultado']
        df_final_reporte = df.drop(columns=['Resultado'])
        df_final_reporte = df_final_reporte.rename(columns={'Num_Permiso': 'Permiso/Autoriza'})
        if 'Local' in df_final_reporte.columns:
            df_final_reporte['Local'] = df_final_reporte['Local'].astype(str).apply(lambda x: x.split('\n')[0])
        ruta_formateada = ruta_original.replace(".xlsx", "_formateado.xlsx")
        writer = pd.ExcelWriter(ruta_formateada, engine='xlsxwriter')
        df_final_reporte.to_excel(writer, sheet_name='Reporte de Ingresos', index=False, startrow=3)
        workbook, worksheet = writer.book, writer.sheets['Reporte de Ingresos']
        header_format, title_format, summary_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#2B2D42', 'font_color': 'white', 'border': 1}), workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#2B2D42'}), workbook.add_format({'bold': True, 'font_size': 11})
        green_format, red_format, excepcion_format, fao_format = workbook.add_format({'bg_color': '#eaf7ed'}), workbook.add_format({'bg_color': '#fdecea'}), workbook.add_format({'bg_color': '#e0f7fa'}), workbook.add_format({'bg_color': '#fff9c4'})
        total, permitidos = len(df), (columna_resultado == 'VERDE').sum()
        denegados = total - permitidos
        worksheet.merge_range('A1:D1', f"Reporte de Ingresos - {datetime.now().strftime('%d/%m/%Y')}", title_format)
        worksheet.write('A2', 'Resumen:', summary_format)
        worksheet.write('B2', f'Total: {total} | Permitidos: {permitidos} | Denegados: {denegados}')
        for col_num, value in enumerate(df_final_reporte.columns.values):
            worksheet.write(3, col_num, value, header_format)
        num_filas = len(df_final_reporte)
        for i, (idx, row) in enumerate(df_final_reporte.iterrows()):
            fila_excel, formato_a_aplicar, resultado_actual = i + 4, None, columna_resultado.iloc[i]
            if resultado_actual == 'VERDE': formato_a_aplicar = green_format
            elif resultado_actual == 'ROJO': formato_a_aplicar = red_format
            if row['Tipo_Permiso'] == 'Excepcion': formato_a_aplicar = excepcion_format
            elif row['Tipo_Permiso'] == 'FAO': formato_a_aplicar = fao_format
            if formato_a_aplicar: worksheet.set_row(fila_excel, None, formato_a_aplicar)
        for idx, col in enumerate(df_final_reporte):
            max_len = max((df_final_reporte[col].astype(str).map(len).max(), len(str(col)))) + 3
            worksheet.set_column(idx, idx, max_len)
        worksheet.autofilter(3, 0, num_filas + 3, len(df_final_reporte.columns) - 1)
        worksheet.freeze_panes(4, 0)
        writer.close()
        return ruta_formateada
    except Exception as e:
        print(f"Error al crear el reporte formateado: {e}")
        return None

def enviar_email(archivo_adjunto=None):
    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = EMAIL_SENDER, EMAIL_RECEIVER, f"Reporte Diario de Admisión - {datetime.now().strftime('%Y-%m-%d')}"
    msg.attach(MIMEText("Adjunto el reporte diario de admisiones.", 'plain'))
    if archivo_adjunto and os.path.exists(archivo_adjunto):
        try:
            with open(archivo_adjunto, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(archivo_adjunto)}')
            msg.attach(part)
        except Exception as e: print(f"Error al adjuntar archivo: {e}")
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())
        return True
    except Exception as e:
        print(f"Error al enviar el correo: {e}")
        return False

# --- RUTAS DE FLASK ---

@app.route('/')
def home(): return render_template('index.html')

@app.route('/login')
def login_page(): return render_template('login.html')

@app.route('/admin')
def admin_page():
    if 'logged_in' in session and session['logged_in']: return render_template('admin.html')
    return redirect(url_for('login_page'))

@app.route('/perform_login', methods=['POST'])
def perform_login():
    data = request.get_json()
    if data.get('username') == ADMIN_USERNAME and data.get('password') == ADMIN_PASSWORD:
        session['logged_in'] = True
        return jsonify({'success': True})
    return jsonify({'success': False, 'message': 'Usuario o contraseña incorrectos.'})

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login_page'))

@app.route('/verificar_dni', methods=['POST'])
def verificar_dni():
    data = request.get_json()
    dni_ingresado = str(data.get('dni', '')).strip()
    cargar_autorizaciones()
    hoy = datetime.now()
    hora_actual_str = hoy.strftime('%H:%M:%S')
    respuesta = {'acceso': 'DENEGADO', 'nombre': 'No Encontrado', 'mensaje': f'ACCESO DENEGADO: DNI {dni_ingresado} no encontrado o permiso vencido.', 'tipo_permiso': 'N/A', 'num_permiso': 'N/A', 'local': 'N/A', 'tarea': 'N/A', 'vence': 'N/A'}
    if not dni_ingresado.isdigit() or len(dni_ingresado) < 7:
        respuesta['mensaje'] = "Formato de DNI inválido."
        guardar_registro(dni_ingresado, 'N/A', hora_actual_str, 'N/A', 'N/A', 'N/A', 'N/A', 'ROJO')
        return jsonify(respuesta)
    for df, tipo in [(df_excepciones, 'Excepcion'), (df_fap, 'FAP'), (df_fao, 'FAO')]:
        if df.empty or COL_DNI not in df.columns: continue
        match = df[df[COL_DNI] == dni_ingresado]
        if not match.empty:
            persona = match.iloc[0]
            fecha_vencimiento = persona.get(COL_VENCE)
            
            respuesta['nombre'] = persona.get(COL_NOMBRE_APELLIDO, 'N/A')
            respuesta['tipo_permiso'] = tipo
            
            if tipo == 'Excepcion':
                respuesta['num_permiso'] = str(persona.get('Quien_Autoriza', 'N/A'))
            else:
                respuesta['num_permiso'] = str(persona.get(COL_NUM_PERMISO, 'N/A'))

            # Limpieza del dato "Local"
            local_raw = persona.get(COL_LOCAL, 'N/A')
            respuesta['local'] = str(local_raw).split('\n')[0].replace('Local', '').strip()

            respuesta['tarea'] = str(persona.get(COL_TAREA, 'N/A'))
            
            if pd.isna(fecha_vencimiento) or fecha_vencimiento.date() >= hoy.date():
                respuesta['acceso'] = 'PERMITIDO'
                respuesta['mensaje'] = f"ACCESO PERMITIDO ({tipo}): {respuesta['nombre']}"
                respuesta['vence'] = 'Indefinido' if pd.isna(fecha_vencimiento) else fecha_vencimiento.strftime('%d/%m/%Y')
            else:
                respuesta['acceso'] = 'DENEGADO'
                respuesta['vence'] = fecha_vencimiento.strftime('%d/%m/%Y')
                respuesta['mensaje'] = f"ACCESO DENEGADO: Permiso {tipo} vencido el {respuesta['vence']}."
            
            resultado_log = 'VERDE' if respuesta['acceso'] == 'PERMITIDO' else 'ROJO'
            guardar_registro(dni_ingresado, respuesta['nombre'], hora_actual_str, respuesta['tipo_permiso'], respuesta['num_permiso'], respuesta['local'], respuesta['tarea'], resultado_log)
            return jsonify(respuesta)
            
    guardar_registro(dni_ingresado, 'No Encontrado', hora_actual_str, 'N/A', 'N/A', 'N/A', 'N/A', 'ROJO')
    return jsonify(respuesta)

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'logged_in' not in session or not session['logged_in']:
        return jsonify({'success': False, 'message': 'No autorizado.'}), 403
    messages, success = [], True
    def procesar_archivo(file_key, target_path):
        nonlocal success
        archivo = request.files.get(file_key)
        if archivo and archivo.filename != '':
            try:
                temp_path = target_path + ".tmp"
                archivo.save(temp_path)
                pd.read_excel(temp_path, header=1 if file_key.startswith('f') else 0)
                os.replace(temp_path, target_path)
                messages.append(f'Archivo {file_key.split("_")[0].upper()} actualizado.')
            except Exception as e:
                success = False
                messages.append(f'Error al procesar {archivo.filename}: {e}')
                if os.path.exists(temp_path): os.remove(temp_path)
    procesar_archivo('fap_file', EXCEL_FAP)
    procesar_archivo('fao_file', EXCEL_FAO)
    cargar_autorizaciones()
    return jsonify({'success': success, 'message': ' '.join(messages)})

@app.route('/add_excepcion', methods=['POST'])
def add_excepcion():
    if 'logged_in' not in session or not session['logged_in']:
        return jsonify({'success': False, 'message': 'No autorizado.'}), 403
    data = request.get_json()
    nombre, apellido, dni, local, quien_autoriza = data.get('nombre', ''), data.get('apellido', ''), data.get('dni', ''), data.get('local', ''), data.get('quien_autoriza', '')
    if not all([nombre, apellido, dni, local, quien_autoriza]):
        return jsonify({'success': False, 'message': 'Todos los campos son obligatorios.'})
    try:
        columnas = ['Numero', 'Nombre Completo', 'Local', 'Quien Autoriza', 'Fecha de Alta']
        df_excepciones_actual = pd.read_excel(EXCEL_EXCEPCIONES) if os.path.exists(EXCEL_EXCEPCIONES) else pd.DataFrame(columns=columnas)
        if 'Numero' in df_excepciones_actual.columns:
            df_excepciones_actual['Numero'] = df_excepciones_actual['Numero'].astype(str)
        mask = df_excepciones_actual['Numero'] == dni
        if mask.any():
            idx = df_excepciones_actual[mask].index
            df_excepciones_actual.loc[idx, ['Nombre Completo', 'Local', 'Quien Autoriza', 'Fecha de Alta']] = [f"{nombre} {apellido}", local, quien_autoriza, datetime.now().strftime('%d/%m/%Y %H:%M:%S')]
            msg = 'Excepción actualizada.'
        else:
            nuevo_registro = pd.DataFrame([{'Numero': dni, 'Nombre Completo': f"{nombre} {apellido}", 'Local': local, 'Quien Autoriza': quien_autoriza, 'Fecha de Alta': datetime.now().strftime('%d/%m/%Y %H:%M:%S')}])
            df_excepciones_actual = pd.concat([df_excepciones_actual, nuevo_registro], ignore_index=True)
            msg = 'Excepción agregada.'
        df_excepciones_actual.to_excel(EXCEL_EXCEPCIONES, index=False)
        global ult_mod_excepciones
        ult_mod_excepciones = 0 
        cargar_autorizaciones()
        return jsonify({'success': True, 'message': msg})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error al guardar la excepción: {e}'})

@app.route('/enviar_reporte_diario', methods=['POST'])
def enviar_reporte():
    if 'logged_in' not in session or not session['logged_in']:
        return jsonify({'success': False, 'message': 'No autorizado.'}), 403
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    archivo_original = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_hoy}.xlsx')
    if os.path.exists(archivo_original):
        archivo_formateado = crear_reporte_formateado(archivo_original)
        if archivo_formateado:
            if enviar_email(archivo_formateado):
                os.remove(archivo_formateado)
                return jsonify({'success': True, 'message': 'Reporte enviado exitosamente.'})
            else:
                return jsonify({'success': False, 'message': 'Error al enviar el reporte.'})
        else:
            return jsonify({'success': False, 'message': 'Error al generar el reporte formateado.'})
    else:
        return jsonify({'success': False, 'message': 'No hay registros de entradas para hoy.'})

@app.route('/get_daily_records')
def get_daily_records():
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    nombre_archivo = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_hoy}.xlsx')
    if not os.path.exists(nombre_archivo):
        return jsonify({'success': True, 'records': [], 'message': 'Aún no hay registros para hoy.'})
    try:
        df_registros = pd.read_excel(nombre_archivo).fillna('')
        return jsonify({'success': True, 'records': df_registros.to_dict('records')})
    except Exception as e:
        print(f"Error al leer registros diarios: {e}")
        return jsonify({'success': False, 'message': 'Error al procesar registros.', 'records': []}), 500

@app.route('/get_daily_stats')
def get_daily_stats():
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    nombre_archivo = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_hoy}.xlsx')
    if not os.path.exists(nombre_archivo):
        return jsonify({'total': 0, 'permitidos': 0, 'rechazados': 0})
    try:
        df = pd.read_excel(nombre_archivo)
        if df.empty: return jsonify({'total': 0, 'permitidos': 0, 'rechazados': 0})
        total = len(df)
        permitidos = (df['Resultado'] == 'VERDE').sum()
        return jsonify({'total': int(total), 'permitidos': int(permitidos), 'rechazados': int(total - permitidos)})
    except Exception as e:
        print(f"Error al calcular estadísticas: {e}")
        return jsonify({'total': 'Err', 'permitidos': 'Err', 'rechazados': 'Err'})

if __name__ == '__main__':
    webbrowser.open('http://127.0.0.1:5000/')
    app.run(host='127.0.0.1', port=5000)