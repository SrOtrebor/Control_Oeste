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

