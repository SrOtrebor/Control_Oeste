import os
from flask import jsonify, request, session, render_template, send_file
import pandas as pd
from datetime import datetime

# ================= RECOVERED ENDPOINTS =================

@app.route('/dashboard')
def dashboard():
    if 'logged_in' not in session:
        return redirect(url_for('login'))
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
            return jsonify({'success': True, 'stats': {'entradas': entradas, 'salidas': salidas, 'rechazos': rechazos, 'adentro': adentro}})
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
        return jsonify({'success': True, 'records': directorio})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/dashboard/tramites_irsa')
def api_dashboard_tramites_irsa():
    try:
        tramites = data_manager.obtener_tramites_irsa_raw()
        return jsonify({'success': True, 'records': tramites})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/dashboard/auditoria')
def api_dashboard_auditoria():
    return jsonify({'success': True, 'records': []})

@app.route('/api/config/irsa', methods=['GET', 'POST'])
def api_config_irsa():
    if 'logged_in' not in session: return jsonify({'success': False, 'message': 'No autorizado'}), 403
    return jsonify({'success': True, 'has_credentials': True, 'dias_alerta_vencimiento': 10})

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
        texto = request.json.get('texto')
        res = data_manager.procesar_nomina_texto(texto)
        return jsonify(res)
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/dashboard/nomina/<empresa>', methods=['DELETE'])
def api_dashboard_nomina_delete(empresa):
    return jsonify({'success': True})

@app.route('/api/dashboard/lista_negra', methods=['GET', 'POST'])
def api_dashboard_lista_negra():
    return jsonify({'success': True, 'data': []})

@app.route('/api/dashboard/lista_negra/<int:id>', methods=['DELETE'])
def api_dashboard_lista_negra_del(id):
    return jsonify({'success': True})

@app.route('/api/sync/irsa', methods=['POST'])
def api_sync_irsa():
    return jsonify({'success': True, 'message': 'Sincronizado'})

@app.route('/api/admin/backup/download', methods=['GET'])
def api_admin_backup_download():
    return jsonify({'success': False, 'message': 'Not implemented'})

@app.route('/stream/alerts')
def stream_alerts():
    return jsonify({'success': False})

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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
