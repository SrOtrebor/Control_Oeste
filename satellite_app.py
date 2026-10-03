import os
import requests
from flask import Flask, render_template, request, jsonify
from satellite_db import (
    registrar_acceso, obtener_registros_pendientes, marcar_registros_como_sincronizados,
    actualizar_cache, actualizar_ultima_sincronizacion, obtener_config, get_db_connection
)
import logging
from datetime import datetime

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger('SatelliteApp')

def get_master_url():
    config = obtener_config()
    return config.get('master_ip', 'http://127.0.0.1:5000')

@app.route('/')
def index():
    config = obtener_config()
    return render_template('satellite_index.html', config=config)

@app.route('/verificar_dni', methods=['POST'])
def verificar_dni():
    data = request.json
    dni = data.get('dni')
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # Lista Negra
        cursor.execute("SELECT motivo FROM cache_lista_negra WHERE dni = ?", (dni,))
        bloqueado = cursor.fetchone()
        if bloqueado:
            return jsonify({
                'estado': 'rechazado', 
                'motivo': f'Bloqueado: {bloqueado["motivo"]}'
            })
            
        # Nómina Cache
        cursor.execute("SELECT * FROM cache_nominas WHERE dni = ?", (dni,))
        nomina = cursor.fetchone()
        
        if nomina:
            if nomina['activo'] == 0:
                return jsonify({
                    'estado': 'rechazado',
                    'motivo': 'Inactivo en RRHH'
                })
            
            return jsonify({
                'estado': 'permitido',
                'datos': {
                    'nombre': nomina['nombre'],
                    'empresa': 'Interno',
                    'local': nomina['puesto'],
                    'tarea': nomina['categoria']
                }
            })
            
        # Si no está en nómina y no está en lista negra, se deja a discreción como Visita
        return jsonify({
            'estado': 'desconocido',
            'motivo': 'No registrado en base de datos local'
        })


@app.route('/registrar_fichaje', methods=['POST'])
def registrar_fichaje():
    data = request.json
    dni = data.get('dni', '').strip()
    accion = data.get('accion') # entrada, salida, visita_ingreso, visita_salida
    
    if not dni or not accion:
        return jsonify({'success': False, 'message': 'Faltan datos'})
        
    estado = ''
    if accion == 'entrada': estado = 'Entrada'
    elif accion == 'salida': estado = 'Salida'
    elif accion == 'visita_ingreso': estado = 'Visita (Entrada)'
    elif accion == 'visita_salida': estado = 'Visita (Salida)'
    else: estado = accion
    
    tipo_visita = ''
    
    # Validar cache local (Lista Negra y Nómina) solo para entradas regulares
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        if accion == 'entrada':
            cursor.execute("SELECT motivo FROM cache_lista_negra WHERE dni = ?", (dni,))
            bloqueado = cursor.fetchone()
            if bloqueado:
                registrar_acceso(dni, 'Rechazo (Lista Negra)', tipo_visita)
                return jsonify({'success': False, 'message': f'DNI bloqueado: {bloqueado["motivo"]}'})
        
            cursor.execute("SELECT activo FROM cache_nominas WHERE dni = ?", (dni,))
            nomina = cursor.fetchone()
            if nomina and nomina['activo'] == 0:
                registrar_acceso(dni, 'Rechazo (Inactivo)', tipo_visita)
                return jsonify({'success': False, 'message': 'Empleado inactivo en RRHH'})

    # Todo OK, registrar
    registrar_acceso(dni, estado, tipo_visita)
    return jsonify({'success': True})

@app.route('/get_daily_records')
def get_daily_records():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        hoy = datetime.now().strftime('%Y-%m-%d')
        cursor.execute("SELECT * FROM registros_locales WHERE timestamp LIKE ? ORDER BY id DESC LIMIT 50", (f"{hoy}%",))
        rows = [dict(row) for row in cursor.fetchall()]
        
        records = []
        for r in rows:
            # Fake names for offline if needed, or query cache
            cursor.execute("SELECT nombre, puesto FROM cache_nominas WHERE dni = ?", (r['dni'],))
            n = cursor.fetchone()
            records.append({
                'hora': r['timestamp'].split('T')[1][:5] if 'T' in r['timestamp'] else r['timestamp'].split(' ')[1][:5],
                'dni': r['dni'],
                'nombre': n['nombre'] if n else 'Desconocido',
                'estado': r['estado'],
                'permiso': 'N/A',
                'local': n['puesto'] if n else ''
            })
            
        return jsonify({'records': records})

@app.route('/get_dynamic_stats')
def get_dynamic_stats():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        hoy = datetime.now().strftime('%Y-%m-%d')
        cursor.execute("SELECT estado, COUNT(*) as count FROM registros_locales WHERE timestamp LIKE ? GROUP BY estado", (f"{hoy}%",))
        counts = {r['estado']: r['count'] for r in cursor.fetchall()}
        
        entradas = sum([c for e, c in counts.items() if 'Entrada' in e])
        salidas = sum([c for e, c in counts.items() if 'Salida' in e])
        rechazos = sum([c for e, c in counts.items() if 'Rechazo' in e])
        
        return jsonify({
            'totalAdentro': max(0, entradas - salidas),
            'permitidos': entradas + salidas,
            'rechazados': rechazos
        })

@app.route('/api/sync', methods=['POST'])
def api_sync():
    """Sincroniza push y pull con el servidor Master"""
    master_url = get_master_url()
    try:
        pendientes = obtener_registros_pendientes()
        if pendientes:
            res_push = requests.post(f"{master_url}/api/sync/receive_satellite_data", json={'registros': pendientes}, timeout=10)
            if res_push.status_code == 200 and res_push.json().get('success'):
                ids_sincronizados = [r['id'] for r in pendientes]
                marcar_registros_como_sincronizados(ids_sincronizados)
        
        res_puestos = requests.get(f"{master_url}/api/config/puestos", timeout=10)
        if res_puestos.status_code == 200:
            puestos = res_puestos.json().get('data', [])
            actualizar_cache('cache_puestos', puestos)
            
        res_nominas = requests.get(f"{master_url}/api/dashboard/nominas", timeout=10)
        if res_nominas.status_code == 200:
            nominas = res_nominas.json().get('data', [])
            cache_nom = []
            for d in nominas:
                cache_nom.append({'dni': str(d.get('dni', '')), 'nombre': d.get('nombre', ''), 'categoria': d.get('categoria', ''), 'puesto': d.get('puesto', ''), 'activo': 1})
            if cache_nom:
                actualizar_cache('cache_nominas', cache_nom, 'dni')
            
        res_ln = requests.get(f"{master_url}/api/dashboard/lista_negra", timeout=10)
        if res_ln.status_code == 200:
            ln = res_ln.json().get('data', [])
            cache_ln = [{'dni': str(d.get('dni', '')), 'motivo': d.get('motivo', '')} for d in ln]
            if cache_ln:
                actualizar_cache('cache_lista_negra', cache_ln, 'dni')
            
        ultima_fecha = actualizar_ultima_sincronizacion()
        return jsonify({'success': True, 'last_sync': ultima_fecha})
    
    except requests.exceptions.RequestException as e:
        logger.error(f"Error de red al sincronizar: {e}")
        return jsonify({'success': False, 'message': 'No hay conexión con el Master'})

if __name__ == '__main__':
    print("Iniciando Módulo Satélite en puerto 5001...")
    app.run(host='0.0.0.0', port=5001, debug=True)
