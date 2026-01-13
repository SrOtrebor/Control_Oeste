"""
Nuevos endpoints para descarga de reportes por fechas
Agregar al final de app.py antes de if __name__ == '__main__':
"""

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
            archivo_dia = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_{fecha_str}.xlsx')
            
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
            archivo_dia = os.path.join(REGISTROS_FICHAJES_DIR, f'fichajes_{fecha_str}.xlsx')
            
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
