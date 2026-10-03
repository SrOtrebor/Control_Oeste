import os
import re
from datetime import datetime
import pandas as pd

from config import (
    REGISTROS_DIARIOS_DIR, REGISTROS_FICHAJES_DIR,
    COL_DNI, COL_NOMBRE_APELLIDO, COL_NUM_PERMISO, COL_VENCE, COL_LOCAL, COL_TAREA, COL_TIPO_PERMISO
)
from logger_config import get_logger, log_access_event

import data_manager
from data_manager import formatear_excel
from db_queries import (
    guardar_estado_adentro_sqlite, 
    borrar_estado_adentro_sqlite, 
    loggear_acceso_sqlite,
    verificar_dni_sqlite, 
    registrar_fichaje_sqlite
)
from alert_queries import es_persona_de_interes
from alert_manager import send_alert

# Logger para este módulo
logger = get_logger(__name__)


# Conjunto para llevar registro de personas actualmente dentro
personas_adentro = {}

def restaurar_estado_adentro():
    """Recupera el estado de personas_adentro leyendo los registros de hoy."""
    global personas_adentro
    personas_adentro.clear()
    fecha_actual_str = datetime.now().strftime('%Y-%m-%d')
    
    # 1. Recuperar de registros de ingreso
    archivo_ingreso = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_actual_str}.xlsx')
    if os.path.exists(archivo_ingreso):
        try:
            df = pd.read_excel(archivo_ingreso)
            df['DNI'] = df['DNI'].astype(str)
            for _, row in df.iterrows():
                evento = str(row.get('Evento', ''))
                hora_salida = row.get('Hora_Salida')
                dni = row['DNI']
                if 'Entrada' in evento and (pd.isna(hora_salida) or hora_salida == '' or str(hora_salida).lower() == 'nan'):
                    personas_adentro[dni] = row.get('Tipo_Permiso', 'Desconocido')
                elif 'Salida' in evento and dni in personas_adentro:
                    # En la lógica original la salida actualizaba la fila de entrada,
                    # pero por si hay una fila de salida separada, lo sacamos:
                    personas_adentro.pop(dni, None)
        except Exception as e:
            logger.error(f"Error al restaurar ingresos: {e}")

    # 2. Recuperar de registros de fichaje
    archivo_fichaje = os.path.join(REGISTROS_FICHAJES_DIR, f'registros_fichaje_{fecha_actual_str}.xlsx')
    if os.path.exists(archivo_fichaje):
        try:
            df = pd.read_excel(archivo_fichaje)
            df['DNI'] = df['DNI'].astype(str)
            for _, row in df.iterrows():
                hora_entrada = row.get('Hora_Entrada')
                hora_salida = row.get('Hora_Salida')
                dni = row['DNI']
                if pd.notna(hora_entrada) and hora_entrada != '' and (pd.isna(hora_salida) or hora_salida == ''):
                    personas_adentro[dni] = 'NOMINA'
        except Exception as e:
            logger.error(f"Error al restaurar fichajes: {e}")
            
    logger.info(f"Estado restaurado: {len(personas_adentro)} personas adentro hoy.")

def parsear_codigo_barra(scanner_data):
    """
    Parsea datos de escáner de código de barras / QR de DNI argentino.
    Soporta múltiples formatos:
    
    Formato 1 - QR DNI nuevo (2024+):
      {nro_tramite}"{apellido}"{nombre}"{dni}"{ejemplar}"{fecha_nac}"{fecha_venc}"{JWT}
      Ejemplo: 00754162114"CARDOSO PRESTES"ELIAS MARTIN"44678808"F"31-01-03"06-08-26"eyJ...
    
    Formato 2 - Código de barras DNI viejo:
      "{dni}"{tipo}"{num}"{apellido}"{nombre}"{nacionalidad}"{fecha_nac}"{sexo}"...
      Ejemplo: "32759772    "B"1"GOROSITO"MATIAS DANIEL"ARGENTINA"19-11-1986"M"...
    
    Formato 3 - Código de barras estándar (formato original):
      ..."{apellido}"{nombre}"{sexo}"{dni}"..."{fecha_nac}"{fecha_venc}"...
    
    Formato 4 - @ separado por guiones bajos:
      @APELLIDO_NOMBRE_SEXO_DNI_...
    """
    parts = scanner_data.strip().split('"')
    
    if len(parts) >= 7:
        try:
            # --- Formato 1: QR DNI nuevo ---
            # Detectar: parts[0] es un número de trámite (dígitos), parts[3] es el DNI
            # Estructura: tramite"apellido"nombre"dni"ejemplar"nacimiento"vencimiento"JWT
            if parts[0].strip().isdigit() and len(parts[0].strip()) >= 9:
                dni = parts[3].strip()
                if dni.isdigit() and 7 <= len(dni) <= 8:
                    apellido = parts[1].strip()
                    nombre = parts[2].strip()
                    ejemplar = parts[4].strip()
                    fecha_nacimiento = parts[5].strip() if len(parts) > 5 else ''
                    fecha_vencimiento = parts[6].strip() if len(parts) > 6 else ''
                    logger.debug(f"Formato QR DNI nuevo detectado - DNI: {dni}, Nombre: {nombre} {apellido}")
                    return {
                        'dni': dni, 'nombre': nombre, 'apellido': apellido,
                        'ejemplar': ejemplar,
                        'fecha_nacimiento': fecha_nacimiento,
                        'fecha_vencimiento': fecha_vencimiento,
                        'nombre_completo': f"{nombre} {apellido}"
                    }

            # --- Formato 2: Código de barras DNI viejo ---
            # Detectar: parts[0] es vacío (empieza con "), parts[1] es el DNI
            # Estructura: "dni"tipo"num"apellido"nombre"nacionalidad"nacimiento"sexo"...
            if parts[0].strip() == '' and len(parts) >= 8:
                posible_dni = parts[1].strip()
                if posible_dni.isdigit() and 7 <= len(posible_dni) <= 8:
                    apellido = parts[4].strip() if len(parts) > 4 else ''
                    nombre = parts[5].strip() if len(parts) > 5 else ''
                    sexo = parts[8].strip() if len(parts) > 8 else ''
                    fecha_nacimiento = parts[7].strip() if len(parts) > 7 else ''
                    # Fecha de vencimiento está más adelante en este formato
                    fecha_vencimiento = ''
                    for i in range(9, min(len(parts), 15)):
                        p = parts[i].strip()
                        if re.match(r'^\d{2}-\d{2}-\d{4}$', p):
                            fecha_vencimiento = p
                    logger.debug(f"Formato DNI viejo detectado - DNI: {posible_dni}, Nombre: {nombre} {apellido}")
                    return {
                        'dni': posible_dni, 'nombre': nombre, 'apellido': apellido,
                        'sexo': sexo,
                        'fecha_nacimiento': fecha_nacimiento,
                        'fecha_vencimiento': fecha_vencimiento,
                        'nombre_completo': f"{nombre} {apellido}"
                    }

            # --- Formato 3: Código de barras estándar (original) ---
            # Estructura: ..."{apellido}"{nombre}"{sexo}"{dni}"..."{fecha_nac}"{fecha_venc}"...
            apellido = parts[1].strip()
            nombre = parts[2].strip()
            sexo = parts[3].strip()
            dni = parts[4].strip()
            fecha_nacimiento = parts[6].strip() if len(parts) > 6 else ''
            fecha_vencimiento = parts[7].strip() if len(parts) > 7 else ''
            if dni.isdigit() and 7 <= len(dni) <= 8:
                logger.debug(f"Formato estándar detectado - DNI: {dni}, Nombre: {nombre} {apellido}")
                return {
                    'dni': dni, 'nombre': nombre, 'apellido': apellido, 'sexo': sexo,
                    'fecha_nacimiento': fecha_nacimiento, 'fecha_vencimiento': fecha_vencimiento,
                    'nombre_completo': f"{nombre} {apellido}"
                }
        except IndexError:
            pass

    # --- Formato 4: @ separado por guiones bajos ---
    match = re.search(r'@([^_]+)_([^_]+)_([^_]+)_([^_]+)_', scanner_data)
    if match:
        return {
            'dni': match.group(4), 'apellido': match.group(1), 'nombre': match.group(2),
            'sexo': match.group(3), 'nombre_completo': f"{match.group(2)} {match.group(1)}"
        }

    # --- Fallback: buscar cualquier secuencia de 7-8 dígitos ---
    match_dni = re.search(r'\b(\d{7,8})\b', scanner_data)
    if match_dni:
        return {'dni': match_dni.group(1)}

    # --- Último recurso: extraer todos los dígitos ---
    dni_solo_digitos = re.sub(r'\D', '', scanner_data)
    if dni_solo_digitos:
        return {'dni': dni_solo_digitos}

    return None

def registrar_evento(dni, nombre, hora_evento, evento, tipo_permiso, num_permiso, local, tarea, resultado):
    """
    Registra un evento de acceso en el archivo Excel diario.
    - Consolida registros de Entrada y Salida en la misma fila.
    """
    fecha_actual_str = datetime.now().strftime('%Y-%m-%d')
    nombre_archivo = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_actual_str}.xlsx')
    
    columnas_registro = ['DNI', 'Nombre y Apellido', 'Hora_Ingreso', 'Hora_Salida', 'Evento', 'Tipo_Permiso', 'Num_Permiso', 'Local', 'Tarea', 'Resultado']

    try:
        if os.path.exists(nombre_archivo):
            df_registros = pd.read_excel(nombre_archivo)
            # Asegurar que la columna DNI sea string para la comparación
            df_registros['DNI'] = df_registros['DNI'].astype(str)
        else:
            df_registros = pd.DataFrame(columns=columnas_registro)

        # Lógica para consolidar Entrada y Salida
        if evento == 'Salida':
            # Buscar una entrada previa para el mismo DNI que no tenga Hora_Salida
            # Se busca la última entrada del día para ese DNI
            indices = df_registros[
                (df_registros['DNI'] == str(dni)) & 
                (df_registros['Hora_Salida'].isnull()) &
                (df_registros['Evento'].str.contains('Entrada', na=False))
            ].index

            if not indices.empty:
                # Si se encuentra, actualizar la última entrada con la hora de salida
                indice_a_actualizar = indices[-1]
                df_registros.loc[indice_a_actualizar, 'Hora_Salida'] = hora_evento
                df_registros.loc[indice_a_actualizar, 'Resultado'] = 'Registrado' # Actualizar resultado a 'Registrado'
                df_final = df_registros
            else:
                # Si no hay entrada previa, registrar la salida en una nueva fila (comportamiento anómalo)
                nuevo_registro_df = pd.DataFrame([{'DNI': dni, 'Nombre y Apellido': nombre, 'Hora_Ingreso': '', 'Hora_Salida': hora_evento, 'Evento': evento, 'Tipo_Permiso': tipo_permiso, 'Num_Permiso': num_permiso, 'Local': local, 'Tarea': tarea, 'Resultado': resultado}])
                df_final = pd.concat([df_registros, nuevo_registro_df], ignore_index=True)
        else: # Para 'Entrada OK', 'Entrada RECHAZADA', 'Visita Entrada', 'Visita Salida', etc.
            hora_ingreso_val = hora_evento if 'Salida' not in evento else pd.NA
            hora_salida_val = hora_evento if 'Salida' in evento else pd.NA
            
            nuevo_registro_df = pd.DataFrame([{'DNI': dni, 'Nombre y Apellido': nombre, 'Hora_Ingreso': hora_ingreso_val, 'Hora_Salida': hora_salida_val, 'Evento': evento, 'Tipo_Permiso': tipo_permiso, 'Num_Permiso': num_permiso, 'Local': local, 'Tarea': tarea, 'Resultado': resultado}])
            df_final = pd.concat([df_registros, nuevo_registro_df], ignore_index=True)

        # Reordenar columnas para asegurar consistencia
        df_final = df_final.reindex(columns=columnas_registro)
        
        df_final.to_excel(nombre_archivo, index=False)
        formatear_excel(nombre_archivo)

    except Exception as e:
        logger.error(f"Error crítico al guardar registro en Excel: {e}", exc_info=True)

def registrar_evento_fichaje(dni, nombre, fecha, hora_entrada, hora_salida):
    # (Tu función de fichaje original, no se cambia)
    nombre_archivo = os.path.join(REGISTROS_FICHAJES_DIR, f'registros_fichaje_{fecha}.xlsx')
    columnas = ['DNI', 'Nombre y Apellido', 'Fecha', 'Hora_Entrada', 'Hora_Salida']
    try:
        if os.path.exists(nombre_archivo):
            # Especificar dtype para columnas problemáticas y rellenar NaNs
            df_fichajes = pd.read_excel(nombre_archivo, dtype={
                'DNI': str,
                'Hora_Entrada': str,
                'Hora_Salida': str
            })
            df_fichajes.fillna('', inplace=True)
        else:
            df_fichajes = pd.DataFrame(columns=columnas)
        idx = df_fichajes[(df_fichajes['DNI'] == dni) & (df_fichajes['Fecha'] == fecha)].index
        if not idx.empty:
            df_fichajes.loc[idx[0], 'Hora_Salida'] = hora_salida
        else:
            nuevo_registro = pd.DataFrame([{'DNI': dni, 'Nombre y Apellido': nombre, 'Fecha': fecha, 'Hora_Entrada': hora_entrada, 'Hora_Salida': hora_salida}])
            df_fichajes = pd.concat([df_fichajes, nuevo_registro], ignore_index=True)
        df_fichajes.to_excel(nombre_archivo, index=False)
        formatear_excel(nombre_archivo)
        return True
    except Exception as e:
        logger.error(f"Error crítico al guardar registro de fichaje: {e}", exc_info=True)
        return False

def verificar_dni(scanner_data, mode, puerta="Master", tipo_visita=""):
    logger.debug("Iniciando nueva verificación de DNI")
    parsed_data = parsear_codigo_barra(scanner_data)
    
    if not parsed_data or 'dni' not in parsed_data:
        logger.warning(f"DNI no pudo ser parseado del scanner_data: {scanner_data[:50]}")
        return {'acceso': 'DENEGADO', 'mensaje': 'Formato de DNI no válido o DNI no encontrado.'}

    dni_ingresado_str = parsed_data.get('dni')
    nombre_completo_scanner = parsed_data.get('nombre_completo', 'N/A')
    
    hoy = datetime.now()
    hora_actual_str = hoy.strftime('%H:%M:%S')
    dni_limpio_str = re.sub(r'[\.\s-]', '', str(dni_ingresado_str)).strip()
    logger.debug(f"DNI parseado y limpiado: '{dni_limpio_str}'")

    # --- CHECK LISTA NEGRA (ALERTA SILENCIOSA O DIRECTA) ---
    es_interes, motivo_interes = es_persona_de_interes(dni_limpio_str)
    if es_interes:
        send_alert(dni_limpio_str, nombre_completo_scanner, f"ALERTA (Lista Negra): {motivo_interes}", nivel="danger")

    # --- LÓGICA DE SALIDA / VISITA (MIGRADA A SQLITE LOGGING) ---
    if mode == 'salida':
        if dni_limpio_str in personas_adentro:
            entry_type = personas_adentro.pop(dni_limpio_str)

            if entry_type == 'VISITA':
                registrar_evento(dni_limpio_str, nombre_completo_scanner, hora_actual_str, 'Visita Salida', 'VISITA', 'N/A', 'N/A', 'Visita', 'REGISTRADO')
                loggear_acceso_sqlite(dni_limpio_str, nombre_completo_scanner, 'SALIDA', 'VISITA', 'N/A', 'N/A', 'N/A', 'REGISTRADO', puerta=puerta, tipo_visita=tipo_visita)
            else:
                registrar_evento(dni_limpio_str, nombre_completo_scanner, hora_actual_str, 'Salida', 'N/A', 'N/A', 'N/A', 'Salida', 'REGISTRADO')
                loggear_acceso_sqlite(dni_limpio_str, nombre_completo_scanner, 'SALIDA', entry_type, 'N/A', 'N/A', 'N/A', 'REGISTRADO', puerta=puerta, tipo_visita="")
            
            borrar_estado_adentro_sqlite(dni_limpio_str)
            return {'acceso': 'PERMITIDO', 'mensaje': 'Salida Registrada', 'nombre': nombre_completo_scanner}
        else:
            if not es_interes:
                send_alert(dni_limpio_str, nombre_completo_scanner, "Intento de salida, pero no estaba registrado adentro.", nivel="warning")
            return {'acceso': 'DENEGADO', 'mensaje': 'Error: Persona no registrada adentro', 'nombre': ''}

    if mode == 'visita':
        personas_adentro[dni_limpio_str] = 'VISITA'
        registrar_evento(dni_limpio_str, nombre_completo_scanner, hora_actual_str, 'Visita Entrada', 'VISITA', 'N/A', 'N/A', 'Visita', 'AUTORIZADO')
        loggear_acceso_sqlite(dni_limpio_str, nombre_completo_scanner, 'ENTRADA_VISITA', 'VISITA', 'N/A', 'N/A', 'N/A', 'AUTORIZADO', puerta=puerta, tipo_visita=tipo_visita)
        guardar_estado_adentro_sqlite(dni_limpio_str, nombre_completo_scanner, hora_actual_str, 'VISITA', 'N/A', puerta=puerta)
        return {'acceso': 'PERMITIDO', 'mensaje': 'Visita Registrada', 'nombre': nombre_completo_scanner}

    # --- LÓGICA DE ENTRADA (MIGRADA A SQLITE + IRSA DIGITAL) ---
    if mode == 'entrada':
        logger.debug("Modo 'entrada' seleccionado. Verificando...")
        
        permiso = None
        
        # 1. Validar en listado directo de IRSA (FAOs y FAPs digitales)
        try:
            directorio_irsa = data_manager.obtener_directorio_irsa()
            for p in directorio_irsa:
                if re.sub(r'[\.\s-]', '', str(p.get('dni', ''))).strip() == dni_limpio_str:
                    permiso = {
                        'nombre': p.get('nombre', 'Desconocido'),
                        'local': p.get('empresa', 'N/A'),
                        'tarea': 'Trabajos Generales',
                        'vence': p.get('vence', 'Indefinido'),
                        'tipo_permiso': p.get('tipo', 'IRSA'),
                        'num_permiso': p.get('numero', 'N/A')
                    }
                    logger.debug(f"Permiso encontrado en IRSA Digital: {permiso}")
                    break
        except Exception as e:
            logger.error(f"Error al verificar en directorio IRSA digital: {e}")

        # 2. Si no está en IRSA directo, buscar en base de datos SQLite (Nóminas, Excepciones, etc)
        if not permiso:
            permiso = verificar_dni_sqlite(dni_limpio_str)
        
        if permiso:
            nombre = permiso['nombre']
            local = permiso['local']
            tarea = permiso['tarea']
            vence_str = permiso['vence']
            tipo_permiso = permiso['tipo_permiso']
            num_permiso = permiso['num_permiso']
            
            logger.info(f"Acceso PERMITIDO - DNI: {dni_limpio_str}, Nombre: {nombre}, Tipo: {tipo_permiso}")
            log_access_event(dni_limpio_str, nombre, 'PERMITIDO', tipo_permiso, f'Local: {local}, Vence: {vence_str}')
            
            personas_adentro[dni_limpio_str] = tipo_permiso
            
            # Dual Logging: Guardamos en Excel (antiguo) y en SQLite (nuevo)
            registrar_evento(dni_limpio_str, nombre, hora_actual_str, 'Entrada OK', tipo_permiso, num_permiso, local, tarea, 'AUTORIZADO')
            loggear_acceso_sqlite(dni_limpio_str, nombre, 'ENTRADA', tipo_permiso, num_permiso, local, tarea, 'AUTORIZADO', puerta=puerta, tipo_visita="")
            guardar_estado_adentro_sqlite(dni_limpio_str, nombre, hora_actual_str, tipo_permiso, local, puerta=puerta)
            
            return {
                'acceso': 'PERMITIDO', 'nombre': nombre, 'mensaje': f'ACCESO PERMITIDO ({tipo_permiso}): {nombre}', 
                'tipo_permiso': tipo_permiso, 'num_permiso': num_permiso, 'local': local, 'tarea': tarea, 'vence': vence_str
            }
        else:
            logger.warning(f"Acceso DENEGADO - DNI: {dni_limpio_str} no encontrado o sin permiso vigente")
            log_access_event(dni_limpio_str, 'No Autorizado', 'DENEGADO', 'N/A', 'Sin permiso válido')
            mensaje = f'ACCESO DENEGADO: DNI {dni_limpio_str} no encontrado o sin permiso vigente.'
            
            if not es_interes: # Si ya mandamos alerta por lista negra, no duplicamos
                send_alert(dni_limpio_str, nombre_completo_scanner, "ACCESO DENEGADO: Sin permiso válido", nivel="warning")
            
            registrar_evento(dni_limpio_str, 'No Autorizado', hora_actual_str, 'Entrada RECHAZADA', 'N/A', 'N/A', 'N/A', 'N/A', 'DENEGADO')
            loggear_acceso_sqlite(dni_limpio_str, 'No Autorizado', 'ENTRADA_RECHAZADA', 'N/A', 'N/A', 'N/A', 'N/A', 'DENEGADO', 'Sin permiso válido', puerta=puerta)
            
            return {'acceso': 'DENEGADO', 'nombre': 'No Autorizado', 'mensaje': mensaje}

    return {'acceso': 'DENEGADO', 'mensaje': 'Modo no reconocido.'}

def registrar_fichaje(scanner_data, mode, puerta="Master", puesto_seleccionado=""):
    parsed_data = parsear_codigo_barra(scanner_data)
    if not parsed_data or 'dni' not in parsed_data:
        return {'acceso': 'DENEGADO', 'mensaje': 'DNI no válido.', 'nombre': ''}

    dni = parsed_data.get('dni')
    now = datetime.now()
    fecha_hoy_str = now.strftime('%Y-%m-%d')
    hora_actual_str = now.strftime('%H:%M:%S')

    data_manager.cargar_autorizaciones()
    nombre_completo = parsed_data.get('nombre_completo', 'N/A')
    
    df_nominas_persistentes = data_manager.get_df_nominas_persistentes()
    persona_nomina = df_nominas_persistentes[df_nominas_persistentes[COL_DNI] == dni]
    if not persona_nomina.empty:
        nombre_completo = persona_nomina.iloc[0].get('Nombre y Apellido', nombre_completo)
    
    nombre_archivo_fichajes = os.path.join(REGISTROS_FICHAJES_DIR, f'registros_fichaje_{fecha_hoy_str}.xlsx')

    if mode == 'punch-in':
        personas_adentro[dni] = 'NOMINA'
        registrar_evento_fichaje(dni, nombre_completo, fecha_hoy_str, hora_actual_str, '')
        registrar_fichaje_sqlite(dni, nombre_completo, 'ENTRADA', puerta=puerta, puesto_seleccionado=puesto_seleccionado)
        registrar_evento(dni, nombre_completo, hora_actual_str, 'Fichaje Entrada', 'NOMINA', 'N/A', 'N/A', 'Fichaje', 'AUTORIZADO')
        loggear_acceso_sqlite(dni, nombre_completo, 'FICHAJE_ENTRADA', 'NOMINA', 'N/A', 'N/A', 'N/A', 'AUTORIZADO', puerta=puerta)
        guardar_estado_adentro_sqlite(dni, nombre_completo, hora_actual_str, 'NOMINA', 'N/A', puerta=puerta)
        
        return {
            'acceso': 'PERMITIDO', 'mensaje': 'Entrada Registrada Correctamente',
            'nombre': nombre_completo, 'hora_entrada': hora_actual_str
        }

    elif mode == 'punch-out':
        if not os.path.exists(nombre_archivo_fichajes):
            return {'acceso': 'DENEGADO', 'mensaje': 'Error: No hay registros de entrada hoy.', 'nombre': nombre_completo}
        df_fichajes_hoy = pd.read_excel(nombre_archivo_fichajes)
        df_fichajes_hoy['DNI'] = df_fichajes_hoy['DNI'].astype(str)
        
        # Logica original
        if dni not in df_fichajes_hoy['DNI'].values:
             return {'acceso': 'DENEGADO', 'mensaje': 'Error: No se encontró registro de entrada para hoy.', 'nombre': nombre_completo}
             
        registrar_evento_fichaje(dni, nombre_completo, fecha_hoy_str, '', hora_actual_str)
        registrar_fichaje_sqlite(dni, nombre_completo, 'SALIDA', puerta=puerta)
        registrar_evento(dni, nombre_completo, hora_actual_str, 'Fichaje Salida', 'NOMINA', 'N/A', 'N/A', 'Fichaje', 'REGISTRADO')
        loggear_acceso_sqlite(dni, nombre_completo, 'FICHAJE_SALIDA', 'NOMINA', 'N/A', 'N/A', 'N/A', 'REGISTRADO', puerta=puerta)
        borrar_estado_adentro_sqlite(dni)
        registro_entrada = df_fichajes_hoy[df_fichajes_hoy['DNI'] == dni]
        if registro_entrada.empty:
            return {'acceso': 'DENEGADO', 'mensaje': 'Error: No se encontró registro de entrada para hoy.', 'nombre': nombre_completo}
        personas_adentro.pop(dni, None)
        hora_entrada = registro_entrada.iloc[0].get('Hora_Entrada', 'N/A')
        registrar_evento_fichaje(dni, nombre_completo, fecha_hoy_str, hora_entrada, hora_actual_str)
        registrar_evento(dni, nombre_completo, hora_actual_str, 'Fichaje Salida', 'NOMINA', 'N/A', 'N/A', 'Fichaje', 'REGISTRADO')
        return {
            'acceso': 'PERMITIDO', 'mensaje': 'Salida Registrada Correctamente',
            'nombre': nombre_completo, 'hora_entrada': hora_entrada, 'hora_salida': hora_actual_str
        }

    return {'acceso': 'DENEGADO', 'mensaje': 'Modo de fichaje no reconocido.', 'nombre': ''}