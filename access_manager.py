import os
import re
from datetime import datetime
import pandas as pd

from config import (
    REGISTROS_DIARIOS_DIR, REGISTROS_FICHAJES_DIR,
    COL_DNI, COL_NOMBRE_APELLIDO, COL_NUM_PERMISO, COL_VENCE, COL_LOCAL, COL_TAREA, COL_TIPO_PERMISO
)

# --- INICIO DE LA CORRECCIÓN ---
# Importamos el MÓDULO 'data_manager' en lugar de las variables
import data_manager
from data_manager import cargar_autorizaciones
# --- FIN DE LA CORRECCIÓN ---


# Conjunto para llevar registro de personas actualmente dentro
personas_adentro = set()

def parsear_codigo_barra(scanner_data):
    # (Tu lógica de parseo original está aquí, no se cambia)
    parts = scanner_data.strip().split('"')
    if len(parts) >= 8:
        try:
            apellido = parts[1].strip()
            nombre = parts[2].strip()
            sexo = parts[3].strip()
            dni = parts[4].strip()
            fecha_nacimiento = parts[6].strip()
            fecha_vencimiento = parts[7].strip()
            if dni.isdigit() and 7 <= len(dni) <= 8:
                return {
                    'dni': dni, 'nombre': nombre, 'apellido': apellido, 'sexo': sexo,
                    'fecha_nacimiento': fecha_nacimiento, 'fecha_vencimiento': fecha_vencimiento,
                    'nombre_completo': f"{nombre} {apellido}"
                }
        except IndexError:
            pass

    match = re.search(r'@([^_]+)_([^_]+)_([^_]+)_([^_]+)_', scanner_data)
    if match:
        return {
            'dni': match.group(4), 'apellido': match.group(1), 'nombre': match.group(2),
            'sexo': match.group(3), 'nombre_completo': f"{match.group(2)} {match.group(1)}"
        }

    match_dni = re.search(r'\b(\d{7,8})\b', scanner_data)
    if match_dni:
        return {'dni': match_dni.group(1)}

    dni_solo_digitos = re.sub(r'\D', '', scanner_data)
    if dni_solo_digitos:
        return {'dni': dni_solo_digitos}

    return None

def registrar_evento(dni, nombre, hora_ingreso, evento, tipo_permiso, num_permiso, local, tarea, resultado):
    # (Tu función de registro original, no se cambia)
    fecha_actual_str = datetime.now().strftime('%Y-%m-%d')
    nombre_archivo = os.path.join(REGISTROS_DIARIOS_DIR, f'registros_ingreso_{fecha_actual_str}.xlsx')
    columnas_registro = ['DNI', 'Nombre y Apellido', 'Hora_Ingreso', 'Evento', 'Tipo_Permiso', 'Num_Permiso', 'Local', 'Tarea', 'Resultado']
    nuevo_registro_df = pd.DataFrame([{'DNI': dni, 'Nombre y Apellido': nombre, 'Hora_Ingreso': hora_ingreso, 'Evento': evento, 'Tipo_Permiso': tipo_permiso, 'Num_Permiso': num_permiso, 'Local': local, 'Tarea': tarea, 'Resultado': resultado}])
    try:
        if not os.path.exists(nombre_archivo):
            df_registros = pd.DataFrame(columns=columnas_registro)
        else:
            df_registros = pd.read_excel(nombre_archivo)
        df_final = pd.concat([df_registros, nuevo_registro_df], ignore_index=True)
        df_final.to_excel(nombre_archivo, index=False)
    except Exception as e:
        print(f"Error Crítico al guardar registro en Excel: {e}")

def registrar_evento_fichaje(dni, nombre, fecha, hora_entrada, hora_salida):
    # (Tu función de fichaje original, no se cambia)
    nombre_archivo = os.path.join(REGISTROS_FICHAJES_DIR, f'registros_fichaje_{fecha}.xlsx')
    columnas = ['DNI', 'Nombre y Apellido', 'Fecha', 'Hora_Entrada', 'Hora_Salida']
    try:
        if os.path.exists(nombre_archivo):
            df_fichajes = pd.read_excel(nombre_archivo)
            df_fichajes['DNI'] = df_fichajes['DNI'].astype(str)
        else:
            df_fichajes = pd.DataFrame(columns=columnas)
        idx = df_fichajes[(df_fichajes['DNI'] == dni) & (df_fichajes['Fecha'] == fecha)].index
        if not idx.empty:
            df_fichajes.loc[idx[0], 'Hora_Salida'] = hora_salida
        else:
            nuevo_registro = pd.DataFrame([{'DNI': dni, 'Nombre y Apellido': nombre, 'Fecha': fecha, 'Hora_Entrada': hora_entrada, 'Hora_Salida': hora_salida}])
            df_fichajes = pd.concat([df_fichajes, nuevo_registro], ignore_index=True)
        df_fichajes.to_excel(nombre_archivo, index=False)
        return True
    except Exception as e:
        print(f"Error Crítico al guardar registro de fichaje: {e}")
        return False

def verificar_dni(scanner_data, mode):
    print("\n--- INICIANDO NUEVA VERIFICACIÓN (CON DEPURACIÓN EXTENDIDA) ---")
    parsed_data = parsear_codigo_barra(scanner_data)
    
    if not parsed_data or 'dni' not in parsed_data:
        print("DEBUG: DNI no pudo ser parseado del scanner_data.")
        return {'acceso': 'DENEGADO', 'mensaje': 'Formato de DNI no válido o DNI no encontrado.'}

    dni_ingresado_str = parsed_data.get('dni')
    nombre_completo_scanner = parsed_data.get('nombre_completo', 'N/A')
    
    hoy = datetime.now()
    hora_actual_str = hoy.strftime('%H:%M:%S')
    dni_limpio_str = re.sub(r'[\.\s-]', '', str(dni_ingresado_str)).strip()
    print(f"DEBUG: DNI parseado y limpiado para búsqueda: '{dni_limpio_str}' (Tipo: {type(dni_limpio_str)})")

    # --- LÓGICA DE SALIDA / VISITA (sin cambios) ---
    if mode == 'salida':
        if dni_limpio_str in personas_adentro:
            personas_adentro.discard(dni_limpio_str)
            registrar_evento(dni_limpio_str, nombre_completo_scanner, hora_actual_str, 'Salida', 'N/A', 'N/A', 'N/A', 'Salida', 'NEUTRO')
            return {'acceso': 'PERMITIDO', 'mensaje': 'Salida Registrada', 'nombre': nombre_completo_scanner}
        else:
            return {'acceso': 'DENEGADO', 'mensaje': 'Error: Persona no registrada adentro', 'nombre': ''}

    if mode == 'visita':
        personas_adentro.add(dni_limpio_str)
        registrar_evento(dni_limpio_str, nombre_completo_scanner, hora_actual_str, 'Visita Entrada', 'VISITA', 'N/A', 'N/A', 'Visita', 'VERDE')
        return {'acceso': 'PERMITIDO', 'mensaje': 'Visita Registrada', 'nombre': nombre_completo_scanner}

    # --- LÓGICA DE ENTRADA (CORREGIDA) ---
    if mode == 'entrada':
        print("DEBUG: Modo 'entrada' seleccionado. Verificando todas las listas.")
        cargar_autorizaciones() 
        
        # --- 1. Verificar en Nóminas Persistentes ---
        print("\nDEBUG: 1. Verificando en Nóminas Persistentes...")
        df_nominas_persistentes = data_manager.get_df_nominas_persistentes()
        if not df_nominas_persistentes.empty and COL_DNI in df_nominas_persistentes.columns:
            match = df_nominas_persistentes[df_nominas_persistentes[COL_DNI] == dni_limpio_str]
            if not match.empty:
                print("   - DNI ENCONTRADO en Nóminas Persistentes.")
                persona = match.iloc[0]
                desde = persona.get('Vigencia Desde')
                hasta = persona.get('Vigencia Hasta')
                if pd.notna(desde) and pd.notna(hasta) and pd.Timestamp(desde) <= hoy <= pd.Timestamp(hasta):
                    print("   - Vigencia Nómina Persistente VÁLIDA. ACCESO PERMITIDO.")
                    nombre = persona.get(COL_NOMBRE_APELLIDO, 'N/A')
                    local = persona.get(COL_LOCAL, 'N/A')
                    tarea = persona.get(COL_TAREA, 'N/A')
                    vence = pd.Timestamp(hasta).strftime('%d/%m/%Y')
                    personas_adentro.add(dni_limpio_str)
                    registrar_evento(dni_limpio_str, nombre, hora_actual_str, 'Entrada OK', 'Nomina Persistente', 'N/A', local, tarea, 'VERDE')
                    return {'acceso': 'PERMITIDO', 'nombre': nombre, 'mensaje': f'ACCESO PERMITIDO (Nomina): {nombre}', 'tipo_permiso': 'Nomina Persistente', 'num_permiso': 'N/A', 'local': local, 'tarea': tarea, 'vence': vence}
                else:
                    print(f"   - Permiso encontrado en Nómina Persistente pero está vencido o las fechas son inválidas.")
            else:
                print("   - DNI no encontrado en Nóminas Persistentes.")

        # --- 2. Verificar en lista FAP ---
        print("\nDEBUG: 2. Verificando en FAP...")
        if not data_manager.df_fap.empty and COL_DNI in data_manager.df_fap.columns:
            match = data_manager.df_fap[data_manager.df_fap[COL_DNI] == dni_limpio_str]
            if not match.empty:
                print("   - DNI ENCONTRADO en FAP.")
                persona = match.iloc[0]
                vence_val = persona.get(COL_VENCE)
                if pd.notna(vence_val) and hoy.date() <= pd.Timestamp(vence_val).date():
                    print("   - Permiso FAP VÁLIDO. ACCESO PERMITIDO.")
                    nombre = persona.get(COL_NOMBRE_APELLIDO, 'N/A')
                    tipo_permiso = persona.get(COL_TIPO_PERMISO, 'FAP')
                    num_permiso = persona.get(COL_NUM_PERMISO, 'N/A')
                    local = persona.get(COL_LOCAL, 'N/A')
                    tarea = persona.get(COL_TAREA, 'N/A')
                    vence_str = pd.Timestamp(vence_val).strftime('%d/%m/%Y')
                    personas_adentro.add(dni_limpio_str)
                    registrar_evento(dni_limpio_str, nombre, hora_actual_str, 'Entrada OK', tipo_permiso, num_permiso, local, tarea, 'VERDE')
                    return {'acceso': 'PERMITIDO', 'nombre': nombre, 'mensaje': f'ACCESO PERMITIDO (FAP): {nombre}', 'tipo_permiso': tipo_permiso, 'num_permiso': num_permiso, 'local': local, 'tarea': tarea, 'vence': vence_str}
                else:
                    print(f"   - Permiso encontrado en FAP pero está vencido.")
            else:
                print("   - DNI no encontrado en FAP.")

        # --- 3. Verificar en lista FAO ---
        print("\nDEBUG: 3. Verificando en FAO...")
        if not data_manager.df_fao.empty and COL_DNI in data_manager.df_fao.columns:
            match = data_manager.df_fao[data_manager.df_fao[COL_DNI] == dni_limpio_str]
            if not match.empty:
                print("   - DNI ENCONTRADO en FAO.")
                persona = match.iloc[0]
                vence_val = persona.get(COL_VENCE)
                if pd.notna(vence_val) and hoy.date() <= pd.Timestamp(vence_val).date():
                    print("   - Permiso FAO VÁLIDO. ACCESO PERMITIDO.")
                    nombre = persona.get(COL_NOMBRE_APELLIDO, 'N/A')
                    tipo_permiso = persona.get(COL_TIPO_PERMISO, 'FAO')
                    num_permiso = persona.get(COL_NUM_PERMISO, 'N/A')
                    local = persona.get(COL_LOCAL, 'N/A')
                    tarea = persona.get(COL_TAREA, 'N/A')
                    vence_str = pd.Timestamp(vence_val).strftime('%d/%m/%Y')
                    personas_adentro.add(dni_limpio_str)
                    registrar_evento(dni_limpio_str, nombre, hora_actual_str, 'Entrada OK', tipo_permiso, num_permiso, local, tarea, 'VERDE')
                    return {'acceso': 'PERMITIDO', 'nombre': nombre, 'mensaje': f'ACCESO PERMITIDO (FAO): {nombre}', 'tipo_permiso': tipo_permiso, 'num_permiso': num_permiso, 'local': local, 'tarea': tarea, 'vence': vence_str}
                else:
                    print(f"   - Permiso encontrado en FAO pero está vencido.")
            else:
                print("   - DNI no encontrado en FAO.")

        # --- 4. Verificar en lista de excepciones ---
        print("\nDEBUG: 4. Verificando en Excepciones...")
        if not data_manager.df_excepciones.empty and COL_DNI in data_manager.df_excepciones.columns:
            match = data_manager.df_excepciones[data_manager.df_excepciones[COL_DNI] == dni_limpio_str]
            if not match.empty:
                print("   - DNI ENCONTRADO en Excepciones.")
                excepcion = match.iloc[0]
                vence_val = excepcion.get(COL_VENCE)
                if pd.notna(vence_val) and hoy.date() <= pd.Timestamp(vence_val).date():
                    print("   - Excepción VÁLIDA. ACCESO PERMITIDO.")
                    nombre = excepcion.get(COL_NOMBRE_APELLIDO, 'N/A')
                    local = excepcion.get(COL_LOCAL, 'N/A')
                    vence_str = pd.Timestamp(vence_val).strftime('%d/%m/%Y')
                    personas_adentro.add(dni_limpio_str)
                    registrar_evento(dni_limpio_str, nombre, hora_actual_str, 'Entrada OK', 'Excepcion', 'N/A', local, 'N/A', 'VERDE')
                    return {'acceso': 'PERMITIDO', 'nombre': nombre, 'mensaje': f'ACCESO PERMITIDO (Excepción): {nombre}', 'tipo_permiso': 'Excepcion', 'num_permiso': 'N/A', 'local': local, 'tarea': 'N/A', 'vence': vence_str}
                else:
                    print(f"   - Permiso encontrado en Excepciones pero está vencido.")
            else:
                print("   - DNI no encontrado en Excepciones.")

        # --- 5. Decisión final si no se encontró permiso válido ---
        print(f"\nDEBUG: 5. Decisión final para DNI '{dni_limpio_str}'. No se encontró permiso válido en ninguna lista.")
        mensaje = f'ACCESO DENEGADO: DNI {dni_limpio_str} no encontrado o sin permiso vigente.'
        registrar_evento(dni_limpio_str, 'No Autorizado', hora_actual_str, 'Entrada RECHAZADA', 'N/A', 'N/A', 'N/A', 'N/A', 'ROJO')
        return {'acceso': 'DENEGADO', 'nombre': 'No Autorizado', 'mensaje': mensaje}

    return {'acceso': 'DENEGADO', 'mensaje': 'Modo no reconocido.'}

def registrar_fichaje(scanner_data, mode):
    parsed_data = parsear_codigo_barra(scanner_data)
    if not parsed_data or 'dni' not in parsed_data:
        return {'acceso': 'DENEGADO', 'mensaje': 'DNI no válido.', 'nombre': ''}

    dni = parsed_data.get('dni')
    now = datetime.now()
    fecha_hoy_str = now.strftime('%Y-%m-%d')
    hora_actual_str = now.strftime('%H:%M:%S')

    cargar_autorizaciones()
    nombre_completo = parsed_data.get('nombre_completo', 'N/A')
    
    df_nominas_persistentes = data_manager.get_df_nominas_persistentes()
    persona_nomina = df_nominas_persistentes[df_nominas_persistentes[COL_DNI] == dni]
    if not persona_nomina.empty:
        nombre_completo = persona_nomina.iloc[0].get('Nombre y Apellido', nombre_completo)
    
    nombre_archivo_fichajes = os.path.join(REGISTROS_FICHAJES_DIR, f'registros_fichaje_{fecha_hoy_str}.xlsx')

    if mode == 'punch-in':
        personas_adentro.add(dni)
        registrar_evento_fichaje(dni, nombre_completo, fecha_hoy_str, hora_actual_str, '')
        return {
            'acceso': 'PERMITIDO', 'mensaje': 'Entrada Registrada Correctamente',
            'nombre': nombre_completo, 'hora_entrada': hora_actual_str
        }

    elif mode == 'punch-out':
        if not os.path.exists(nombre_archivo_fichajes):
            return {'acceso': 'DENEGADO', 'mensaje': 'Error: No hay registros de entrada hoy.', 'nombre': nombre_completo}
        df_fichajes_hoy = pd.read_excel(nombre_archivo_fichajes)
        df_fichajes_hoy['DNI'] = df_fichajes_hoy['DNI'].astype(str)
        registro_entrada = df_fichajes_hoy[df_fichajes_hoy['DNI'] == dni]
        if registro_entrada.empty:
            return {'acceso': 'DENEGADO', 'mensaje': 'Error: No se encontró registro de entrada para hoy.', 'nombre': nombre_completo}
        personas_adentro.discard(dni)
        hora_entrada = registro_entrada.iloc[0].get('Hora_Entrada', 'N/A')
        registrar_evento_fichaje(dni, nombre_completo, fecha_hoy_str, hora_entrada, hora_actual_str)
        return {
            'acceso': 'PERMITIDO', 'mensaje': 'Salida Registrada Correctamente',
            'nombre': nombre_completo, 'hora_entrada': hora_entrada, 'hora_salida': hora_actual_str
        }

    return {'acceso': 'DENEGADO', 'mensaje': 'Modo de fichaje no reconocido.', 'nombre': ''}