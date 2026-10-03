import os
from datetime import datetime
from database import get_db_connection
from logger_config import get_logger, log_access_event

logger = get_logger(__name__)

def verificar_dni_sqlite(dni_limpio_str):
    """
    Busca el DNI en la base de datos (Autorizaciones).
    Devuelve un diccionario con los datos del permiso si es válido, o None si no.
    """
    hoy = datetime.now().strftime('%Y-%m-%d')
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # Buscar todos los permisos vigentes para este DNI
        cursor.execute('''
            SELECT * FROM autorizaciones 
            WHERE dni = ? AND activo = 1 
            AND (fecha_inicio IS NULL OR fecha_inicio = '' OR fecha_inicio <= ?)
            AND (fecha_fin IS NULL OR fecha_fin = '' OR fecha_fin >= ?)
        ''', (dni_limpio_str, hoy, hoy))
        
        resultados = cursor.fetchall()
        
        if not resultados:
            return None
            
        row = resultados[0]
        nombre_completo = f"{row['nombre']} {row['apellido']}".strip()
        local = row['local'] or "N/A"
        tarea = row['tarea'] or "N/A"
        vencimiento_f = datetime.strptime(row['fecha_fin'], '%Y-%m-%d').strftime('%d/%m/%Y') if row['fecha_fin'] else "Indefinido"
        tipo_permiso = row['tipo_permiso'] or "N/A"
        num_permiso = row['num_permiso'] or "N/A"
        
        return {
            'nombre': nombre_completo,
            'local': local,
            'tarea': tarea,
            'vence': vencimiento_f,
            'tipo_permiso': tipo_permiso,
            'num_permiso': num_permiso
        }

def loggear_acceso_sqlite(dni, nombre, evento, tipo_permiso, num_permiso, local, tarea, resultado, motivo_rechazo="", puerta="Master", tipo_visita=""):
    """Registra el evento en la tabla registros_accesos de la base de datos"""
    now = datetime.now()
    sql_date = now.strftime('%Y-%m-%d')
    sql_time = now.strftime('%H:%M:%S')
    
    with get_db_connection() as conn:
        conn.execute('''
            INSERT INTO registros_accesos (fecha, hora, dni, nombre, evento, tipo_permiso, local, resultado, motivo_rechazo, puerta, tipo_visita)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (sql_date, sql_time, str(dni), str(nombre), evento, tipo_permiso, local, resultado, motivo_rechazo, puerta, tipo_visita))
        
def guardar_estado_adentro_sqlite(dni, nombre, hora, tipo, local, puerta="Master"):
    with get_db_connection() as conn:
        conn.execute('''
            INSERT OR REPLACE INTO estado_adentro (dni, nombre, hora_ingreso, tipo, local, puerta)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (str(dni), str(nombre), str(hora), str(tipo), str(local), puerta))

def borrar_estado_adentro_sqlite(dni):
    with get_db_connection() as conn:
        conn.execute('DELETE FROM estado_adentro WHERE dni = ?', (str(dni),))

from hr_calculator import calcular_horas_turno

def registrar_fichaje_sqlite(dni, nombre, evento, puerta="Master", puesto_seleccionado=""):
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    hora_actual = datetime.now().strftime('%H:%M:%S')
    dni_str = str(dni).strip()
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # Obtener categoria y puesto_historico desde autorizaciones si no se envía un puesto seleccionado
        cursor.execute("SELECT categoria, puesto_especifico FROM autorizaciones WHERE dni = ? AND activo = 1 ORDER BY id DESC LIMIT 1", (dni_str,))
        row_auth = cursor.fetchone()
        categoria = row_auth['categoria'] if row_auth else 'N/A'
        puesto = puesto_seleccionado if puesto_seleccionado else (row_auth['puesto_especifico'] if row_auth else '')

        if evento == "ENTRADA":
            cursor.execute('''
                INSERT INTO registros_fichajes (fecha, dni, nombre, hora_entrada, estado, categoria, puesto_historico, puerta)
                VALUES (?, ?, ?, ?, 'EN CURSO', ?, ?, ?)
            ''', (fecha_hoy, dni_str, str(nombre), hora_actual, categoria, puesto, puerta))
        elif evento == "SALIDA":
            cursor.execute('''
                SELECT id, hora_entrada, categoria FROM registros_fichajes 
                WHERE dni = ? AND fecha = ? AND estado = 'EN CURSO' 
                ORDER BY id DESC LIMIT 1
            ''', (dni_str, fecha_hoy))
            row = cursor.fetchone()
            if row:
                hora_in = row['hora_entrada']
                cat = row['categoria']
                # Si la categoría contiene "Limpieza" (ignorando mayúsculas), no aplica nocturnidad.
                aplica_nocturnidad = False if 'limpieza' in str(cat).lower() else True
                
                h_diurnas, h_nocturnas = calcular_horas_turno(hora_in, hora_actual, aplica_nocturnidad)
                
                cursor.execute('''
                    UPDATE registros_fichajes 
                    SET hora_salida = ?, estado = 'COMPLETADO', horas_diurnas = ?, horas_nocturnas = ?, puerta = ?
                    WHERE id = ?
                ''', (hora_actual, h_diurnas, h_nocturnas, puerta, row['id']))
            else:
                cursor.execute('''
                    INSERT INTO registros_fichajes (fecha, dni, nombre, hora_salida, estado, categoria, puesto_historico, puerta)
                    VALUES (?, ?, ?, ?, 'HUERFANO_SALIDA', ?, ?, ?)
                ''', (fecha_hoy, dni_str, str(nombre), hora_actual, categoria, puesto, puerta))

# --- GESTION DE PUESTOS OPERATIVOS ---

def obtener_puestos_operativos():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, nombre, activo FROM puestos_operativos ORDER BY nombre")
        return [dict(row) for row in cursor.fetchall()]

def agregar_puesto_operativo(nombre):
    with get_db_connection() as conn:
        try:
            conn.execute("INSERT INTO puestos_operativos (nombre) VALUES (?)", (nombre.strip(),))
            return True, "Puesto agregado correctamente."
        except sqlite3.IntegrityError:
            return False, "El puesto ya existe."

def eliminar_puesto_operativo(puesto_id):
    with get_db_connection() as conn:
        conn.execute("DELETE FROM puestos_operativos WHERE id = ?", (puesto_id,))
        return True, "Puesto eliminado."


# --- GESTION DE PUESTOS FISICOS ---
def obtener_puestos_fisicos():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, nombre, sector, hora_inicio, hora_fin, activo FROM puestos_fisicos ORDER BY sector, nombre")
        return [dict(row) for row in cursor.fetchall()]

def agregar_puesto_fisico(nombre, sector, hora_inicio, hora_fin):
    with get_db_connection() as conn:
        try:
            conn.execute("INSERT INTO puestos_fisicos (nombre, sector, hora_inicio, hora_fin) VALUES (?, ?, ?, ?)", 
                         (nombre.strip(), sector, hora_inicio, hora_fin))
            return True, "Puesto agregado correctamente."
        except sqlite3.IntegrityError:
            return False, "El puesto ya existe."

def eliminar_puesto_fisico(puesto_id):
    with get_db_connection() as conn:
        conn.execute("DELETE FROM puestos_fisicos WHERE id = ?", (puesto_id,))
        return True, "Puesto eliminado."

# --- GESTION DE EMPRESAS ---
def obtener_empresas():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, nombre, activo FROM empresas_contratistas ORDER BY nombre")
        return [dict(row) for row in cursor.fetchall()]

def agregar_empresa(nombre):
    with get_db_connection() as conn:
        try:
            conn.execute("INSERT INTO empresas_contratistas (nombre) VALUES (?)", (nombre.strip(),))
            return True, "Empresa agregada correctamente."
        except sqlite3.IntegrityError:
            return False, "La empresa ya existe."

def eliminar_empresa(empresa_id):
    with get_db_connection() as conn:
        conn.execute("DELETE FROM empresas_contratistas WHERE id = ?", (empresa_id,))
        return True, "Empresa eliminada."


# --- GESTION DE NOMINA (ABM PERSONAL) ---
def guardar_empleado_nomina(data):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        emp_id = data.get('id')
        dni = data.get('dni')
        nombre = data.get('nombre')
        categoria = data.get('categoria')
        activo = 1 if data.get('activo') else 0
        huella = data.get('huella_template')
        foto_b64 = data.get('foto_b64')
        id_empresa = data.get('id_empresa') or None
        tipo_empleado = data.get('tipo_empleado') or None

        if emp_id:
            # Update
            cursor.execute('''
                UPDATE autorizaciones 
                SET dni=?, nombre=?, tarea=?, activo=?, huella_template=?, foto_path=?, id_empresa=?, tipo_empleado=?
                WHERE id=? AND tipo_permiso='NOMINA'
            ''', (dni, nombre, categoria, activo, huella, foto_b64, id_empresa, tipo_empleado, emp_id))
        else:
            # Insert
            cursor.execute('''
                INSERT INTO autorizaciones (dni, nombre, tarea, tipo_permiso, activo, huella_template, foto_path, id_empresa, tipo_empleado)
                VALUES (?, ?, ?, 'NOMINA', ?, ?, ?, ?, ?)
            ''', (dni, nombre, categoria, activo, huella, foto_b64, id_empresa, tipo_empleado))
            
        return True, "Empleado guardado correctamente."

def eliminar_empleado_nomina(emp_id):
    with get_db_connection() as conn:
        conn.execute("DELETE FROM autorizaciones WHERE id=? AND tipo_permiso='NOMINA'", (emp_id,))
        return True, "Empleado eliminado."
