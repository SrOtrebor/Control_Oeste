"""
Documentación de Funciones Principales - Control de Acceso AOS

Este documento contiene la documentación completa de las funciones principales
del sistema, incluyendo descripciones, parámetros, retornos y ejemplos de uso.
"""

# ============================================================================
# ACCESS_MANAGER.PY
# ============================================================================

## parsear_codigo_barra(scanner_data)
"""
Parsea datos del escáner de código de barras de DNI.

Maneja múltiples formatos de DNI según la edición del documento:
- Formato 1: DNI nuevo con comillas (8+ campos separados por ")
- Formato 2: Formato con @ y guiones bajos (@apellido_nombre_sexo_dni_)
- Formato 3: Búsqueda de DNI de 7-8 dígitos en el texto
- Formato 4: Extracción de todos los dígitos

Args:
    scanner_data (str): Datos crudos del escáner de código de barras

Returns:
    dict or None: Diccionario con datos parseados o None si falla
        - dni (str): DNI limpio (7-8 dígitos)
        - nombre (str): Nombre (si está disponible)
        - apellido (str): Apellido (si está disponible)
        - nombre_completo (str): Nombre completo (si está disponible)
        - sexo (str): Sexo (opcional)
        - fecha_nacimiento (str): Fecha de nacimiento (opcional)
        - fecha_vencimiento (str): Fecha de vencimiento del DNI (opcional)

Ejemplo:
    >>> parsear_codigo_barra('"PEREZ"JUAN"M"12345678"...')
    {'dni': '12345678', 'nombre': 'JUAN', 'apellido': 'PEREZ', ...}
"""

## registrar_evento(dni, nombre, hora_evento, evento, tipo_permiso, num_permiso, local, tarea, resultado)
"""
Registra un evento de acceso en el archivo Excel diario.

Crea automáticamente el archivo si no existe y lo formatea con colores.
Todos los eventos se registran con timestamp para auditoría.

Args:
    dni (str): DNI de la persona
    nombre (str): Nombre completo
    hora_evento (str): Hora del evento en formato HH:MM:SS
    evento (str): Tipo de evento ('Entrada OK', 'Salida', 'Entrada RECHAZADA', etc.)
    tipo_permiso (str): Tipo de permiso (FAP, FAO, Nómina Persistente, Excepción, VISITA)
    num_permiso (str): Número de permiso o quien autoriza
    local (str): Local asociado al permiso
    tarea (str): Tarea asociada (para FAO)
    resultado (str): Resultado del evento ('AUTORIZADO', 'DENEGADO', 'REGISTRADO')

Efectos secundarios:
    - Crea archivo Excel en registros_diarios/ si no existe
    - Aplica formato con colores (verde para autorizados, rojo para denegados)
    - Registra en logs

Ejemplo:
    >>> registrar_evento('12345678', 'Juan Perez', '09:30:00', 'Entrada OK', 
                         'FAP', 'FAP-001', 'Local A', 'N/A', 'AUTORIZADO')
"""

## registrar_evento_fichaje(dni, nombre, fecha, hora_entrada, hora_salida)
"""
Registra un evento de fichaje (entrada/salida) de empleados.

Crea un registro consolidado diario con hora de entrada y salida.
Si ya existe un registro para el DNI en el día, actualiza la hora de salida.

Args:
    dni (str): DNI del empleado
    nombre (str): Nombre completo del empleado
    fecha (str): Fecha en formato YYYY-MM-DD
    hora_entrada (str): Hora de entrada en formato HH:MM:SS (o None)
    hora_salida (str): Hora de salida en formato HH:MM:SS (o None)

Returns:
    bool: True si se registró correctamente, False si hubo error

Efectos secundarios:
    - Crea archivo Excel en registros_fichajes/ si no existe
    - Actualiza registro existente si el empleado ya fichó entrada

Ejemplo:
    >>> registrar_evento_fichaje('12345678', 'Juan Perez', '2026-01-12', '09:00:00', None)
    True
"""

## verificar_dni(scanner_data, mode)
"""
Función principal de verificación de acceso.

Verifica el DNI contra todas las listas autorizadas (Nóminas, FAP, FAO, Excepciones)
y determina si se permite o deniega el acceso. Maneja tres modos de operación.

Args:
    scanner_data (str): Datos crudos del escáner de DNI
    mode (str): Modo de operación
        - 'entrada': Verificación de acceso normal
        - 'salida': Registro de salida de visitante
        - 'visita': Registro de entrada de visitante

Returns:
    dict: Resultado de la verificación con las siguientes claves:
        - acceso (str): 'PERMITIDO' o 'DENEGADO'
        - nombre (str): Nombre de la persona
        - mensaje (str): Mensaje descriptivo del resultado
        - tipo_permiso (str): Tipo de permiso encontrado (opcional)
        - num_permiso (str): Número de permiso (opcional)
        - local (str): Local asociado (opcional)
        - tarea (str): Tarea asociada (opcional)
        - vence (str): Fecha de vencimiento (opcional)

Lógica de verificación (modo 'entrada'):
    1. Parsea el DNI del scanner
    2. Verifica en Nóminas Persistentes (con vigencia)
    3. Verifica en lista FAP (con vigencia)
    4. Verifica en lista FAO (con vigencia)
    5. Verifica en Excepciones (con vigencia)
    6. Si no encuentra en ninguna, DENIEGA

Efectos secundarios:
    - Registra evento en Excel diario
    - Actualiza set personas_adentro
    - Registra en logs (access_events.log)

Ejemplo:
    >>> verificar_dni('"PEREZ"JUAN"M"12345678"...', 'entrada')
    {'acceso': 'PERMITIDO', 'nombre': 'JUAN PEREZ', 'mensaje': 'ACCESO PERMITIDO (FAP): JUAN PEREZ', ...}
"""

## registrar_fichaje(dni, nombre, modo)
"""
Registra fichaje de entrada o salida de empleados.

Wrapper de alto nivel para registrar_evento_fichaje que maneja la lógica
de entrada/salida y actualiza el set de personas_adentro.

Args:
    dni (str): DNI del empleado
    nombre (str): Nombre completo
    modo (str): 'entrada' o 'salida'

Returns:
    dict: Resultado del fichaje
        - success (bool): True si se registró correctamente
        - mensaje (str): Mensaje descriptivo

Ejemplo:
    >>> registrar_fichaje('12345678', 'Juan Perez', 'entrada')
    {'success': True, 'mensaje': 'Entrada registrada para Juan Perez'}
"""

# ============================================================================
# DATA_MANAGER.PY
# ============================================================================

## cargar_autorizaciones()
"""
Carga y procesa todos los archivos de autorización (FAP, FAO, Excepciones, Nóminas).

Implementa un sistema de caché inteligente que solo recarga archivos si fueron modificados.
Normaliza los datos y estandariza las columnas para facilitar las búsquedas.

Efectos secundarios:
    - Actualiza variables globales: df_fap, df_fao, df_excepciones, df_nominas
    - Actualiza timestamps de última modificación
    - Registra en logs las operaciones realizadas

Archivos procesados:
    - ListadoFAPs.xlsx (header en fila 1)
    - ListadoFAOs.xlsx (header en fila 1)
    - excepciones.xlsx
    - nominas_persistentes.xlsx

Transformaciones aplicadas:
    - Normalización de DNIs (elimina puntos, espacios, guiones)
    - Renombrado de columnas a nombres estándar
    - Concatenación de nombre y apellido
    - Conversión de fechas a formato datetime
"""

## formatear_excel(nombre_archivo)
"""
Aplica formato visual a archivos Excel (colores, anchos de columna, alineación).

Args:
    nombre_archivo (str): Ruta completa al archivo Excel

Formato aplicado:
    - Header: Fondo azul, texto blanco, negrita
    - Filas alternas: Fondo gris claro
    - Ajuste automático de ancho de columnas
    - Alineación centrada

Efectos secundarios:
    - Modifica el archivo Excel en disco
    - Registra en logs si hay errores
"""

## procesar_nomina_texto(texto_nomina, empresa, vigencia_desde, vigencia_hasta)
"""
Procesa texto pegado de nómina y extrae DNIs y nombres.

Maneja múltiples formatos de entrada:
    - CUIL + Nombre completo
    - DNI + Nombre completo
    - Diferentes separadores (espacios, tabs, comas)

Args:
    texto_nomina (str): Texto con la nómina (múltiples líneas)
    empresa (str): Nombre de la empresa
    vigencia_desde (str): Fecha inicio de vigencia (YYYY-MM-DD)
    vigencia_hasta (str): Fecha fin de vigencia (YYYY-MM-DD)

Returns:
    list: Lista de diccionarios con personas procesadas
        Cada diccionario contiene: DNI, Nombre, Apellido, Local (empresa), Vigencia

Lógica de procesamiento:
    1. Divide texto en líneas
    2. Ignora líneas vacías y encabezados
    3. Extrae CUIL/DNI usando regex
    4. Extrae nombre completo
    5. Separa nombre y apellido
    6. Valida datos antes de agregar

Ejemplo:
    >>> procesar_nomina_texto("20-12345678-9 PEREZ JUAN\\n...", "Empresa A", "2026-01-01", "2026-12-31")
    [{'DNI': '12345678', 'Nombre': 'JUAN', 'Apellido': 'PEREZ', ...}]
"""

## generar_reporte_consolidado(fecha)
"""
Genera un reporte consolidado de accesos del día.

Combina registros de entrada y salida en una sola fila por persona.
Calcula tiempo de permanencia si hay entrada y salida.

Args:
    fecha (str): Fecha del reporte en formato YYYY-MM-DD

Returns:
    pd.DataFrame: DataFrame consolidado con columnas:
        - DNI
        - Nombre
        - Hora Ingreso
        - Hora Salida
        - Evento
        - Permiso/Autoriza
        - Local
        - Tarea
        - Resultado

Lógica de consolidación:
    - Agrupa por DNI
    - Primera entrada = Hora Ingreso
    - Última salida = Hora Salida
    - Combina información de múltiples eventos
"""

# ============================================================================
# UTILS.PY
# ============================================================================

## validar_archivo_fap(filepath)
"""
Valida que un archivo FAP tenga el formato esperado.

Verificaciones realizadas:
    - Existencia de columnas requeridas
    - Archivo no vacío
    - DNIs en formato válido (7-8 dígitos)

Args:
    filepath (str): Ruta al archivo Excel FAP

Returns:
    tuple: (es_válido, mensaje)
        - es_válido (bool): True si el archivo es válido
        - mensaje (str): Descripción del resultado o error

Columnas requeridas:
    - Numero (DNI)
    - Nombre
    - Apellido
    - FAP (número de permiso)
    - Fecha Fin (vencimiento)
    - Marca (local)

Ejemplo:
    >>> validar_archivo_fap('ListadoFAPs.xlsx')
    (True, 'Archivo válido: 150 registros')
"""

## validar_archivo_fao(filepath)
"""
Valida que un archivo FAO tenga el formato esperado.

Similar a validar_archivo_fap pero incluye validación de columna 'Tarea'.

Args:
    filepath (str): Ruta al archivo Excel FAO

Returns:
    tuple: (es_válido, mensaje)

Columnas requeridas:
    - Numero, Nombre, Apellido, FAO, Fecha Fin, Marca, Tarea
"""

## crear_backup(filepath, tipo='manual')
"""
Crea un backup de un archivo con timestamp.

Args:
    filepath (str): Ruta al archivo a respaldar
    tipo (str): Tipo de backup ('manual', 'auto', 'pre-update')

Returns:
    tuple: (éxito, ruta_backup_o_error)
        - éxito (bool): True si se creó correctamente
        - ruta_backup_o_error (str): Ruta del backup o mensaje de error

Estructura de backups:
    backups/
    ├── manual/
    ├── auto/
    └── pre-update/

Formato de nombre:
    {nombre_original}_{YYYYMMDD_HHMMSS}.xlsx

Ejemplo:
    >>> crear_backup('ListadoFAPs.xlsx', 'pre-update')
    (True, 'backups/pre-update/ListadoFAPs_20260112_203000.xlsx')
"""

## limpiar_backups_antiguos(dias=30, tipo='auto')
"""
Elimina backups más antiguos que X días.

Args:
    dias (int): Días de antigüedad para eliminar
    tipo (str): Tipo de backup a limpiar

Returns:
    int: Cantidad de archivos eliminados

Ejemplo:
    >>> limpiar_backups_antiguos(30, 'auto')
    5  # Eliminó 5 archivos
"""

# ============================================================================
# LOGGER_CONFIG.PY
# ============================================================================

## get_logger(name, level=logging.INFO)
"""
Obtiene un logger configurado para un módulo específico.

Configura automáticamente:
    - Archivo de log por módulo
    - Log general de aplicación
    - Log de errores
    - Salida a consola (WARNING+)
    - Formato consistente con timestamps

Args:
    name (str): Nombre del módulo (usar __name__)
    level (int): Nivel de logging (default: INFO)

Returns:
    logging.Logger: Logger configurado

Ejemplo:
    >>> logger = get_logger(__name__)
    >>> logger.info("Operación exitosa")
    >>> logger.error("Error crítico", exc_info=True)
"""

## log_access_event(dni, nombre, resultado, tipo_permiso='N/A', mensaje='')
"""
Registra un evento de acceso de forma estructurada.

Crea un log especial en access_events.log para auditoría.

Args:
    dni (str): DNI de la persona
    nombre (str): Nombre completo
    resultado (str): 'PERMITIDO' o 'DENEGADO'
    tipo_permiso (str): Tipo de permiso
    mensaje (str): Mensaje adicional

Ejemplo:
    >>> log_access_event('12345678', 'Juan Perez', 'PERMITIDO', 'FAP', 'Local: A')
"""
