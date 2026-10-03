# Control de Acceso AOS - Resumen de Actualización V2 (Fase A - D)

Este documento detalla todas las mejoras, módulos y cambios arquitectónicos implementados en la segunda gran versión del sistema de Control de Acceso.

## 1. Migración a Base de Datos (SQLite)
El mayor salto tecnológico fue abandonar la dependencia exclusiva de archivos de Excel (que causaban bloqueos y cuellos de botella) para pasar a un motor de base de datos relacional robusto (`control_acceso.db`).

- **Doble Registro (Dual Logging):** Para mantener compatibilidad hacia atrás y facilitar la auditoría, el sistema ahora guarda cada acceso y fichaje *tanto* en SQLite como en los Excels diarios antiguos.
- **Eficiencia y Concurrencia:** SQLite permite lecturas inmediatas y evita el error de "Permission Denied" cuando alguien tiene el Excel abierto.
- **Scripts de Migración:** Se construyeron scripts automáticos (`migrate_to_sqlite.py`, `migrate_db_rrhh.py`, `migrate_db_alertas.py`) para transferir todas las autorizaciones, nóminas y configuraciones antiguas al nuevo formato.

## 2. Dashboard de Supervisión Local (UI Web)
Se creó un panel de control accesible desde la PC del Supervisor (o cualquier PC de la red local) vía navegador web, sin requerir internet.

- **Faro UDP (Beacon):** El servidor emite una señal a través de la red local para que el "Lanzador del Supervisor" (un acceso directo en la PC) encuentre automáticamente la IP del servidor de la garita. No requiere configuración manual de IP estática.
- **Monitoreo en Vivo (Stats):** Visualización en tiempo real de Entradas, Salidas, Rechazos, Fichajes de RRHH y Personas Adentro.
- **Auditoría Avanzada:** Un motor de búsqueda interno para cruzar fechas y obtener reportes inmediatos, con la opción de exportar directamente a Excel.

## 3. Módulo Avanzado de RRHH y Fichadas (Fase F)
Se reescribió por completo la lógica de cómo el sistema procesa a los vigiladores y personal de limpieza.

- **Motor Matemático Nocturno:** El módulo `hr_calculator.py` procesa los turnos usando comparaciones cruzadas de `datetime` para identificar qué minutos caen dentro del horario nocturno (21:00 a 06:00 hs), resolviendo el clásico problema de "turnos que cruzan la medianoche".
- **Puestos Históricos (Trazabilidad):** Cada vez que un empleado ficha, el sistema "congela" su Puesto Específico actual (ej: "R1"). Si el mes que viene lo cambian a "R3", su historial del mes pasado seguirá diciendo "R1".
- **ABM en el Dashboard:** Una pestaña dedicada ("Nóminas") para dar de alta empleados, asignarles Categoría (ej: "Seguridad Mall", "Limpieza"), Puesto ("Portón 3") y bloquearles el acceso con 1 clic.
- **Sábana RRHH Automática:** La exportación desde la pestaña Auditoría ahora cruza todas las fichadas y separa matemáticamente las "Horas Diurnas" y "Horas Nocturnas" del personal.

## 4. Alertas de Seguridad en Tiempo Real (Fase D)
El sistema ahora reacciona de manera instantánea ante un evento sospechoso.

- **Tecnología SSE (Server-Sent Events):** Se implementó un túnel de conexión permanente (sin polling/sin lag) entre el servidor Flask y el Dashboard. 
- **Cartel Rojo y Alarma:** Si un acceso es denegado o salta una persona marcada, el dashboard de la PC oscurece la pantalla, reproduce una sirena policial corta y dibuja un enorme cartel rojo. El cartel es persistente y obliga al supervisor a hacer clic en "Enterado".
- **Lista Negra (Personas de Interés):** Desde la pestaña de "Configuración", se pueden agregar DNIs bajo sospecha. Aunque el guardia en la garita los deje pasar, el supervisor será notificado por el cartel rojo de forma inmediata.

## 5. Actualizaciones Remotas y Reinicio (Fase E)
Se construyó un sistema "OTA" (Over-The-Air) para inyectar parches a la garita sin tener que levantarse de la silla ni usar AnyDesk.
- **Subida de Parches:** Desde la pestaña de Configuración, el supervisor puede subir un archivo `.py` o `.zip` validado con su contraseña.
- **Reinicio en Caliente:** Al recibir el archivo, la garita aplica el parche y reinicia su propio proceso en 2 segundos, y el dashboard se auto-recarga.
- **Reinicio Manual:** Un botón de pánico en el Dashboard para forzar el reinicio de la aplicación si se congela.

## 6. Backups y Seguridad de Datos (Fase G)
El sistema ahora está protegido contra fallos de disco o borrado accidental de la base de datos `control_acceso.db`.
- **Cronjob Diario:** El sistema se autocomprime en un `.zip` todos los días a las 03:00 AM y se guarda localmente en la carpeta `backups/`. Mantiene las últimas 2 semanas y limpia lo viejo automáticamente.
- **Respaldo Manual:** Un botón de "Descargar Copia de Seguridad" en el Dashboard permite descargar instantáneamente la base de datos completa a la PC del supervisor o a un pendrive.

## 7. Próximos Pasos (Pendientes)
- **Lector de Huellas Digitales:** Esperando adquisición de hardware para armar el integrador USB.
- **BUGFIX - Autenticación API IRSA:** Reparar el proceso de inicio de sesión hacia el portal de locatarios que fue alterado recientemente.
