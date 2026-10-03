# Documentación de Migración a Firebase (Estado Actual)

## 1. Arquitectura Híbrida Implementada
El sistema ha migrado de una API monolítica local a una **Arquitectura Híbrida Orientada a Eventos (IoT)** usando Firebase:

*   **Frontend (Dashboard Web):** 
    *   Alojado en Firebase Hosting (o localmente vía `firebase serve`).
    *   Conectado directamente a Firestore en la nube para actualizaciones en tiempo real (ej: `onSnapshot` para nóminas y alertas).
    *   **Autenticación pendiente:** Actualmente las rutas están expuestas; el próximo paso crítico es proteger `scan.html` y `dashboard.html` usando Firebase Auth.
*   **Backend (Obrero Local en Python):**
    *   El script `firebase_worker.py` corre localmente como un servicio de fondo en la garita.
    *   Escucha una colección especial llamada `system_commands` en Firestore.
    *   Ejecuta tareas pesadas locales (ej: extraer nóminas de IRSA con `irsa_client.py` usando `requests` y `pandas`) y reporta el progreso a la nube.

## 2. Trabajo Realizado Hoy
*   **Limpieza de UI:** Se eliminaron botones obsoletos en la pestaña "Nóminas" (Editar/Borrar) para cumplir con el requisito de que sea un listado de lectura estricto.
*   **Obrero IoT:** Se creó e inicializó `firebase_worker.py` conectado con `serviceAccountKey.json`.
*   **Gestión de Credenciales de IRSA:**
    *   Las credenciales ahora se configuran en el Dashboard y se guardan directamente encriptadas en Firestore (`config/irsa`).
    *   El obrero descarga estas credenciales frescas de Firestore antes de cada sincronización y las encripta localmente (usando Fernet en `irsa_config.py`).
*   **Solución a Tiempos de Espera (Timeouts):**
    *   Se aumentó el límite de espera del portal IRSA de 20s a 45s.
    *   Se forzó el uso de rutas absolutas (`BASE_DIR`) en el worker para prevenir errores de caché vacía.
*   **Carga de Trámites:**
    *   Se implementó la lógica en el worker para volcar el listado general de FAOs y FAPs a la colección `tramites_irsa` en Firestore, además de registrar a cada persona individualmente en `nominas`.

## 3. Problemas Conocidos (A resolver en la próxima sesión)
*   **Nros. de Trámite y Vencimientos no visibles:** 
    *   A pesar de que la lógica de extracción y subida a Firestore (`nro_tramite` y `fecha_vencimiento`) fue programada en `firebase_worker.py` y `dashboard.js`, los datos no se están renderizando correctamente en la tabla de "Nóminas".
    *   *Sospecha:* Es probable que el lote subido no contenga los campos debido a que el usuario no volvió a ejecutar el botón de "Extraer Nóminas IRSA Ahora" después de la última corrección del código, o bien los nombres de las variables en JS (`r.nro_tramite` / `r.fecha_vencimiento`) tienen alguna discordancia con la base local.
*   **Pestaña FAP/FAO Vacía ("No se encontraron trámites"):**
    *   El dashboard fue actualizado para leer directamente desde Firestore (`collection(db, 'tramites_irsa')`), pero aparentemente la colección no se está poblando correctamente, o la sincronización anterior se realizó antes de que se agregara la instrucción `tramites_ref.document().set(...)`.

## 4. Próximos Pasos (To-Do)
1.  **Re-ejecutar Sincronización:** Ejecutar manualmente la sincronización para verificar la correcta ingesta de `nro_tramite`, `fecha_vencimiento` y los documentos completos en `tramites_irsa`.
2.  **Verificación de Interfaz:** Corregir cualquier desajuste visual o de mapeo de campos en las tablas "Nóminas" y "FAP/FAO".
3.  **Sistema OTA:** Probar el sistema de actualización remota (`UPDATE_OTA`) de la garita.
4.  **Autenticación de Seguridad (Auth):** Implementar login obligatorio para acceder al dashboard y proteger los datos corporativos sensibles.


### 30 SEP 2026 - Estrategia H�brida Offline-First para Excepciones/ART

Se actualiz� la forma en la que el Dashboard procesa y guarda las listas de Excepciones y n�minas ART para ajustarse a los requerimientos Offline-First y a la realidad de m�ltiples FAOs por persona.

1. **Guardado Directo en Nube (Offline-First):**
El Dashboard (dashboard.js) ya no delega la grabaci�n en la nube al Obrero Local. En su lugar, el navegador escribe *directamente* en la colecci�n 
ominas de Firestore. Esto permite que el sistema web funcione a la perfecci�n, guardando de forma local (v�a IndexedDB) y sincronizando silenciosamente sin dejar la interfaz cargando ('Esperando al obrero...').

2. **Evitar Borrado de Otras Empresas:**
Se corrigi� la query de limpieza. Anteriormente, si una persona ten�a cargado un permiso para 'Empresa A' y se sub�a uno para 'Empresa B', el sistema borraba el de A. Ahora la eliminaci�n de solapamientos se filtra espec�ficamente por \empresa\ para permitir que una misma persona trabaje en dos locales distintos.

3. **Sincronizaci�n H�brida al Hardware:**
Aunque la web guarda directo en \
ominas\, los molinetes dependen del SQLite f�sico. Por lo tanto, adem�s de escribir en \
ominas\, la web env�a silenciosamente un comando \SAVE_EXCEPCIONES\ a \system_commands\. El \irebase_worker.py\ local lo recibe en background y replica los cambios a la tabla \utorizaciones\ de SQLite para que el molinete conceda los accesos.

4. **Regex de Fallback de Texto:**
Se agreg� un regex global en el parseador de PDF de la ART para que, si el usuario pega un bloque de texto chato sin saltos de l�nea, el sistema igual pueda identificar y separar a las personas bas�ndose en el r�gimen (Ej: 'R�gimen General').
