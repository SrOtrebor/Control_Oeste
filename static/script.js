document.addEventListener('DOMContentLoaded', () => {

    // --- FUNCIONES GLOBALES (Accesibles desde todas las páginas) ---

    async function loadDailyRecords(targetTableBody) {
        if (!targetTableBody) return;
        try {
            const response = await fetch('/get_daily_records');
            const data = await response.json();
            targetTableBody.innerHTML = '';
            if (data.success && data.records.length > 0) {
                data.records.reverse();
                data.records.forEach(record => {
                    const row = targetTableBody.insertRow();
                    if (record.Resultado === 'VERDE') row.classList.add('fila-verde');
                    else if (record.Resultado === 'ROJO') row.classList.add('fila-roja');
                    if (record.Tipo_Permiso === 'FAO') row.classList.add('fila-fao');

                    row.insertCell().textContent = record.Hora_Ingreso ? record.Hora_Ingreso.split(' ')[1] || record.Hora_Ingreso : '';
                    row.insertCell().textContent = record.DNI || '';
                    row.insertCell().textContent = record['Nombre y Apellido'] || '';
                    row.insertCell().textContent = record.Num_Permiso || '';
                    row.insertCell().textContent = record.Local || '';
                    row.insertCell().textContent = record.Tarea || '';
                });
            } else {
                const row = targetTableBody.insertRow();
                const cell = row.insertCell(0);
                cell.colSpan = 6;
                cell.textContent = data.message || 'No hay registros para mostrar.';
                cell.style.textAlign = 'center';
            }
        } catch (error) {
            console.error('Error al cargar registros diarios:', error);
        }
    }

    async function updateStats() {
        const statTotalEl = document.getElementById('statTotal');
        if (!statTotalEl) return;
        try {
            const response = await fetch('/get_daily_stats');
            const stats = await response.json();
            statTotalEl.textContent = stats.total;
            document.getElementById('statPermitidos').textContent = stats.permitidos;
            document.getElementById('statRechazados').textContent = stats.rechazados;
        } catch (error) {
            console.error("Error al actualizar estadísticas:", error);
        }
    }


    // --- LÓGICA ESPECÍFICA PARA CADA PÁGINA ---

    // Lógica para la Página de Admisión (index.html)
    const dniInput = document.getElementById('dniInput');
    if (dniInput) {
        const resultadoDiv = document.getElementById('resultado');
        const mensajeResultado = document.getElementById('mensajeResultado');
        const nombrePersona = document.getElementById('nombrePersona');
        const infoLocal = document.getElementById('infoLocal');
        const infoTarea = document.getElementById('infoTarea');
        const infoVence = document.getElementById('infoVence');
        const ingresosListBody = document.getElementById('ingresosListBody');

        dniInput.focus();
        dniInput.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                verificarDNI();
            }
        });

       async function verificarDNI() {
    const dniInput = document.getElementById('dniInput');
    const scannerData = dniInput.value.trim();
    if (!scannerData) return;

    let dniParaVerificar = '';

    // --- LÓGICA DE PARSEO UNIVERSAL ---
    // Busca la primera secuencia de exactamente 7 u 8 dígitos en toda la cadena.
    const match = scannerData.match(/\b\d{7,8}\b/);
    
    if (match) {
        // Si encuentra una, ese es nuestro DNI.
        dniParaVerificar = match[0];
        console.log("DNI extraído con método universal:", dniParaVerificar);
    } else {
        // Si no encuentra un número de 7 u 8 dígitos, puede que sea un ingreso manual
        // o un formato muy raro. Tomamos todos los números juntos.
        const numeros = scannerData.replace(/\D/g, '');
        if (numeros.length >= 7) {
            dniParaVerificar = numeros;
        }
    }
    // --- FIN DE LA LÓGICA ---

    if (!dniParaVerificar) {
        mostrarResultado('red', 'No se pudo encontrar un DNI válido en el código.', '', '', '', '');
        dniInput.value = '';
        dniInput.focus();
        return;
    }

    try {
        const response = await fetch('/verificar_dni', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ dni: dniParaVerificar }) });
        const data = await response.json();
        
        const vencimientoTexto = data.vence && data.vence !== 'N/A' ? (data.acceso === 'PERMITIDO' ? `Vence: ${data.vence}`: `Vencido el: ${data.vence}`) : '';
        mostrarResultado(data.acceso === 'PERMITIDO' ? 'green' : 'red', data.mensaje, data.nombre, data.local, data.tarea, vencimientoTexto);

    } catch (error) {
        console.error('Error al verificar DNI:', error);
        mostrarResultado('red', 'Error de comunicación.', '', '', '', '');
    } finally {
        dniInput.value = '';
        dniInput.focus();
        if (typeof updateStats === 'function') updateStats();
        if (typeof loadDailyRecords === 'function') {
            const ingresosListBody = document.getElementById('ingresosListBody');
            const adminIngresosListBody = document.getElementById('adminIngresosListBody');
            if (ingresosListBody) loadDailyRecords(ingresosListBody);
            if (adminIngresosListBody) loadDailyRecords(adminIngresosListBody);
        }
    }
}

        function mostrarResultado(colorClass, mensaje, nombre, localInfo, tareaInfo, venceInfo) {
            resultadoDiv.className = `result-display luz-${colorClass}`;
            mensajeResultado.textContent = mensaje;
            nombrePersona.textContent = nombre && nombre !== 'No Encontrado' ? `Nombre: ${nombre}` : '';
            infoLocal.textContent = '';
            infoTarea.textContent = '';
            infoVence.textContent = '';
            if (colorClass === 'green') {
                infoLocal.textContent = localInfo && localInfo !== 'N/A' ? `Local: ${localInfo}` : '';
                infoTarea.textContent = tareaInfo && tareaInfo !== 'N/A' ? `Tarea: ${tareaInfo}` : '';
                infoVence.textContent = venceInfo;
            }
        }
        
        loadDailyRecords(ingresosListBody);
        updateStats();
    }

    // Lógica para la Página de Login
    const loginBtn = document.getElementById('loginBtn');
    if (loginBtn) {
        const usernameInput = document.getElementById('username');
        const passwordInput = document.getElementById('password');
        const loginMessage = document.getElementById('loginMessage');
        loginBtn.addEventListener('click', async () => {
            const username = usernameInput.value.trim();
            const password = passwordInput.value.trim();
            try {
                const response = await fetch('/perform_login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username, password }) });
                const data = await response.json();
                if (data.success) {
                    window.location.href = '/admin';
                } else {
                    loginMessage.textContent = data.message;
                }
            } catch (error) {
                loginMessage.textContent = 'Error de comunicación.';
            }
        });
    }

    // Lógica para la Página de Administración
    const adminPageIdentifier = document.getElementById('adminIngresosListBody');
    if (adminPageIdentifier) {
        loadDailyRecords(adminPageIdentifier);
        
        const fapInput = document.getElementById('fapFileInput');
        const faoInput = document.getElementById('faoFileInput');
        const fapFileName = document.getElementById('fapFileName');
        const faoFileName = document.getElementById('faoFileName');

        if (fapInput) {
            fapInput.addEventListener('change', () => {
                fapFileName.textContent = fapInput.files.length > 0 ? fapInput.files[0].name : 'Ningún archivo seleccionado';
            });
        }
        if (faoInput) {
            faoInput.addEventListener('change', () => {
                faoFileName.textContent = faoInput.files.length > 0 ? faoInput.files[0].name : 'Ningún archivo seleccionado';
            });
        }

        const uploadExcelBtn = document.getElementById('uploadExcelBtn');
        uploadExcelBtn.addEventListener('click', async () => {
            const uploadStatus = document.getElementById('uploadStatus');
            const formData = new FormData();
            if (fapInput.files.length > 0) formData.append('fap_file', fapInput.files[0]);
            if (faoInput.files.length > 0) formData.append('fao_file', faoInput.files[0]);
            if (!formData.entries().next().value) {
                uploadStatus.textContent = 'Por favor, seleccione al menos un archivo.'; return;
            }
            uploadStatus.textContent = 'Cargando...';
            try {
                const response = await fetch('/upload_excel', { method: 'POST', body: formData });
                const data = await response.json();
                uploadStatus.textContent = data.message;
            } catch (error) {
                uploadStatus.textContent = 'Error de comunicación.';
            }
        });

        const addExcepcionBtn = document.getElementById('addExcepcionBtn');
        addExcepcionBtn.addEventListener('click', async () => {
            const excepcionForm = document.querySelector('.excepcion-form');
            const excepcionData = {
                nombre: document.getElementById('nombreExcepcion').value.trim(),
                apellido: document.getElementById('apellidoExcepcion').value.trim(),
                dni: document.getElementById('dniExcepcion').value.trim(),
                local: document.getElementById('localExcepcion').value.trim(),
                quien_autoriza: document.getElementById('autorizaExcepcion').value.trim()
            };
            const excepcionStatus = document.getElementById('excepcionStatus');
            excepcionStatus.textContent = 'Guardando...';
            try {
                const response = await fetch('/add_excepcion', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(excepcionData) });
                const data = await response.json();
                excepcionStatus.textContent = data.message;
                if(data.success && excepcionForm.parentElement) {
                    excepcionForm.closest('form').reset();
                }
            } catch (error) {
                excepcionStatus.textContent = 'Error al procesar la respuesta.';
            }
        });
        
        const enviarReporteBtn = document.getElementById('enviarReporteBtn');
        enviarReporteBtn.addEventListener('click', async () => {
            const reporteStatus = document.getElementById('reporteStatus');
            reporteStatus.textContent = 'Enviando...';
            try {
                const response = await fetch('/enviar_reporte_diario', { method: 'POST' });
                const data = await response.json();
                reporteStatus.textContent = data.message;
            } catch (error) {
                reporteStatus.textContent = 'Error de comunicación.';
            }
        });
    }
});