document.addEventListener('DOMContentLoaded', () => {

    // --- ESTADO GLOBAL ---
    let currentMode = 'entrada'; // 'entrada', 'visita', 'salida'
    let currentPunchMode = null; // 'punch-in', 'punch-out'

    // --- FUNCIONES GLOBALES ---

    async function registrarFichaje() {
        const dniInput = document.getElementById('dniInput');
        const scannerData = dniInput.value.trim();
        if (!scannerData) return;

        try {
            const response = await fetch('/registrar_fichaje', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ scanner_data: scannerData, mode: currentPunchMode })
            });
            const data = await response.json();
            
            // Usar la misma función de resultado para mostrar el estado del fichaje
            mostrarResultado(data.acceso === 'PERMITIDO' ? 'gris' : 'red', data.mensaje, data.nombre, `Entrada: ${data.hora_entrada || ''}`, `Salida: ${data.hora_salida || ''}`, '');

        } catch (error) {
            console.error('Error al registrar fichaje:', error);
            mostrarResultado('red', 'Error de red', '', '', '', '');
        } finally {
            dniInput.value = '';
            dniInput.focus();
            // No es necesario recargar estadísticas o registros aquí, a menos que se quiera unificar la vista
        }
    }

    async function loadDailyRecords(targetTableBody) {
        if (!targetTableBody) return;
        try {
            const response = await fetch('/get_daily_records');
            const data = await response.json();
            targetTableBody.innerHTML = '';
            if (data.success && data.records.length > 0) {
                data.records.reverse(); // Muestra los más recientes primero
                data.records.forEach(record => {
                    const row = targetTableBody.insertRow();
                    if (record.Resultado === 'VERDE') row.classList.add('fila-verde');
                    else if (record.Resultado === 'ROJO') row.classList.add('fila-roja');
                    if (record.Tipo_Permiso === 'FAO') row.classList.add('fila-fao');

                    row.insertCell().textContent = record.Hora_Ingreso ? record.Hora_Ingreso.split(' ')[1] || record.Hora_Ingreso : '';
                    row.insertCell().textContent = record.DNI || '';
                    row.insertCell().textContent = record['Nombre y Apellido'] || '';
                    row.insertCell().textContent = record.Evento || '';
                    row.insertCell().textContent = record.Num_Permiso || '';
                    row.insertCell().textContent = record.Local || '';
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
        const statTotalAdentroEl = document.getElementById('statTotalAdentro');
        if (!statTotalAdentroEl) return;
        try {
            const response = await fetch('/get_dynamic_stats');
            const stats = await response.json();
            statTotalAdentroEl.textContent = stats.total_adentro;
            document.getElementById('statPermitidos').textContent = stats.permitidos;
            document.getElementById('statRechazados').textContent = stats.rechazados;
        } catch (error) {
            console.error("Error al actualizar estadísticas:", error);
        }
    }

    function setMode(newMode) {
        currentMode = newMode;
        currentPunchMode = null; // Desactiva el modo fichador
        const dniInput = document.getElementById('dniInput');
        const headerTitle = document.getElementById('header-title');

        // Gestionar clases activas en botones de control
        ['btnEntrada', 'btnRegVisita', 'btnSalida'].forEach(id => {
            document.getElementById(id)?.classList.remove('active-mode');
        });
        const activeBtnId = `btn${newMode.charAt(0).toUpperCase() + newMode.slice(1)}`;
        const activeBtn = document.getElementById(activeBtnId === 'btnRegVisita' ? 'btnRegVisita' : activeBtnId);
        activeBtn?.classList.add('active-mode');

        // Quitar clase activa de botones de fichador
        ['btnFichadorEntrada', 'btnFichadorSalida'].forEach(id => {
            document.getElementById(id)?.classList.remove('active-mode');
        });
        
        const header = document.querySelector('.header');

        // Gestionar clases de modo en el body y título
        document.body.classList.remove('control-acceso-mode', 'registrar-visita-mode', 'salida-visita-mode', 'registrar-entrada-mode', 'registrar-salida-mode');
        header.classList.remove('control-acceso-mode', 'registrar-visita-mode', 'salida-visita-mode', 'registrar-entrada-mode', 'registrar-salida-mode');

        if (newMode === 'entrada') {
            document.body.classList.add('control-acceso-mode');
            header.classList.add('control-acceso-mode');
            headerTitle.textContent = 'Control de Acceso';
        }
        if (newMode === 'visita') {
            document.body.classList.add('registrar-visita-mode');
            header.classList.add('registrar-visita-mode');
            headerTitle.textContent = 'Registrar Visita';
        }
        if (newMode === 'salida') {
            document.body.classList.add('salida-visita-mode');
            header.classList.add('salida-visita-mode');
            headerTitle.textContent = 'Salida de Visita';
        }

        if (dniInput) dniInput.focus();
    }

    function setPunchMode(newPunchMode) {
        currentPunchMode = newPunchMode;
        currentMode = null; // Desactiva el modo de control de acceso
        const dniInput = document.getElementById('dniInput');
        const headerTitle = document.getElementById('header-title');

        // Quitar clase activa de botones de control
        ['btnEntrada', 'btnRegVisita', 'btnSalida'].forEach(id => {
            document.getElementById(id)?.classList.remove('active-mode');
        });

        const header = document.querySelector('.header');

        // Gestionar clases activas en botones de fichador
        ['btnFichadorEntrada', 'btnFichadorSalida'].forEach(id => {
            document.getElementById(id)?.classList.remove('active-mode');
        });
        const activeBtn = document.getElementById(newPunchMode === 'punch-in' ? 'btnFichadorEntrada' : 'btnFichadorSalida');
        activeBtn?.classList.add('active-mode');

        // Gestionar clases de modo en el body
        document.body.classList.remove('control-acceso-mode', 'registrar-visita-mode', 'salida-visita-mode', 'registrar-entrada-mode', 'registrar-salida-mode');
        header.classList.remove('control-acceso-mode', 'registrar-visita-mode', 'salida-visita-mode', 'registrar-entrada-mode', 'registrar-salida-mode');

        if (newPunchMode === 'punch-in') {
            document.body.classList.add('registrar-entrada-mode');
            header.classList.add('registrar-entrada-mode');
            headerTitle.textContent = 'Registrar Entrada';
        }
        if (newPunchMode === 'punch-out') {
            document.body.classList.add('registrar-salida-mode');
            header.classList.add('registrar-salida-mode');
            headerTitle.textContent = 'Registrar Salida';
        }

        if (dniInput) dniInput.focus();
    }

    // --- LÓGICA DE LA PÁGINA DE ADMISIÓN (index.html) ---
    if (document.getElementById('dniInput')) {
        const dniInput = document.getElementById('dniInput');
        const resultadoDiv = document.getElementById('resultado');
        const mensajeResultado = document.getElementById('mensajeResultado');
        const nombrePersona = document.getElementById('nombrePersona');
        const infoLocal = document.getElementById('infoLocal');
        const infoTarea = document.getElementById('infoTarea');
        const infoVence = document.getElementById('infoVence');
        const ingresosListBody = document.getElementById('ingresosListBody');

        document.getElementById('btnEntrada').addEventListener('click', () => setMode('entrada'));
        document.getElementById('btnRegVisita').addEventListener('click', () => setMode('visita'));
        document.getElementById('btnSalida').addEventListener('click', () => setMode('salida'));
        document.getElementById('btnFichadorEntrada').addEventListener('click', () => setPunchMode('punch-in'));
        document.getElementById('btnFichadorSalida').addEventListener('click', () => setPunchMode('punch-out'));

        document.getElementById('searchIngresos').addEventListener('keyup', (e) => {
            const searchTerm = e.target.value.toLowerCase();
            const rows = ingresosListBody.getElementsByTagName('tr');
            Array.from(rows).forEach(row => {
                row.style.display = row.textContent.toLowerCase().includes(searchTerm) ? '' : 'none';
            });
        });

        async function handleDniScan() {
            if (currentPunchMode) {
                registrarFichaje();
            } else {
                verificarDNI();
            }
        }

        async function verificarDNI() {
            const scannerData = dniInput.value.trim();
            if (!scannerData) return;

            try {
                const response = await fetch('/verificar_dni', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ scanner_data: scannerData, mode: currentMode })
                });
                const data = await response.json();
                
                const vencimientoTexto = data.vence && data.vence !== 'N/A' ? (data.acceso === 'PERMITIDO' ? `Vence: ${data.vence}`: `Vencido el: ${data.vence}`) : '';
                mostrarResultado(data.acceso === 'PERMITIDO' ? 'verde' : 'roja', data.mensaje, data.nombre, data.local, data.tarea, vencimientoTexto);

            } catch (error) {
                console.error('Error al verificar DNI:', error);
                mostrarResultado('red', 'Error de red', '', '', '', '');
            } finally {
                dniInput.value = '';
                dniInput.focus();
                updateStats();
                loadDailyRecords(ingresosListBody);
            }
        }

        function mostrarResultado(colorClass, mensaje, nombre, localInfo, tareaInfo, venceInfo) {
            resultadoDiv.className = `result-display luz-${colorClass}`;
            mensajeResultado.textContent = mensaje;
            nombrePersona.textContent = nombre && nombre !== 'No Encontrado' ? `Nombre: ${nombre}` : '';
            infoLocal.textContent = '';
            infoTarea.textContent = '';
            infoVence.textContent = '';

            if (colorClass === 'verde' || colorClass === 'gris' || (nombre && nombre !== 'No Encontrado')) {
                infoLocal.textContent = localInfo && localInfo !== 'N/A' ? `${localInfo}` : '';
                infoTarea.textContent = tareaInfo && tareaInfo !== 'N/A' ? `${tareaInfo}` : '';
                infoVence.textContent = venceInfo;
            }
        }

        dniInput.focus();
        dniInput.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                handleDniScan();
            }
        });
        
        loadDailyRecords(ingresosListBody);
        updateStats();
        setInterval(() => {
            updateStats();
            loadDailyRecords(ingresosListBody);
        }, 5000);

        setMode('entrada');
    }

    // --- LÓGICA DE LA PÁGINA DE LOGIN ---
    if (document.getElementById('loginBtn')) {
        const loginBtn = document.getElementById('loginBtn');
        const usernameInput = document.getElementById('username');
        const passwordInput = document.getElementById('password');
        const loginMessage = document.getElementById('loginMessage');

        async function handleLogin() {
            const username = usernameInput.value;
            const password = passwordInput.value;

            try {
                const response = await fetch('/perform_login', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ username, password })
                });
                const data = await response.json();
                if (data.success) {
                    window.location.href = '/admin';
                } else {
                    loginMessage.textContent = data.message || 'Error al iniciar sesión.';
                }
            } catch (error) {
                console.error('Error en el login:', error);
                loginMessage.textContent = 'Error de red al intentar iniciar sesión.';
            }
        }

        loginBtn.addEventListener('click', handleLogin);
        [usernameInput, passwordInput].forEach(input => {
            input.addEventListener('keydown', (event) => {
                if (event.key === 'Enter') handleLogin();
            });
        });
    }

    // --- LÓGICA DE LA PÁGINA DE ADMIN ---
    if (document.body.contains(document.getElementById('adminIngresosListBody'))) {
        
        // Carga inicial de registros
        const adminIngresosListBody = document.getElementById('adminIngresosListBody');
        loadDailyRecords(adminIngresosListBody);
        setInterval(() => loadDailyRecords(adminIngresosListBody), 10000);

        // Carga de FAP/FAO
        const fapFileInput = document.getElementById('fapFileInput');
        const faoFileInput = document.getElementById('faoFileInput');
        const fapFileName = document.getElementById('fapFileName');
        const faoFileName = document.getElementById('faoFileName');
        const uploadExcelBtn = document.getElementById('uploadExcelBtn');
        const uploadStatus = document.getElementById('uploadStatus');

        fapFileInput.addEventListener('change', () => {
            fapFileName.textContent = fapFileInput.files.length > 0 ? fapFileInput.files[0].name : 'Ningún archivo seleccionado';
        });

        faoFileInput.addEventListener('change', () => {
            faoFileName.textContent = faoFileInput.files.length > 0 ? faoFileInput.files[0].name : 'Ningún archivo seleccionado';
        });

        uploadExcelBtn.addEventListener('click', async () => {
            const formData = new FormData();
            if (fapFileInput.files.length > 0) {
                formData.append('fapFile', fapFileInput.files[0]);
            }
            if (faoFileInput.files.length > 0) {
                formData.append('faoFile', faoFileInput.files[0]);
            }

            if (formData.has('fapFile') || formData.has('faoFile')) {
                try {
                    const response = await fetch('/upload_excel', {
                        method: 'POST',
                        body: formData
                    });
                    const data = await response.json();
                    uploadStatus.textContent = data.message;
                    uploadStatus.className = data.success ? 'status-message success' : 'status-message error';
                } catch (error) {
                    console.error('Error al subir archivos:', error);
                    uploadStatus.textContent = 'Error de red al intentar subir los archivos.';
                    uploadStatus.className = 'status-message error';
                }
            } else {
                uploadStatus.textContent = 'Por favor, seleccione al menos un archivo (FAP o FAO).';
                uploadStatus.className = 'status-message error';
            }
        });

        // Autorizar Excepción
        const addExcepcionBtn = document.getElementById('addExcepcionBtn');
        addExcepcionBtn.addEventListener('click', async () => {
            const nombre = document.getElementById('nombreExcepcion').value;
            const apellido = document.getElementById('apellidoExcepcion').value;
            const dni = document.getElementById('dniExcepcion').value;
            const local = document.getElementById('localExcepcion').value;
            const autoriza = document.getElementById('autorizaExcepcion').value;
            const excepcionStatus = document.getElementById('excepcionStatus');

            if (nombre && apellido && dni && local && autoriza) {
                try {
                    const response = await fetch('/agregar_excepcion', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ nombre, apellido, dni, local, autoriza })
                    });
                    const data = await response.json();
                    excepcionStatus.textContent = data.message;
                    excepcionStatus.className = data.success ? 'status-message success' : 'status-message error';
                    if (data.success) {
                        document.getElementById('nombreExcepcion').value = '';
                        document.getElementById('apellidoExcepcion').value = '';
                        document.getElementById('dniExcepcion').value = '';
                        document.getElementById('localExcepcion').value = '';
                        document.getElementById('autorizaExcepcion').value = '';
                    }
                } catch (error) {
                    console.error('Error al agregar excepción:', error);
                    excepcionStatus.textContent = 'Error de red al intentar agregar la excepción.';
                    excepcionStatus.className = 'status-message error';
                }
            } else {
                excepcionStatus.textContent = 'Por favor, complete todos los campos para la excepción.';
                excepcionStatus.className = 'status-message error';
            }
        });

        // Carga de Nómina Persistente
        const btnPreviewNomina = document.getElementById('btnPreviewNomina');
        const btnSaveNomina = document.getElementById('btnSaveNomina');
        const nominaTexto = document.getElementById('nominaTexto');
        const previewContainer = document.getElementById('previewNominaContainer');
        const previewBody = document.getElementById('previewNominaBody');
        const nominaParseStatus = document.getElementById('nominaParseStatus');
        const nominaSaveStatus = document.getElementById('nominaSaveStatus');
        const nominaEmpresa = document.getElementById('nominaEmpresa');
        const nominaVigenciaDesde = document.getElementById('nominaVigenciaDesde');
        const nominaVigenciaHasta = document.getElementById('nominaVigenciaHasta');
        const nominasGuardadasBody = document.getElementById('nominasGuardadasBody');

        let isUpdateMode = false;
        let originalNominaData = {};

        async function loadNominasGuardadas() {
            try {
                const response = await fetch('/get_nominas_guardadas');
                const data = await response.json();
                nominasGuardadasBody.innerHTML = '';
                if (data.success && data.nominas) {
                    data.nominas.forEach(n => {
                        const row = nominasGuardadasBody.insertRow();
                        row.insertCell().textContent = n.empresa;
                        row.insertCell().textContent = n.vigencia;
                        row.insertCell().textContent = n.cantidad_personas; // Celda de cantidad
                        
                        const actionsCell = row.insertCell();
                        const editBtn = document.createElement('button');
                        editBtn.textContent = 'Editar';
                        editBtn.classList.add('action-button', 'edit');
                        editBtn.onclick = () => setupEditMode(n.empresa, n.vigencia);
                        
                        const deleteBtn = document.createElement('button');
                        deleteBtn.textContent = 'Borrar';
                        deleteBtn.classList.add('action-button', 'delete');
                        deleteBtn.onclick = () => deleteNomina(n.empresa, n.vigencia);

                        actionsCell.appendChild(editBtn); // Añadir botón de editar
                        actionsCell.appendChild(deleteBtn); // Añadir botón de borrar
                    });
                }
            } catch (error) {
                console.error('Error al cargar nóminas guardadas:', error);
            }
        }

        function renderPreviewTable(personas) {
            previewBody.innerHTML = '';
            personas.forEach(p => {
                const row = previewBody.insertRow();
                row.insertCell().innerHTML = `<input type="text" value="${p.dni || ''}" class="editable-input">`;
                row.insertCell().innerHTML = `<input type="text" value="${p.apellido || ''}" class="editable-input">`;
                row.insertCell().innerHTML = `<input type="text" value="${p.nombre || ''}" class="editable-input">`;
            });
            previewContainer.style.display = 'block';
        }

        btnPreviewNomina.addEventListener('click', async () => {
            const texto = nominaTexto.value;
            if (!texto.trim()) {
                nominaParseStatus.textContent = 'El campo de texto está vacío.';
                nominaParseStatus.className = 'status-message error';
                return;
            }
            try {
                const response = await fetch('/parse_nomina', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ texto_pegado: texto })
                });
                const data = await response.json();
                if (data.success) {
                    renderPreviewTable(data.nomina);
                    nominaParseStatus.textContent = `Se previsualizaron ${data.nomina.length} personas.`;
                    nominaParseStatus.className = 'status-message success';
                } else {
                    nominaParseStatus.textContent = data.message || 'No se pudo procesar el texto.';
                    nominaParseStatus.className = 'status-message error';
                    previewContainer.style.display = 'none';
                }
            } catch (error) {
                console.error('Error al previsualizar:', error);
                nominaParseStatus.textContent = 'Error de red.';
                nominaParseStatus.className = 'status-message error';
            }
        });

        btnSaveNomina.addEventListener('click', async () => {
            const nomina = Array.from(previewBody.rows).map(row => ({
                dni: row.cells[0].querySelector('input').value.trim(),
                apellido: row.cells[1].querySelector('input').value.trim(),
                nombre: row.cells[2].querySelector('input').value.trim()
            })).filter(p => p.dni);

            if (nomina.length === 0) {
                nominaSaveStatus.textContent = 'La tabla de previsualización está vacía.';
                nominaSaveStatus.className = 'status-message error';
                return;
            }

            const payload = {
                nomina: nomina,
                empresa: nominaEmpresa.value,
                vigencia_desde: nominaVigenciaDesde.value,
                vigencia_hasta: nominaVigenciaHasta.value,
                is_update: isUpdateMode,
                original_empresa: isUpdateMode ? originalNominaData.empresa : null,
                original_vigencia_desde: isUpdateMode ? originalNominaData.vigencia_desde : null,
            };

            try {
                const response = await fetch('/save_nomina', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const data = await response.json();
                nominaSaveStatus.textContent = data.message;
                nominaSaveStatus.className = data.success ? 'status-message success' : 'status-message error';
                if (data.success) {
                    resetNominaForm();
                    loadNominasGuardadas();
                }
            } catch (error) {
                console.error('Error al guardar:', error);
                nominaSaveStatus.textContent = 'Error de red.';
                nominaSaveStatus.className = 'status-message error';
            }
        });

        async function setupEditMode(empresa, vigencia) {
            try {
                const response = await fetch('/get_nomina_detalle', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ empresa, vigencia })
                });
                const data = await response.json();
                if (data.success) {
                    const detalle = data.detalle;
                    
                    nominaEmpresa.value = empresa;
                    nominaVigenciaDesde.value = detalle.vigencia_desde;
                    nominaVigenciaHasta.value = detalle.vigencia_hasta;
                    nominaTexto.value = '';
                    
                    renderPreviewTable(detalle.personas);
                    
                    isUpdateMode = true;
                    originalNominaData = { empresa, vigencia };
                    
                    btnSaveNomina.textContent = 'Actualizar Nómina';
                    document.getElementById('management-title').textContent = 'Editando Nómina';
                    window.scrollTo(0, 0);
                } else {
                    alert(data.message);
                }
            } catch (error) {
                console.error('Error al cargar para editar:', error);
            }
        }

        async function deleteNomina(empresa, vigencia) {
            if (!confirm(`¿Seguro que desea eliminar la nómina de "${empresa}" con vigencia ${vigencia}?`)) return;

            try {
                const response = await fetch('/delete_nomina', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ empresa, vigencia })
                });
                const data = await response.json();
                alert(data.message);
                if (data.success) {
                    loadNominasGuardadas();
                }
            } catch (error) {
                console.error('Error al eliminar:', error);
                alert('Error de red al intentar eliminar.');
            }
        }

        function resetNominaForm() {
            nominaEmpresa.value = '';
            nominaVigenciaDesde.value = '';
            nominaVigenciaHasta.value = '';
            nominaTexto.value = '';
            previewContainer.style.display = 'none';
            previewBody.innerHTML = '';
            isUpdateMode = false;
            originalNominaData = {};
            btnSaveNomina.textContent = 'Guardar Nómina';
            document.getElementById('management-title').textContent = 'Gestión de Nóminas Persistentes';
            nominaParseStatus.textContent = '';
            nominaSaveStatus.textContent = '';
        }

        loadNominasGuardadas();

        // Descargar Reporte Diario
        const descargarReporteBtn = document.getElementById('descargarReporteBtn');
        if (descargarReporteBtn) {
            descargarReporteBtn.addEventListener('click', () => {
                window.location.href = '/descargar_reporte_diario';
            });
        }

        const descargarFichajesBtn = document.getElementById('descargarFichajesBtn');
        if (descargarFichajesBtn) {
            descargarFichajesBtn.addEventListener('click', () => {
                window.location.href = '/descargar_reporte_fichajes';
            });
        }

        // Enviar Reporte por Email
        const enviarReporteBtn = document.getElementById('enviarReporteBtn');
        if (enviarReporteBtn) {
            enviarReporteBtn.addEventListener('click', async () => {
                const reporteStatus = document.getElementById('reporteStatus');
                reporteStatus.textContent = 'Enviando reporte...';
                reporteStatus.className = 'status-message'; // Reset class

                try {
                    const response = await fetch('/send_report_email', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' }
                    });
                    const data = await response.json();
                    reporteStatus.textContent = data.message;
                    reporteStatus.className = data.success ? 'status-message success' : 'status-message error';
                } catch (error) {
                    console.error('Error al enviar reporte por email:', error);
                    reporteStatus.textContent = 'Error de red al intentar enviar el reporte.';
                    reporteStatus.className = 'status-message error';
                }
            });
        }


    }
});
