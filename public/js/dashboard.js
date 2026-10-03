/* =============================================
   DASHBOARD JS - Control de Acceso AOS (FIREBASE)
   ============================================= */
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getAuth, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-auth.js";
import { getFirestore, collection, query, where, onSnapshot, getDocs, doc, getDoc, setDoc, addDoc, serverTimestamp, deleteDoc } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-firestore.js";

const firebaseConfig = {
    apiKey: "AIzaSyBg0ht9KWgGrhGmHkT7mzCRgJ2CS6e4-lQ",
    authDomain: "control-acceso-c2f82.firebaseapp.com",
    projectId: "control-acceso-c2f82",
    storageBucket: "control-acceso-c2f82.firebasestorage.app",
    messagingSenderId: "654089830879",
    appId: "1:654089830879:web:96b5a40c8fc6292422446a"
};
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
// Importaciones combinadas en la linea 5
// --- Estado global ---
let accesosData = [];
let fichajesData = [];
let auditoriaData = [];
// Eliminamos REFRESH_INTERVAL porque ahora es Tiempo Real!

// --- Inicialización ---
document.addEventListener('DOMContentLoaded', () => {

    // --- PROTECCIÓN DE LOGIN ---
    onAuthStateChanged(auth, async (user) => {
        if (!user) {
            window.location.href = '/index.html';
            return;
        }
        
        try {
            const userDoc = await getDoc(doc(db, "usuarios", user.uid));
            if (userDoc.exists()) {
                const data = userDoc.data();
                if (data.rol !== 'supervisor' && data.rol !== 'super_admin') {
                    alert("Acceso denegado. No tienes permisos de supervisor.");
                    window.location.href = '/scan.html';
                }
                
                // Solo el Super Admin puede ver y crear usuarios
                if (data.rol !== 'super_admin') {
                    const tabUsuarios = document.querySelector('.dash-tab[data-tab="usuarios"]');
                    if(tabUsuarios) tabUsuarios.style.display = 'none';
                }
            } else {
                signOut(auth);
            }
        } catch(e) {
            console.error(e);
        }
    });

    const btnLogout = document.getElementById('btnLogout');
    if (btnLogout) {
        btnLogout.addEventListener('click', () => {
            signOut(auth);
        });
    }

    initTabs();
    initFilters();
    initAuditoria();
    initConfig();
    loadDashboardData();
});

// =============================================
// TABS
// =============================================
function initTabs() {
    document.querySelectorAll('.dash-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('.dash-tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.dash-panel').forEach(p => p.classList.add('hidden'));
            tab.classList.add('active');
            const panelId = 'panel-' + tab.dataset.tab;
            const panel = document.getElementById(panelId);
            if (panel) panel.classList.remove('hidden');
        });
    });
}

window.filterAndGoTo = function(tabName, filterId, filterValue) {
    const tabBtn = document.querySelector(`.dash-tab[data-tab="${tabName}"]`);
    if (tabBtn) tabBtn.click();
    
    if (filterId) {
        const input = document.getElementById(filterId);
        if (input) {
            input.value = filterValue;
            input.dispatchEvent(new Event('input'));
            input.dispatchEvent(new Event('change'));
        }
    }
    document.querySelector('.dash-tabs-container').scrollIntoView({ behavior: 'smooth' });
};

// =============================================
// CARGA DE DATOS (FIREBASE REALTIME)
// =============================================
async function loadDashboardData() {
    setupAccesosListener();
    setupListaNegraListener();
    setupNominasListener();
    // Las demás funciones se migrarán después
    // loadFichajes();
}

function formatDateToLocalString() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

function setupAccesosListener() {
    const today = formatDateToLocalString();
    const q = query(collection(db, "accesos"), where("Fecha", "==", today));
    
    onSnapshot(q, (snapshot) => {
        accesosData = [];
        let entradas = 0;
        let salidas = 0;
        let rechazos = 0;
        
        snapshot.forEach((doc) => {
            const data = doc.data();
            accesosData.push(data);
            
            const res = (data.Resultado || '').toUpperCase();
            if (res.includes('AUTORIZADO') || res === 'VERDE' || res.includes('REGISTRADO')) {
                if (data.Evento === 'ENTRADA' || data.Evento === 'Entrada OK') entradas++;
                if (data.Evento === 'SALIDA' || data.Evento === 'Salida') salidas++;
            } else if (res.includes('DENEGADO') || res === 'ROJO') {
                rechazos++;
            }
        });
        
        // Ordenar por hora descendente
        accesosData.sort((a, b) => {
            const timeA = a.Hora_Ingreso || a.Hora_Salida || '00:00:00';
            const timeB = b.Hora_Ingreso || b.Hora_Salida || '00:00:00';
            return timeB.localeCompare(timeA);
        });
        
        // Actualizar UI
        renderAccesos();
        updateStatsUI(entradas, salidas, rechazos);
    });
}

function updateStatsUI(entradas, salidas, rechazos) {
    document.getElementById('statEntradas').textContent = entradas;
    document.getElementById('statSalidas').textContent = salidas;
    document.getElementById('statRechazos').textContent = rechazos;
    
    // Escuchar live_stats desde Firebase en vez del servidor local
    const docRef = doc(db, 'config', 'live_stats');
    onSnapshot(docRef, (docSnap) => {
        if (docSnap.exists()) {
            const data = docSnap.data();
            const statAdentro = document.getElementById('statAdentro');
            if (statAdentro) {
                statAdentro.textContent = data.total_adentro || 0;
            }
        }
    });
    
    const now = new Date();
    document.getElementById('lastUpdate').textContent = now.toLocaleTimeString('es-AR', {hour12: false});
    
    animateStatCards();
}

// =============================================
// RENDERIZADO DE TABLAS
// =============================================
function renderAccesos() {
    const tbody = document.getElementById('bodyAccesos');
    if (!tbody) return;
    const filterDniEl = document.getElementById('filterAccesoDni');
    const filterResultadoEl = document.getElementById('filterAccesoResultado');
    const filterDni = filterDniEl ? filterDniEl.value.toLowerCase() : '';
    const filterResultado = filterResultadoEl ? filterResultadoEl.value : '';

    let filtered = accesosData;

    if (filterDni) {
        filtered = filtered.filter(r => 
            (r.DNI || '').toLowerCase().includes(filterDni) ||
            (r['Nombre y Apellido'] || '').toLowerCase().includes(filterDni)
        );
    }
    if (filterResultado) {
        filtered = filtered.filter(r => (r.Resultado || '').toUpperCase().includes(filterResultado));
    }

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="empty-state">No hay registros de acceso para hoy</td></tr>';
        return;
    }

    tbody.innerHTML = filtered.map(r => {
        const hora = r.Hora_Ingreso || r.Hora_Salida || '';
        const resultado = (r.Resultado || '').toUpperCase();
        let badgeClass = 'badge-info';
        if (resultado.includes('AUTORIZADO') || resultado === 'VERDE') badgeClass = 'badge-success';
        else if (resultado.includes('DENEGADO') || resultado === 'ROJO') badgeClass = 'badge-danger';
        else if (resultado.includes('REGISTRADO')) badgeClass = 'badge-info';

        let comentarioHtml = r.Comentario ? `<br><small style="font-style: italic; color: #fbbf24;">📝 ${r.Comentario}</small>` : '';

        return `<tr>
            <td>${hora}</td>
            <td><strong>${r.DNI || ''}</strong></td>
            <td>${r['Nombre y Apellido'] || ''}${comentarioHtml}</td>
            <td>${r.Evento || ''}</td>
            <td>${r.Tipo_Permiso || ''}</td>
            <td>${r.Local || ''}</td>
            <td><span class="badge ${badgeClass}">${r.Resultado || ''}</span></td>
        </tr>`;
    }).join('');
}

function renderFichajes() {
    const tbody = document.getElementById('bodyFichajes');
    if (!tbody) return;
    const filterDniEl = document.getElementById('filterFichajeDni');
    const filterDni = filterDniEl ? filterDniEl.value.toLowerCase() : '';

    let filtered = fichajesData;

    if (filterDni) {
        filtered = filtered.filter(r => 
            (r.DNI || '').toLowerCase().includes(filterDni) ||
            (r['Nombre y Apellido'] || '').toLowerCase().includes(filterDni)
        );
    }

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-state">No hay fichajes registrados hoy</td></tr>';
        return;
    }

    tbody.innerHTML = filtered.map(r => {
        const tieneSalida = r.Hora_Salida && r.Hora_Salida !== '' && r.Hora_Salida !== 'nan' && r.Hora_Salida !== null;

        return `<tr>
            <td><strong>${r.DNI || ''}</strong></td>
            <td>${r['Nombre y Apellido'] || ''}</td>
            <td><span class="badge badge-info">${r.Categoria || 'N/A'}</span></td>
            <td>${r.Puesto || '-'}</td>
            <td>${r.Hora_Entrada || ''}</td>
            <td>${tieneSalida ? r.Hora_Salida : '<span class="badge badge-warning">En curso</span>'}</td>
            <td>${tieneSalida ? (r.horas_diurnas || '0') + 'h' : '-'}</td>
            <td>${tieneSalida ? (r.horas_nocturnas || '0') + 'h' : '-'}</td>
        </tr>`;
    }).join('');
}

function renderPersonasAdentro(personas) {
    const grid = document.getElementById('personasAdentroGrid');

    if (!personas || personas.length === 0) {
        grid.innerHTML = '<p class="empty-state">No hay personas dentro del predio en este momento</p>';
        return;
    }

    grid.innerHTML = personas.map(p => {
        const tipo = (p.tipo_permiso || 'default').toLowerCase();
        let iconClass = 'tipo-default';
        let icon = '👤';

        if (tipo.includes('fap')) { iconClass = 'tipo-fap'; icon = '🟢'; }
        else if (tipo.includes('fao')) { iconClass = 'tipo-fao'; icon = '🟡'; }
        else if (tipo.includes('nomina')) { iconClass = 'tipo-nomina'; icon = '🟣'; }
        else if (tipo.includes('excepcion')) { iconClass = 'tipo-excepcion'; icon = '🔵'; }
        else if (tipo.includes('visita')) { iconClass = 'tipo-visita'; icon = '🔴'; }

        return `<div class="persona-card">
            <div class="persona-icon ${iconClass}">${icon}</div>
            <div class="persona-info">
                <span class="persona-dni">${p.dni}</span>
                <span class="persona-tipo">${p.tipo_permiso}</span>
            </div>
        </div>`;
    }).join('');
}

// =============================================
// FILTROS
// =============================================
function initFilters() {
    const filterAccDni = document.getElementById('filterAccesoDni');
    const filterAccRes = document.getElementById('filterAccesoResultado');
    const filterFichDni = document.getElementById('filterFichajeDni');
    if (filterAccDni) filterAccDni.addEventListener('input', renderAccesos);
    if (filterAccRes) filterAccRes.addEventListener('change', renderAccesos);
    if (filterFichDni) filterFichDni.addEventListener('input', renderFichajes);
}

// =============================================
// AUDITORÍA
// =============================================
function initAuditoria() {
    // Default: últimos 7 días
    const hoy = new Date();
    const hace7 = new Date(hoy);
    hace7.setDate(hace7.getDate() - 7);

    document.getElementById('auditHasta').value = formatDate(hoy);
    document.getElementById('auditDesde').value = formatDate(hace7);

    document.getElementById('btnBuscarAudit').addEventListener('click', buscarAuditoria);
    document.getElementById('btnExportarAudit').addEventListener('click', exportarAuditoria);
}

async function buscarAuditoria() {
    const desde = document.getElementById('auditDesde').value;
    const hasta = document.getElementById('auditHasta').value;
    const tipo = document.getElementById('auditTipo').value;

    if (!desde || !hasta) {
        alert('Seleccioná ambas fechas');
        return;
    }

    const btn = document.getElementById('btnBuscarAudit');
    btn.disabled = true;
    btn.innerHTML = '<span class="material-icons">hourglass_top</span> Buscando...';

    try {
        const res = await fetch(`/api/dashboard/auditoria?desde=${desde}&hasta=${hasta}&tipo=${tipo}`);
        const data = await res.json();

        if (data.success) {
            auditoriaData = data.records;
            renderAuditoria(tipo);

            const info = document.getElementById('auditResultInfo');
            info.style.display = 'flex';
            info.innerHTML = `<span class="material-icons">info</span> ${data.total} registros encontrados — Período: ${data.periodo}`;
        } else {
            alert(data.message);
        }
    } catch (e) {
        console.error('Error en auditoría:', e);
        alert('Error al buscar registros');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<span class="material-icons">search</span> Buscar';
    }
}

function renderAuditoria(tipo) {
    const thead = document.getElementById('headAuditoria');
    const tbody = document.getElementById('bodyAuditoria');

    if (auditoriaData.length === 0) {
        thead.innerHTML = '<tr><th colspan="7" class="empty-state">No se encontraron registros para el período seleccionado</th></tr>';
        tbody.innerHTML = '';
        return;
    }

    if (tipo === 'fichajes') {
        thead.innerHTML = `<tr>
            <th>Fecha</th><th>DNI</th><th>Nombre</th>
            <th>Hora Entrada</th><th>Hora Salida</th>
        </tr>`;
        tbody.innerHTML = auditoriaData.map(r => `<tr>
            <td>${r.Fecha || ''}</td>
            <td><strong>${r.DNI || ''}</strong></td>
            <td>${r['Nombre y Apellido'] || ''}</td>
            <td>${r.Hora_Entrada || ''}</td>
            <td>${r.Hora_Salida || '—'}</td>
        </tr>`).join('');
    } else {
        thead.innerHTML = `<tr>
            <th>Fecha</th><th>Hora</th><th>DNI</th><th>Nombre</th>
            <th>Evento</th><th>Tipo Permiso</th><th>Resultado</th>
        </tr>`;
        tbody.innerHTML = auditoriaData.map(r => {
            const resultado = (r.Resultado || '').toUpperCase();
            let badgeClass = 'badge-info';
            if (resultado.includes('AUTORIZADO') || resultado === 'VERDE') badgeClass = 'badge-success';
            else if (resultado.includes('DENEGADO') || resultado === 'ROJO') badgeClass = 'badge-danger';

            return `<tr>
                <td>${r.Fecha || ''}</td>
                <td>${r.Hora_Ingreso || r.Hora_Salida || ''}</td>
                <td><strong>${r.DNI || ''}</strong></td>
                <td>${r['Nombre y Apellido'] || ''}</td>
                <td>${r.Evento || ''}</td>
                <td>${r.Tipo_Permiso || ''}</td>
                <td><span class="badge ${badgeClass}">${r.Resultado || ''}</span></td>
            </tr>`;
        }).join('');
    }
}

function exportarAuditoria() {
    const desde = document.getElementById('auditDesde').value;
    const hasta = document.getElementById('auditHasta').value;
    const tipo = document.getElementById('auditTipo').value;

    if (!desde || !hasta) {
        alert('Seleccioná ambas fechas primero');
        return;
    }

    // Usar los endpoints de descarga existentes
    if (tipo === 'fichajes') {
        window.location.href = `/descargar_fichajes_fechas?desde=${desde}&hasta=${hasta}`;
    } else {
        window.location.href = `/descargar_reporte_fechas?desde=${desde}&hasta=${hasta}`;
    }
}

// =============================================
// HELPERS
// =============================================
function formatDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

function animateStatCards() {
    document.querySelectorAll('.stat-value').forEach(el => {
        el.style.transition = 'transform 0.2s ease';
        el.style.transform = 'scale(1.1)';
        setTimeout(() => { el.style.transform = 'scale(1)'; }, 200);
    });
}

// =============================================
// CONFIGURACIÓN IRSA
// =============================================
function initConfig() {
    const btnSave = document.getElementById('btnSaveIrsa');
    const btnSync = document.getElementById('btnSyncIrsa');
    if (btnSave) btnSave.addEventListener('click', saveIrsaConfig);
    if (btnSync) btnSync.addEventListener('click', syncIrsa);
    loadIrsaConfig();
}

async function loadIrsaConfig() {
    try {
        const docRef = doc(db, 'config', 'irsa');
        const docSnap = await getDoc(docRef);
        
        if (docSnap.exists()) {
            const data = docSnap.data();
            if (data.username) {
                document.getElementById('irsaUser').value = data.username;
            }
            if (data.dias_alerta_vencimiento !== undefined) {
                document.getElementById('irsaDiasAlerta').value = data.dias_alerta_vencimiento;
                window.diasAlertaVencimiento = parseInt(data.dias_alerta_vencimiento) || 0;
            }
            
            const hasCredsEl = document.getElementById('irsaHasCreds');
            hasCredsEl.textContent = data.password ? 'Configuradas (OK)' : 'No configuradas';
            hasCredsEl.style.color = data.password ? 'var(--dash-success)' : 'var(--dash-danger)';
            
            // Si el Obrero actualiza la fecha de ultima sincronizacion en este mismo documento:
            const lastSyncTimeEl = document.getElementById('irsaLastSyncTime');
            if (data.last_sync_time) {
                lastSyncTimeEl.textContent = new Date(data.last_sync_time).toLocaleString();
            } else {
                lastSyncTimeEl.textContent = 'Nunca';
            }
        } else {
            const hasCredsEl = document.getElementById('irsaHasCreds');
            hasCredsEl.textContent = 'No configuradas';
            hasCredsEl.style.color = 'var(--dash-danger)';
            document.getElementById('irsaLastSyncTime').textContent = 'Nunca';
        }
    } catch (e) {
        console.error("Error loading IRSA config from Firebase:", e);
    }
}

async function saveIrsaConfig() {
    const userEl = document.getElementById('irsaUser');
    const passEl = document.getElementById('irsaPass');
    const diasEl = document.getElementById('irsaDiasAlerta');
    const msgEl = document.getElementById('irsaStatusMsg');
    const btn = document.getElementById('btnSaveIrsa');

    if (!userEl || !passEl) {
        alert('Error: no se encontraron los campos del formulario IRSA.');
        return;
    }

    const user = userEl.value.trim();
    const pass = passEl.value;
    const diasAlerta = diasEl ? diasEl.value : 0;

    if (!user) {
        const msg = 'El usuario no puede estar vacío.';
        if (msgEl) { msgEl.textContent = msg; msgEl.style.color = 'var(--dash-danger)'; msgEl.style.display = 'block'; }
        else alert(msg);
        return;
    }

    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="material-icons">hourglass_top</span> Guardando...'; }

    try {
        const docRef = doc(db, 'config', 'irsa');
        await setDoc(docRef, {
            username: user,
            password: pass,
            dias_alerta_vencimiento: parseInt(diasAlerta)
        }, { merge: true });

        if (msgEl) {
            msgEl.textContent = 'Credenciales guardadas en la nube.';
            msgEl.style.color = 'var(--dash-success)';
            msgEl.style.display = 'block';
        }

        if (passEl) passEl.value = '';
        
        // Mantener compatibilidad temporal disparando el fetch local en silencio
        fetch('/api/config/irsa', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({ username: user, password: pass, dias_alerta_vencimiento: diasAlerta })
        }).catch(e => console.log('API local no responde (esperado en web)'));

    } catch (e) {
        console.error("Error guardando credenciales en Firebase:", e);
        if (msgEl) {
            msgEl.textContent = 'Error de conexión con la nube.';
            msgEl.style.color = 'var(--dash-danger)';
            msgEl.style.display = 'block';
        }
    } finally {
        if (btn) { btn.disabled = false; btn.innerHTML = '<span class="material-icons">save</span> Guardar Credenciales'; }
    }
}

async function syncIrsa() {
    const btn = document.getElementById('btnSyncIrsa');
    const msgEl = document.getElementById('irsaStatusMsg');
    
    btn.disabled = true;
    btn.innerHTML = '<span class="material-icons">cloud_sync</span> Sincronizando...';
    
    try {
        const res = await fetch('/api/sync/irsa', {method: 'POST'});
        const data = await res.json();
        
        msgEl.textContent = data.message;
        msgEl.style.color = data.success ? 'var(--dash-success)' : 'var(--dash-danger)';
        msgEl.style.display = 'block';
        
        loadIrsaConfig(); // Actualizar el panel de estado abajo
        loadAccesos(); // Recargar tablas si hubo cambios
        loadTramitesIRSA(); // Actualizar trámites FAP/FAO
    } catch (e) {
        msgEl.textContent = 'Error de conexión al servidor';
        msgEl.style.color = 'var(--dash-danger)';
        msgEl.style.display = 'block';
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<span class="material-icons">sync</span> Sincronizar Ahora';
    }
}
// --- NÓMINAS (RRHH) MANAGEMENT ---
let nominasData = [];

function setupNominasListener() {
    const q = query(collection(db, "nominas"));
    onSnapshot(q, (snapshot) => {
        nominasData = [];
        snapshot.forEach((doc) => {
            const data = doc.data();
            data.id = doc.id; // Guardamos el ID del documento (DNI)
            nominasData.push(data);
        });
        renderNominas();
    });
}

function renderNominas() {
    const tbody = document.getElementById('bodyDirectorio');
    if (!tbody) return;
    
    const filterText = (document.getElementById('filterDirectorio')?.value || '').toLowerCase();
    
    let filtered = nominasData.filter(r => {
        const dni = (r.dni || '').toLowerCase();
        const nombre = (r.nombre || '').toLowerCase();
        const empresa = (r.puesto_especifico || r.empresa || '').toLowerCase();
        return dni.includes(filterText) || nombre.includes(filterText) || empresa.includes(filterText);
    });

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-state">No se encontraron registros.</td></tr>';
        return;
    }

    tbody.innerHTML = filtered.map(r => `
        <tr>
            <td><strong>${r.dni}</strong></td>
            <td>${r.nombre}</td>
            <td>${r.puesto_especifico || r.empresa || '-'}</td>
            <td><span class="badge badge-info">${r.categoria || 'N/A'}</span></td>
            <td>${r.nro_tramite || '-'}</td>
            <td>${r.fecha_vencimiento || '-'}</td>
        </tr>
    `).join('');
}

// Modal logic for Nomina
const modalNomina = document.getElementById('modalNomina');
const closeNominaModal = document.getElementById('closeNominaModal');
const btnNuevaNomina = document.getElementById('btnNuevaNomina');
const formNomina = document.getElementById('formNomina');

if(btnNuevaNomina) {
    btnNuevaNomina.onclick = () => {
        document.getElementById('modalNominaTitle').textContent = 'Agregar Empleado a Nómina';
        formNomina.reset();
        document.getElementById('nominaId').value = '';
        modalNomina.style.display = 'flex';
    };
}

if(closeNominaModal) {
    closeNominaModal.onclick = () => modalNomina.style.display = 'none';
}

window.addEventListener('click', (e) => {
    if (e.target == modalNomina) {
        modalNomina.style.display = 'none';
    }
});

function editNomina(r) {
    document.getElementById('modalNominaTitle').textContent = 'Editar Empleado';
    document.getElementById('nominaId').value = r.id;
    document.getElementById('nominaDni').value = r.dni;
    document.getElementById('nominaNombre').value = r.nombre;
    document.getElementById('nominaCategoria').value = r.categoria || '';
    document.getElementById('nominaPuesto').value = r.puesto_especifico || '';
    document.getElementById('nominaActivo').value = r.activo ? '1' : '0';
    modalNomina.style.display = 'flex';
}

async function deleteNomina(id) {
    if(!confirm("¿Está seguro de eliminar este empleado de la nómina?")) return;
    try {
        const res = await fetch(`/api/dashboard/nomina/${id}`, { method: 'DELETE' });
        const data = await res.json();
        if(data.success) {
            loadNominas();
        } else {
            alert("Error al borrar: " + data.message);
        }
    } catch(e) {
        console.error(e);
        alert("Error de red");
    }
}

if(formNomina) {
    formNomina.onsubmit = async (e) => {
        e.preventDefault();
        const payload = {
            id: document.getElementById('nominaId').value,
            dni: document.getElementById('nominaDni').value.trim(),
            nombre: document.getElementById('nominaNombre').value.trim(),
            categoria: document.getElementById('nominaCategoria').value,
            puesto: document.getElementById('nominaPuesto').value.trim(),
            activo: document.getElementById('nominaActivo').value === '1'
        };

        try {
            const res = await fetch('/api/dashboard/nomina', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const data = await res.json();
            if(data.success) {
                modalNomina.style.display = 'none';
                loadNominas();
            } else {
                alert("Error al guardar: " + data.message);
            }
        } catch (e) {
            console.error(e);
            alert("Error de red");
        }
    };
}
// --- FIN NÓMINAS ---
// --- SSE ALERT LISTENER ---
function initSSEAlerts() {
    const eventSource = new EventSource('/stream/alerts');
    const modalAlerta = document.getElementById('modalAlertaSeguridad');
    const btnEnterado = document.getElementById('btnEnteradoAlerta');
    const audio = document.getElementById('alertSound');

    eventSource.onmessage = function(event) {
        try {
            const data = JSON.parse(event.data);
            console.log("ALERTA RECIBIDA:", data);
            
            document.getElementById('alertaNombre').textContent = data.nombre || 'DESCONOCIDO';
            document.getElementById('alertaDNI').textContent = 'DNI: ' + data.dni;
            document.getElementById('alertaMotivo').textContent = data.motivo;
            
            modalAlerta.style.display = 'flex';
            
            if (audio) {
                audio.currentTime = 0;
                audio.play().catch(e => console.error("Auto-play prevented", e));
            }
        } catch (e) {
            console.error("Error parseando alerta", e);
        }
    };

    eventSource.onerror = function(err) {
        console.error("SSE connection error", err);
    };

    if (btnEnterado) {
        btnEnterado.onclick = () => {
            modalAlerta.style.display = 'none';
            if (audio) {
                audio.pause();
                audio.currentTime = 0;
            }
        };
    }
}

document.addEventListener('DOMContentLoaded', () => {
    // Inicializar SSE después de todo
    setTimeout(initSSEAlerts, 1000);
});
// --- LISTA NEGRA (PERSONAS DE INTERÉS) ---
let listaNegraData = [];

function setupListaNegraListener() {
    const q = query(collection(db, "lista_negra"));
    onSnapshot(q, (snapshot) => {
        listaNegraData = [];
        snapshot.forEach((doc) => {
            const data = doc.data();
            data.id = doc.id;
            listaNegraData.push(data);
        });
        renderListaNegra();
    });
}

function renderListaNegra() {
    const tbody = document.getElementById('bodyListaNegra');
    if (!tbody) return;
    
    if (listaNegraData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" class="empty-state">No hay personas en la lista negra</td></tr>';
        return;
    }

    tbody.innerHTML = listaNegraData.map(r => `
        <tr>
            <td><strong>${r.dni}</strong></td>
            <td>${r.nombre || '-'}</td>
            <td><span style="color: #ff3b30; font-weight: 500;">${r.motivo}</span></td>
            <td>
                 <!-- Boton borrar deshabilitado hasta portar funcion -->
                <button class="dash-btn dash-btn-danger" style="padding: 0.25rem 0.5rem; font-size: 0.8rem; opacity: 0.5; cursor: not-allowed;">Quitar</button>
            </td>
        </tr>
    `).join('');
}

async function deleteListaNegra(id) {
    if(!confirm("¿Quitar a esta persona de la Lista Negra? Ya no generará alertas silenciosas.")) return;
    try {
        const res = await fetch(`/api/dashboard/lista_negra/${id}`, { method: 'DELETE' });
        const data = await res.json();
        if(data.success) loadListaNegra();
    } catch(e) {
        console.error(e);
        alert("Error de red");
    }
}

const formListaNegra = document.getElementById('formListaNegra');
if (formListaNegra) {
    formListaNegra.onsubmit = async (e) => {
        e.preventDefault();
        const payload = {
            dni: document.getElementById('lnDni').value.trim(),
            nombre: document.getElementById('lnNombre').value.trim(),
            motivo: document.getElementById('lnMotivo').value.trim()
        };

        try {
            const res = await fetch('/api/dashboard/lista_negra', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const data = await res.json();
            if(data.success) {
                formListaNegra.reset();
                loadListaNegra();
            } else {
                alert("Error al guardar: " + data.message);
            }
        } catch (e) {
            console.error(e);
            alert("Error de red");
        }
    };
}
// --- MÓDULO OTA: ACTUALIZACIONES REMOTAS Y REINICIO ---
const formOtaUpdate = document.getElementById('formOtaUpdate');
const otaStatusMsg = document.getElementById('otaStatusMsg');
const btnRestartApp = document.getElementById('btnRestartApp');

function showOtaStatus(msg, type = 'info') {
    otaStatusMsg.textContent = msg;
    otaStatusMsg.style.display = 'block';
    if (type === 'error') {
        otaStatusMsg.style.color = '#ff3b30';
    } else if (type === 'success') {
        otaStatusMsg.style.color = '#34c759';
    } else {
        otaStatusMsg.style.color = 'var(--dash-text-muted)';
    }
}

async function waitForServerRestart() {
    let retries = 0;
    const maxRetries = 15; // 15 segundos aprox
    
    // Crear el modal de "Reiniciando" si no existe
    let restartModal = document.getElementById('restartModalOverlay');
    if (!restartModal) {
        restartModal = document.createElement('div');
        restartModal.id = 'restartModalOverlay';
        restartModal.innerHTML = `
            <div style="position: fixed; top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.8);z-index:99999;display:flex;flex-direction:column;justify-content:center;align-items:center;color:white;font-family:sans-serif;">
                <span class="material-icons" style="font-size:4rem; margin-bottom:1rem; animation: spin 2s linear infinite;">autorenew</span>
                <h2>Reiniciando el Servidor de la Garita...</h2>
                <p>Aplicando cambios y reconectando...</p>
                <style>@keyframes spin { 100% { transform:rotate(360deg); } }</style>
            </div>
        `;
        document.body.appendChild(restartModal);
    }
    restartModal.style.display = 'flex';

    const checkServer = async () => {
        try {
            // Intentar cargar la raíz o un endpoint que sepamos que existe (ping)
            await fetch('/', { method: 'HEAD', cache: 'no-store' });
            // Si no da error, el server está up!
            window.location.reload(); 
        } catch (e) {
            retries++;
            if (retries > maxRetries) {
                restartModal.innerHTML = `
                    <div style="position: fixed; top:0;left:0;width:100%;height:100%;background:rgba(255,0,0,0.9);z-index:99999;display:flex;flex-direction:column;justify-content:center;align-items:center;color:white;">
                        <span class="material-icons" style="font-size:4rem;">error</span>
                        <h2>Fallo al reconectar</h2>
                        <p>El servidor está tardando demasiado en reiniciar. Verifica la PC de la garita.</p>
                        <button onclick="window.location.reload()" style="padding:1rem;margin-top:1rem;background:white;color:black;border:none;cursor:pointer;">Reintentar Conexión</button>
                    </div>
                `;
            } else {
                setTimeout(checkServer, 1000); // intentar cada segundo
            }
        }
    };
    
    // Esperar 2 segundos antes del primer chequeo para darle tiempo a morir
    setTimeout(checkServer, 2000);
}

if (formOtaUpdate) {
    formOtaUpdate.onsubmit = async (e) => {
        e.preventDefault();
        
        const fileInput = document.getElementById('otaFile');
        const pwdInput = document.getElementById('otaPassword');
        
        if (!fileInput.files.length) return;
        
        const file = fileInput.files[0];
        const pwd = pwdInput.value;
        
        const formData = new FormData();
        formData.append('file', file);
        formData.append('password', pwd);
        
        const btn = document.getElementById('btnUploadOta');
        btn.disabled = true;
        btn.innerHTML = '<span class="material-icons">hourglass_empty</span> Subiendo...';
        showOtaStatus('Subiendo archivo...', 'info');
        
        try {
            const res = await fetch('/api/admin/update_code', {
                method: 'POST',
                body: formData
            });
            const data = await res.json();
            
            if (data.success) {
                showOtaStatus(data.message, 'success');
                waitForServerRestart();
            } else {
                showOtaStatus('Error: ' + data.message, 'error');
                btn.disabled = false;
                btn.innerHTML = '<span class="material-icons">upload</span> Aplicar Parche y Reiniciar';
            }
        } catch (err) {
            showOtaStatus('Error de red al subir', 'error');
            btn.disabled = false;
            btn.innerHTML = '<span class="material-icons">upload</span> Aplicar Parche y Reiniciar';
        }
    };
}

if (btnRestartApp) {
    btnRestartApp.onclick = async () => {
        const pwd = prompt("Ingrese contraseña de Administrador para reiniciar la aplicación:");
        if (!pwd) return;
        
        const origText = btnRestartApp.innerHTML;
        btnRestartApp.innerHTML = "Reiniciando...";
        btnRestartApp.disabled = true;
        
        try {
            const res = await fetch('/api/admin/restart', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ password: pwd })
            });
            const data = await res.json();
            if (data.success) {
                waitForServerRestart();
            } else {
                alert("Error al reiniciar: " + data.message);
                btnRestartApp.innerHTML = origText;
                btnRestartApp.disabled = false;
            }
        } catch (e) {
            alert("Error de red");
            btnRestartApp.innerHTML = origText;
            btnRestartApp.disabled = false;
        }
    };
}
// --- MÓDULO BACKUP ---
const btnDownloadBackup = document.getElementById('btnDownloadBackup');
const backupStatusMsg = document.getElementById('backupStatusMsg');

if (btnDownloadBackup) {
    btnDownloadBackup.onclick = async () => {
        btnDownloadBackup.disabled = true;
        btnDownloadBackup.innerHTML = '<span class="material-icons">hourglass_empty</span> Generando Respaldo...';
        
        try {
            // El endpoint devuelve el archivo zip directo o un json de error
            const res = await fetch('/api/admin/backup/download');
            
            if (res.ok && res.headers.get('content-type') === 'application/zip') {
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                // Extraer el nombre del header o poner uno por defecto
                const disposition = res.headers.get('Content-Disposition');
                let filename = 'backup_aos.zip';
                if (disposition && disposition.indexOf('attachment') !== -1) {
                    const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                    const matches = filenameRegex.exec(disposition);
                    if (matches != null && matches[1]) { 
                        filename = matches[1].replace(/['"]/g, '');
                    }
                }
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                
                if (backupStatusMsg) {
                    backupStatusMsg.style.display = 'block';
                    backupStatusMsg.style.color = '#34c759';
                    backupStatusMsg.textContent = '¡Respaldo descargado exitosamente!';
                }
            } else {
                const data = await res.json();
                alert("Error generando respaldo: " + (data.message || 'Error desconocido'));
            }
        } catch (e) {
            console.error(e);
            alert("Error de red al intentar descargar el respaldo.");
        } finally {
            btnDownloadBackup.disabled = false;
            btnDownloadBackup.innerHTML = '<span class="material-icons">download</span> Descargar Copia de Seguridad (.zip)';
        }
    };
}

// --- DIRECTORIO IRSA LOGIC ---
let directorioData = [];

async function loadDirectorioIrsa() {
    const tbody = document.getElementById('bodyDirectorio');
    if (!tbody) return;
    
    tbody.innerHTML = '<tr><td colspan="5" class="empty-state">Cargando directorio...</td></tr>';
    try {
        const response = await fetch('/api/dashboard/directorio_irsa');
        const data = await response.json();
        
        if (data.success) {
            directorioData = data.data;
            renderDirectorio();
        } else {
            tbody.innerHTML = `<tr><td colspan="5" class="empty-state error">Error: ${data.message}</td></tr>`;
        }
    } catch (error) {
        console.error('Error cargando directorio:', error);
        tbody.innerHTML = '<tr><td colspan="5" class="empty-state error">Error de conexión.</td></tr>';
    }
}

function renderDirectorio() {
    const tbody = document.getElementById('bodyDirectorio');
    const filterText = (document.getElementById('filterDirectorio')?.value || '').toLowerCase();
    const filterTipo = document.getElementById('filterDirectorioTipo')?.value || 'TODOS';
    
    if (!tbody) return;

    let filtered = directorioData.filter(item => {
        const matchesText = item.dni.toLowerCase().includes(filterText) || item.nombre.toLowerCase().includes(filterText) || item.empresa.toLowerCase().includes(filterText);
        const matchesTipo = filterTipo === 'TODOS' || item.tipo === filterTipo;
        return matchesText && matchesTipo;
    });

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="empty-state">No se encontraron registros.</td></tr>';
        return;
    }

    filtered.sort((a, b) => a.nombre.localeCompare(b.nombre));

    const hoy = new Date();
    hoy.setHours(0,0,0,0);

    let html = '';
    filtered.forEach(item => {
        let vencimientoClass = '';
        if (item.vence) {
            const parts = item.vence.split('/');
            if (parts.length === 3) {
                const fVence = new Date(parts[2], parts[1]-1, parts[0]);
                const umbral = new Date(hoy);
                if (window.diasAlertaVencimiento) {
                    umbral.setDate(umbral.getDate() + window.diasAlertaVencimiento);
                }

                if (fVence < hoy) {
                    vencimientoClass = 'vencido-text';
                } else if (fVence <= umbral && window.diasAlertaVencimiento > 0) {
                    vencimientoClass = 'por-vencer-text';
                }
            }
        }
        
        html += `
            <tr>
                <td>${item.dni}</td>
                <td>${item.nombre}</td>
                <td>${item.empresa}</td>
                <td class="${vencimientoClass}">${item.vence || 'N/A'}</td>
                <td><span class="status-badge ${item.tipo === 'FAO' ? 'status-ok' : 'status-denied'}">${item.tipo}</span></td>
                <td>${item.numero || '-'}</td>
            </tr>
        `;
    });
    tbody.innerHTML = html;
}

let tramitesData = [];

async function loadTramitesIRSA() {
    try {
        const querySnapshot = await getDocs(collection(db, "tramites_irsa"));
        tramitesData = [];
        querySnapshot.forEach((doc) => {
            tramitesData.push(doc.data());
        });
        renderTramites();
    } catch(e) {
        console.error("Error loading tramites IRSA from Firestore", e);
    }
}

function renderTramites() {
    const tbody = document.getElementById('bodyTramites');
    const filterText = (document.getElementById('filterTramites')?.value || '').toLowerCase();
    const filterEstado = document.getElementById('filterTramitesEstado')?.value || 'TODOS';
    
    if (!tbody) return;

    let filtered = tramitesData.filter(item => {
        const matchesText = item.empresa.toLowerCase().includes(filterText) || item.id.toString().includes(filterText);
        const matchesEstado = filterEstado === 'TODOS' || item.estado_codigo === filterEstado;
        return matchesText && matchesEstado;
    });

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="empty-state">No se encontraron trámites.</td></tr>';
        return;
    }

    filtered.sort((a, b) => b.id - a.id);

    let html = '';
    filtered.forEach(item => {
        let badgeClass = 'status-denied';
        let estadoStr = 'Desconocido';
        if (item.estado_codigo === '6') { badgeClass = 'status-ok'; estadoStr = 'Aprobado'; }
        else if (item.estado_codigo === '4') { badgeClass = 'status-ok'; estadoStr = 'Aprobado Parcial'; }
        else if (item.estado_codigo === '3') { badgeClass = ''; estadoStr = 'Enviado'; }
        else if (item.estado_codigo === '5') { badgeClass = 'status-denied'; estadoStr = 'Requiere Cambios'; }
        
        let personalHtml = '';
        if (item.personal && item.personal.length > 0) {
            personalHtml = `
                <tr class="accordion-content" id="accordion-${item.id}" style="display: none; background-color: var(--dash-surface-hover);">
                    <td colspan="7" style="padding: 1rem;">
                        <div style="font-weight: 500; margin-bottom: 0.5rem; color: var(--dash-primary);">Personas Autorizadas:</div>
                        <ul style="list-style-type: disc; padding-left: 1.5rem; margin: 0; color: var(--dash-text-muted);">
                            ${item.personal.map(p => `<li>${p.nombre || ''} ${p.apellido || ''} - DNI: ${p.numeroDocumento || ''} - Estado: ${p.activo ? 'Activo' : 'Inactivo'}</li>`).join('')}
                        </ul>
                    </td>
                </tr>
            `;
        } else {
            personalHtml = `
                <tr class="accordion-content" id="accordion-${item.id}" style="display: none; background-color: var(--dash-surface-hover);">
                    <td colspan="7" style="padding: 1rem; color: var(--dash-text-muted); font-style: italic;">
                        No hay personas registradas en este trámite.
                    </td>
                </tr>
            `;
        }

        html += `
            <tr class="accordion-header" onclick="toggleAccordion('accordion-${item.id}')" style="cursor: pointer;">
                <td><span class="status-badge ${item.tipo === 'FAO' ? 'status-ok' : 'status-denied'}">${item.tipo}</span></td>
                <td style="font-weight: 500;">#${item.id}</td>
                <td>${item.empresa || '-'}</td>
                <td>${item.fecha_inicio || '-'}</td>
                <td>${item.fecha_fin || '-'}</td>
                <td><span class="status-badge ${badgeClass}">${estadoStr}</span></td>
                <td style="text-align:center;"><span class="material-icons" style="font-size: 1.5rem; vertical-align: middle; color: var(--dash-primary);">expand_more</span></td>
            </tr>
            ${personalHtml}
        `;
    });
    tbody.innerHTML = html;
}

window.toggleAccordion = function(id) {
    const el = document.getElementById(id);
    if (el) {
        el.style.display = el.style.display === 'none' ? 'table-row' : 'none';
    }
};

document.addEventListener('DOMContentLoaded', () => {
    const filterInput = document.getElementById('filterDirectorio');
    const filterSelect = document.getElementById('filterDirectorioTipo');
    const filterTramites = document.getElementById('filterTramites');
    const filterTramitesEstado = document.getElementById('filterTramitesEstado');

    if (filterInput) {
        filterInput.addEventListener('input', renderDirectorio);
        filterInput.addEventListener('input', renderNominas);
    }
    if (filterSelect) {
        filterSelect.addEventListener('change', renderDirectorio);
        filterSelect.addEventListener('change', renderNominas);
    }
    if (filterTramites) filterTramites.addEventListener('input', renderTramites);
    if (filterTramitesEstado) filterTramitesEstado.addEventListener('change', renderTramites);

    const tabs = document.querySelectorAll('.dash-tab');
    tabs.forEach(tab => {
        tab.addEventListener('click', (e) => {
            const target = e.currentTarget.getAttribute('data-tab');
            if (target === 'tramites') {
                loadTramitesIRSA();
            }
        });
    });
});

// --- MOTOR IOT: Enviar comandos al Obrero Python ---
async function sendWorkerCommand(actionName, btnElement, origText, statusElement) {
    btnElement.disabled = true;
    btnElement.innerHTML = '<span class="material-icons">hourglass_empty</span> Ejecutando...';
    statusElement.textContent = 'Enviando comando a la garita...';
    statusElement.style.color = 'var(--dash-text-muted)';
    
    try {
        const cmdRef = await addDoc(collection(db, 'system_commands'), {
            action: actionName,
            status: 'PENDING',
            createdAt: serverTimestamp()
        });
        
        // Escuchar la respuesta del worker
        const unsubscribe = onSnapshot(cmdRef, (docSnap) => {
            const data = docSnap.data();
            if (data.status === 'IN_PROGRESS') {
                statusElement.textContent = 'Obrero ejecutando tarea...';
                statusElement.style.color = 'var(--dash-primary)';
            } else if (data.status === 'DONE') {
                statusElement.textContent = `Éxito: ${data.result}`;
                statusElement.style.color = 'var(--dash-success)';
                btnElement.disabled = false;
                btnElement.innerHTML = origText;
                unsubscribe();
            } else if (data.status === 'ERROR') {
                statusElement.textContent = `Error: ${data.result}`;
                statusElement.style.color = 'var(--dash-danger)';
                btnElement.disabled = false;
                btnElement.innerHTML = origText;
                unsubscribe();
            }
        });
    } catch (e) {
        statusElement.textContent = 'Error de conexión con la nube.';
        statusElement.style.color = 'var(--dash-danger)';
        btnElement.disabled = false;
        btnElement.innerHTML = origText;
    }
}

document.addEventListener('DOMContentLoaded', () => {
    const btnIrsa = document.getElementById('btnTriggerIrsa');
    const btnOta = document.getElementById('btnTriggerOta');
    const statusMsg = document.getElementById('workerStatusMsg');
    
    if (btnIrsa) {
        btnIrsa.addEventListener('click', () => {
            const orig = '<span class="material-icons">sync</span> Extraer Nóminas IRSA Ahora';
            sendWorkerCommand('SYNC_IRSA', btnIrsa, orig, statusMsg);
        });
    }
    
    if (btnOta) {
        btnOta.addEventListener('click', () => {
            if(!confirm("¿Actualizar el código fuente de la garita principal? Se reiniciará el motor.")) return;
            const orig = '<span class="material-icons">system_update_alt</span> Actualizar Código de Garita (OTA)';
            sendWorkerCommand('UPDATE_OTA', btnOta, orig, statusMsg);
        });
    }
    
    const btnManual = document.getElementById('btnManualSync');
    const manualInput = document.getElementById('manualFaoInput');
    if (btnManual && manualInput) {
        btnManual.addEventListener('click', () => {
            const val = manualInput.value.trim();
            if(!val) {
                alert("Por favor ingrese el número de FAO.");
                return;
            }
            const orig = '<span class="material-icons">download</span> Descarga Forzada';
            // We need to send custom data, so we use a modified sendWorkerCommand or direct fetch
            btnManual.disabled = true;
            btnManual.innerHTML = 'Forzando...';
            statusMsg.textContent = 'Enviando orden a la nube...';
            statusMsg.style.color = 'var(--dash-text-muted)';
            
            db.collection('system_commands').add({
                action: 'FORCE_SYNC_FAO',
                fao_id: val,
                status: 'PENDING',
                createdAt: firebase.firestore.FieldValue.serverTimestamp()
            }).then(docRef => {
                statusMsg.textContent = 'Orden en cola. El Obrero procesará la descarga en unos segundos.';
                statusMsg.style.color = 'var(--dash-primary)';
                
                const unsub = db.collection('system_commands').doc(docRef.id).onSnapshot(doc => {
                    const data = doc.data();
                    if(data.status === 'DONE' || data.status === 'ERROR') {
                        btnManual.disabled = false;
                        btnManual.innerHTML = orig;
                        statusMsg.textContent = data.result || 'Comando finalizado.';
                        statusMsg.style.color = data.status === 'DONE' ? 'var(--dash-success)' : 'var(--dash-danger)';
                        manualInput.value = '';
                        unsub();
                    } else if (data.status === 'IN_PROGRESS') {
                        statusMsg.textContent = 'Descargando desde IRSA...';
                    }
                });
            }).catch(err => {
                btnManual.disabled = false;
                btnManual.innerHTML = orig;
                statusMsg.textContent = 'Error: ' + err.message;
                statusMsg.style.color = 'var(--dash-danger)';
            });
        });
    }
});
// --- CONFIG: Puestos Operativos ---
const formPuestos = document.getElementById('formPuestos');
const listaPuestos = document.getElementById('listaPuestos');

async function loadPuestos() {
    if (!listaPuestos) return;
    try {
        const res = await fetch('/api/config/puestos');
        const data = await res.json();
        if (data.success) {
            if (data.data.length === 0) {
                listaPuestos.innerHTML = '<li style="color: var(--dash-text-muted); text-align: center; font-size: 0.85rem;">No hay puestos registrados</li>';
            } else {
                listaPuestos.innerHTML = data.data.map(p => 
                    '<li style="display: flex; justify-content: space-between; align-items: center; padding: 0.5rem; background: var(--dash-surface); border-radius: var(--dash-radius-sm);">' +
                        '<span>' + p.nombre + '</span>' +
                        '<button type="button" class="dash-btn dash-btn-danger" style="padding: 0.25rem 0.5rem; font-size: 0.75rem;" onclick="deletePuesto(' + p.id + ')">' +
                            '<span class="material-icons" style="font-size: 1rem;">delete</span>' +
                        '</button>' +
                    '</li>'
                ).join('');
            }
        }
    } catch (e) {
        console.error('Error loading puestos:', e);
    }
}

if (formPuestos) {
    formPuestos.onsubmit = async (e) => {
        e.preventDefault();
        const input = document.getElementById('nuevoPuestoNombre');
        const nombre = input.value.trim();
        if (!nombre) return;
        
        try {
            const res = await fetch('/api/config/puestos', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ nombre })
            });
            const data = await res.json();
            if (data.success) {
                input.value = '';
                loadPuestos();
            } else {
                alert('Error al crear puesto: ' + data.message);
            }
        } catch (e) {
            alert('Error de red');
        }
    };
}

window.deletePuesto = async function(id) {
    if (!confirm('¿Seguro que desea eliminar este puesto?')) return;
    try {
        const res = await fetch('/api/config/puestos/' + id, { method: 'DELETE' });
        const data = await res.json();
        if (data.success) {
            loadPuestos();
        } else {
            alert('Error al eliminar puesto: ' + data.message);
        }
    } catch (e) {
        alert('Error de red');
    }
};

// Cargar puestos cuando se abre la tab de config
(() => {
    // Si estamos en la página del dashboard, hook al cambio de tab
    const configTabBtn = document.querySelector('button[data-tab="config"]');
    if (configTabBtn) {
        configTabBtn.addEventListener('click', loadPuestos);
    }
    // Opcionalmente cargarlos de entrada
    if(document.getElementById('panel-config') && !document.getElementById('panel-config').classList.contains('hidden')){
        loadPuestos();
    }
})();


// --- ABM Puestos Fisicos ---
async function cargarPuestosFisicos() {
    const lista = document.getElementById('listaPuestosFisicos');
    if (!lista) return;
    try {
        const res = await fetch('/api/config/puestos_fisicos');
        const data = await res.json();
        if (data.success) {
            lista.innerHTML = '';
            data.data.forEach(p => {
                const li = document.createElement('li');
                li.style.display = 'flex';
                li.style.justifyContent = 'space-between';
                li.style.padding = '0.5rem';
                li.style.background = 'rgba(255,255,255,0.05)';
                li.style.borderRadius = '4px';
                
                li.innerHTML = `
                    <span>${p.nombre} <small style="color: #aaa;">(${p.sector})</small> <small style="color: #888;">[${p.hora_inicio || ''} - ${p.hora_fin || ''}]</small></span>
                    <button class="dash-btn" style="background: transparent; color: #ff3b30; padding: 2px 5px;" onclick="eliminarPuestoFisico(${p.id})">
                        <span class="material-icons" style="font-size: 1.1rem;">delete</span>
                    </button>
                `;
                lista.appendChild(li);
            });
        }
    } catch (e) { console.error('Error', e); }
}

const formPuestosFisicos = document.getElementById('formPuestosFisicos');
if (formPuestosFisicos) {
    formPuestosFisicos.addEventListener('submit', async (e) => {
        e.preventDefault();
        const payload = {
            nombre: document.getElementById('nuevoPuestoFisicoNombre').value,
            sector: document.getElementById('nuevoPuestoFisicoSector').value,
            hora_inicio: document.getElementById('nuevoPuestoFisicoInicio').value,
            hora_fin: document.getElementById('nuevoPuestoFisicoFin').value
        };
        try {
            const res = await fetch('/api/config/puestos_fisicos', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(payload)
            });
            const data = await res.json();
            if (data.success) {
                document.getElementById('nuevoPuestoFisicoNombre').value = '';
                document.getElementById('nuevoPuestoFisicoInicio').value = '';
                document.getElementById('nuevoPuestoFisicoFin').value = '';
                cargarPuestosFisicos();
            } else { alert('Error: ' + data.message); }
        } catch (e) { alert('Error de conexión'); }
    });
}

async function eliminarPuestoFisico(id) {
    if (!confirm('¿Eliminar puesto físico?')) return;
    try {
        const res = await fetch(`/api/config/puestos_fisicos/${id}`, {method: 'DELETE'});
        const data = await res.json();
        if (data.success) cargarPuestosFisicos();
    } catch (e) { alert('Error'); }
}


// --- ABM Empresas ---
async function cargarEmpresas() {
    const lista = document.getElementById('listaEmpresas');
    if (!lista) return;
    try {
        const res = await fetch('/api/config/empresas');
        const data = await res.json();
        if (data.success) {
            lista.innerHTML = '';
            data.data.forEach(e => {
                const li = document.createElement('li');
                li.style.display = 'flex';
                li.style.justifyContent = 'space-between';
                li.style.padding = '0.5rem';
                li.style.background = 'rgba(255,255,255,0.05)';
                li.style.borderRadius = '4px';
                
                li.innerHTML = `
                    <span>${e.nombre}</span>
                    <button class="dash-btn" style="background: transparent; color: #ff3b30; padding: 2px 5px;" onclick="eliminarEmpresa(${e.id})">
                        <span class="material-icons" style="font-size: 1.1rem;">delete</span>
                    </button>
                `;
                lista.appendChild(li);
            });
        }
    } catch (e) { console.error('Error', e); }
}

const formEmpresas = document.getElementById('formEmpresas');
if (formEmpresas) {
    formEmpresas.addEventListener('submit', async (e) => {
        e.preventDefault();
        try {
            const res = await fetch('/api/config/empresas', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ nombre: document.getElementById('nuevaEmpresaNombre').value })
            });
            const data = await res.json();
            if (data.success) {
                document.getElementById('nuevaEmpresaNombre').value = '';
                cargarEmpresas();
            } else { alert('Error: ' + data.message); }
        } catch (e) { alert('Error de conexión'); }
    });
}

async function eliminarEmpresa(id) {
    if (!confirm('¿Eliminar empresa?')) return;
    try {
        const res = await fetch(`/api/config/empresas/${id}`, {method: 'DELETE'});
        const data = await res.json();
        if (data.success) cargarEmpresas();
    } catch (e) { alert('Error'); }
}

document.addEventListener('DOMContentLoaded', () => {
    cargarPuestosFisicos();
    cargarEmpresas();
});

// --- PARSEADOR DE EXCEPCIONES ---
document.addEventListener('DOMContentLoaded', () => {
    const btnPreview = document.getElementById('btnPreviewNomina');
    const btnSave = document.getElementById('btnSaveNomina');
    if(!btnPreview || !btnSave) return;

    btnPreview.addEventListener('click', () => {
        const texto = document.getElementById('nominaTexto').value;
        const status = document.getElementById('nominaParseStatus');
        
        if(!texto.trim()) {
            status.textContent = 'El texto está vacío';
            status.style.color = 'var(--dash-danger)';
            return;
        }

        const lineas = texto.split('\n');
        const personas = [];
        const regexes = [
            /^(?:C\.?U\.?I\.?L\.?[:\s]*)(?<cuil>\d{11})\s+(?<nombre_completo>.*)/i,
            /^(?<cuil>\d{2}-\d{7,8}-\d{1})\s+(?<nombre_completo>.*)/i,
            /^(?<cuil>\d{11})\s+\d+\s+(?<nombre_completo>.*)/i,
            /^(?<cuil>\d{11})\s+(?<nombre_completo>.*)/i,
            /^(?<dni>\d{7,8})\s+(?<nombre_completo>.*)/i,
            /^(?<nombre_completo>.*?)\s+(?<cuil>\d{2}-\d{7,8}-\d{1})$/i
        ];

        lineas.forEach(line => {
            line = line.trim();
            if(!line) return;
            if(line.toUpperCase().includes('CUIL') || line.toUpperCase().includes('APELLIDO')) return;

            for(let r of regexes) {
                const match = line.match(r);
                if(match) {
                    const idNumber = (match.groups.cuil || match.groups.dni || '').replace(/\D/g, '');
                    const dni = idNumber.length === 11 ? idNumber.substring(2, 10) : idNumber;
                    
                    const partesNombre = match.groups.nombre_completo.trim().split(' ').filter(p => p);
                    let apellido = '';
                    let nombre = '';
                    if(partesNombre.length === 1) {
                        apellido = partesNombre[0];
                    } else if(partesNombre.length === 2) {
                        apellido = partesNombre[0];
                        nombre = partesNombre[1];
                    } else {
                        apellido = partesNombre.slice(0, 2).join(' ');
                        nombre = partesNombre.slice(2).join(' ');
                    }

                    personas.push({dni, apellido, nombre});
                    break;
                }
            }
        });

        // FALLBACK: Si no encontró a nadie línea por línea, intenta extraer con una expresión regular global para textos sin saltos de línea
        if (personas.length === 0) {
            // Regex para el formato de Prevención ART y similares pegados como un solo bloque:
            // "20404903706 1AVALOS ABEL ISAIAS Régimen General 20364052619 2AVALOS EDUARDO DANIEL Régimen General"
            // Atrapa el CUIL (11 dígitos opcionalmente con guiones), descarta el número pegado al apellido, atrapa el nombre, y frena antes de la condición u otro CUIL.
            const globalRegex = /(?:C\.?U\.?I\.?L\.?[:\s]*)?(\d{2}-?\d{7,8}-?\d{1}|\d{11})\s+(?:\d+)?([A-ZÁÉÍÓÚÑa-záéíóúñ\s]+?)(?=\s+Régimen General|\s+Condición|\s+\d{2}-?\d{7,8}-?\d{1}|\s+\d{11}|$)/gi;
            let match;
            while ((match = globalRegex.exec(texto)) !== null) {
                const idNumber = match[1].replace(/\D/g, '');
                const dni = idNumber.length === 11 ? idNumber.substring(2, 10) : idNumber;
                
                const partesNombre = match[2].trim().split(' ').filter(p => p);
                let apellido = '';
                let nombre = '';
                if(partesNombre.length === 1) {
                    apellido = partesNombre[0];
                } else if(partesNombre.length === 2) {
                    apellido = partesNombre[0];
                    nombre = partesNombre[1];
                } else {
                    apellido = partesNombre.slice(0, 2).join(' ');
                    nombre = partesNombre.slice(2).join(' ');
                }
                
                // Evitar atrapar encabezados sueltos como "DETALLE DEL PERSONAL" si se colaron
                if(apellido.toUpperCase() !== "DETALLE" && apellido.toUpperCase() !== "EMPRESA") {
                    personas.push({dni, apellido, nombre});
                }
            }
        }

        if(personas.length === 0) {
            status.textContent = 'No se encontró a nadie válido en el texto';
            status.style.color = 'var(--dash-danger)';
            return;
        }

        document.getElementById('previewNominaContainer').style.display = 'block';
        status.textContent = `Se encontraron ${personas.length} personas.`;
        status.style.color = 'var(--dash-success)';

        const tbody = document.getElementById('previewNominaBody');
        tbody.innerHTML = '';
        personas.forEach((p, i) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><input type='text' class='dash-input' style='padding:5px' value='${p.dni}' id='pdni_${i}'></td>
                <td><input type='text' class='dash-input' style='padding:5px' value='${p.apellido}' id='pape_${i}'></td>
                <td><input type='text' class='dash-input' style='padding:5px' value='${p.nombre}' id='pnom_${i}'></td>
                <td><button class='dash-btn dash-btn-danger' onclick='this.parentElement.parentElement.remove()'>Quitar</button></td>
            `;
            tbody.appendChild(tr);
        });
    });

    btnSave.addEventListener('click', async () => {
        const empresa = document.getElementById('nominaEmpresa').value.trim();
        const vigDesde = document.getElementById('nominaVigenciaDesde').value;
        const vigHasta = document.getElementById('nominaVigenciaHasta').value;
        const status = document.getElementById('nominaSaveStatus');

        if(!empresa || !vigDesde || !vigHasta) {
            status.textContent = 'Faltan completar campos de empresa o fechas.';
            status.style.color = 'var(--dash-danger)';
            return;
        }

        const tbody = document.getElementById('previewNominaBody');
        const rows = tbody.querySelectorAll('tr');
        const nomina = [];
        
        rows.forEach(tr => {
            const inputs = tr.querySelectorAll('input');
            const dni = inputs[0].value.trim();
            const apellido = inputs[1].value.trim();
            const nombre = inputs[2].value.trim();
            if(dni) {
                nomina.push({dni, apellido, nombre});
            }
        });

        if(nomina.length === 0) {
            status.textContent = 'La lista está vacía.';
            status.style.color = 'var(--dash-danger)';
            return;
        }

        btnSave.disabled = true;
        btnSave.innerHTML = 'Guardando...';

        try {
            // Guardado Offline First directamente en Firebase
            const batchPromises = nomina.map(async p => {
                // Primero buscamos si hay previos con este DNI para borrarlos
                const q = query(collection(db, 'nominas'), where('dni', '==', p.dni), where('empresa', '==', empresa));
                const snap = await getDocs(q);
                const deletePromises = snap.docs.map(d => deleteDoc(doc(db, 'nominas', d.id)));
                await Promise.all(deletePromises);

                // Insertamos el nuevo
                const nombreCompleto = p.nombre ? p.nombre + ' ' + p.apellido : p.apellido;
                return addDoc(collection(db, 'nominas'), {
                    dni: p.dni,
                    nombre: nombreCompleto,
                    empresa: empresa,
                    fecha_vencimiento: vigHasta, // Mapeado correctamente para la UI
                    activo: true,
                    categoria: 'Excepción/ART',
                    createdAt: serverTimestamp()
                });
            });

            await Promise.all(batchPromises);

            // También enviamos silenciosamente la orden al Obrero para que actualice la base local de los molinetes
            addDoc(collection(db, 'system_commands'), {
                action: 'SAVE_EXCEPCIONES',
                empresa: empresa,
                vigencia_desde: vigDesde,
                vigencia_hasta: vigHasta,
                nomina: nomina,
                status: 'PENDING',
                createdAt: serverTimestamp()
            });
            
            btnSave.disabled = false;
            btnSave.innerHTML = '<span class="material-icons">save</span> Guardar y Aprobar Excepción';
            status.textContent = 'Guardado localmente y sincronizando con la nube...';
            status.style.color = 'var(--dash-success)';
            document.getElementById('nominaTexto').value = '';
            document.getElementById('previewNominaBody').innerHTML = '';
            document.getElementById('previewNominaContainer').style.display = 'none';

        } catch (e) {
            btnSave.disabled = false;
            btnSave.innerHTML = '<span class="material-icons">save</span> Guardar y Aprobar Excepción';
            status.textContent = 'Error: ' + e.message;
            status.style.color = 'var(--dash-danger)';
            console.error(e);
        }
    });
});
