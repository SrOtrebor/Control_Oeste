/* =============================================
   DASHBOARD JS - Control de Acceso AOS (FIREBASE)
   ============================================= */
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getAuth, onAuthStateChanged, signOut, createUserWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-auth.js";
import { getFirestore, collection, query, where, onSnapshot, getDocs, doc, getDoc, setDoc, addDoc, serverTimestamp, deleteDoc, updateDoc } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-firestore.js";

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
let currentUserRole = null;
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
                currentUserRole = data.rol;

                // Solo supervisores y super_admin pueden ingresar al Dashboard
                if (data.rol !== 'supervisor' && data.rol !== 'super_admin') {
                    alert("Acceso denegado. No tienes permisos para acceder al dashboard.");
                    window.location.href = '/scan.html';
                    return;
                }
                
                // CONTROL DE PRIVILEGIOS EXCLUSIVOS:
                // Solo Roberto ('super_admin') puede ver y gestionar "Usuarios" y "Shoppings"
                const isSuperAdmin = (data.rol === 'super_admin');
                document.querySelectorAll('.super-admin-only').forEach(el => {
                    el.style.display = isSuperAdmin ? '' : 'none';
                });

                if (isSuperAdmin) {
                    initCentros();
                    initUsuarios();
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
    initCierresJornada();
    initExcepciones();
    initConfig();
    loadDashboardData();
});

// =============================================
// TABS
// =============================================
function initTabs() {
    document.querySelectorAll('.dash-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            const target = tab.dataset.tab;
            if ((target === 'usuarios' || target === 'shoppings') && currentUserRole !== 'super_admin') {
                return;
            }
            document.querySelectorAll('.dash-tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.dash-panel').forEach(p => p.classList.add('hidden'));
            tab.classList.add('active');
            const panelId = 'panel-' + target;
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
    
    // Escuchar la coleccion estado_adentro en tiempo real para poblar la grilla y el contador
    const estadoAdentroRef = collection(db, 'estado_adentro');
    onSnapshot(estadoAdentroRef, (snapshot) => {
        const statAdentro = document.getElementById('statAdentro');
        if (statAdentro) {
            statAdentro.textContent = snapshot.size;
        }

        let personas = [];
        snapshot.forEach((doc) => {
            personas.push(doc.data());
        });
        
        // Ordenar alfabeticamente por nombre
        personas.sort((a, b) => (a.nombre || '').localeCompare(b.nombre || ''));
        
        renderPersonasAdentro(personas);
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
                <span class="persona-dni" title="DNI: ${p.dni || p.DNI || 'Desconocido'}">${p.nombre || p.Nombre || p.DNI || p.dni || 'Sin nombre'}</span>
                <span class="persona-tipo">${p.tipo_permiso}</span>
            </div>
        </div>`;
    }).join('');
}

// =============================================
// HISTORIAL DE CIERRES DE JORNADA (AUDITORÍA)
// =============================================
let cierresJornadaCache = [];

function initCierresJornada() {
    const q = query(collection(db, 'cierres_jornada'));
    onSnapshot(q, (snapshot) => {
        cierresJornadaCache = [];
        snapshot.forEach(docSnap => {
            const data = docSnap.data();
            data.id = docSnap.id;
            cierresJornadaCache.push(data);
        });

        // Ordenar descendente (más recientes primero)
        cierresJornadaCache.sort((a, b) => {
            const fechaA = (a.fecha || '') + ' ' + (a.hora || '');
            const fechaB = (b.fecha || '') + ' ' + (b.hora || '');
            return fechaB.localeCompare(fechaA);
        });

        renderCierresJornada();
    });

    // Eventos de modal detalle
    const closeModal = document.getElementById('closeModalDetalleCierre');
    const btnCerrarModal = document.getElementById('btnCerrarModalDetalle');
    const modal = document.getElementById('modalDetalleCierre');

    const ocultarModal = () => {
        if (modal) modal.style.display = 'none';
    };

    if (closeModal) closeModal.onclick = ocultarModal;
    if (btnCerrarModal) btnCerrarModal.onclick = ocultarModal;
    window.addEventListener('click', (e) => {
        if (e.target === modal) ocultarModal();
    });
}

function renderCierresJornada() {
    const tbody = document.getElementById('bodyCierresJornada');
    if (!tbody) return;

    if (cierresJornadaCache.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-state">No hay registros de cierre de jornada aún.</td></tr>';
        return;
    }

    tbody.innerHTML = cierresJornadaCache.map(c => {
        const cant = c.total_quedaron_adentro || (c.personas_adentro ? c.personas_adentro.length : 0);
        const cantBadge = cant > 0 
            ? `<span style="background:#ef4444; color:white; padding:3px 8px; border-radius:12px; font-weight:600; font-size:0.85rem;">${cant} personas</span>`
            : `<span style="background:#10b981; color:white; padding:3px 8px; border-radius:12px; font-weight:600; font-size:0.85rem;">0 personas (Limpio)</span>`;

        return `<tr>
            <td><strong>${c.fecha || '-'}</strong></td>
            <td>${c.hora || '-'}</td>
            <td>${c.operario_email || 'Garita'}</td>
            <td>${cantBadge}</td>
            <td style="max-width:250px; font-style:italic; color:var(--dash-text-muted);">${c.observacion || '-'}</td>
            <td>
                <button class="dash-btn dash-btn-sm" onclick="window.verDetalleCierre('${c.id}')" style="background:var(--dash-primary); color:white; padding:4px 10px; font-size:0.8rem;">
                    <span class="material-icons" style="font-size:16px;">visibility</span> Auditar Personas
                </button>
            </td>
        </tr>`;
    }).join('');
}

window.verDetalleCierre = function(cierreId) {
    const cierre = cierresJornadaCache.find(x => x.id === cierreId);
    if (!cierre) return;

    const modal = document.getElementById('modalDetalleCierre');
    const subtitulo = document.getElementById('detalleCierreSubtitulo');
    const tbody = document.getElementById('bodyDetalleCierrePersonas');

    if (subtitulo) {
        subtitulo.innerHTML = `Cierre del <strong>${cierre.fecha} a las ${cierre.hora}</strong> por <strong>${cierre.operario_email}</strong>. Observación: <em>"${cierre.observacion || 'Sin observaciones'}"</em>`;
    }

    const personas = cierre.personas_adentro || [];
    if (personas.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="empty-state" style="text-align:center; padding:2rem;">No había personas dentro al momento del cierre. Todo en orden.</td></tr>';
    } else {
        tbody.innerHTML = personas.map(p => `
            <tr>
                <td><strong>${p.dni || '-'}</strong></td>
                <td>${p.nombre || '-'}</td>
                <td>${p.hora_ingreso || '-'}</td>
                <td><span class="persona-tipo" style="font-size:0.8rem;">${p.tipo_permiso || '-'}</span></td>
                <td>${p.local || '-'}</td>
            </tr>
        `).join('');
    }

    if (modal) modal.style.display = 'flex';
};

// =============================================
// GESTIÓN DE CENTROS COMERCIALES (SHOPPINGS)
// =============================================
let centrosCache = [];

function initCentros() {
    const bodyCentros = document.getElementById('bodyCentros');
    const badgeTotal = document.getElementById('badgeTotalCentros');
    const selectCentroUsuario = document.getElementById('newUserCentro');
    if (!bodyCentros) return;

    onSnapshot(collection(db, "centros"), (snapshot) => {
        centrosCache = [];
        if (snapshot.empty) {
            bodyCentros.innerHTML = '<tr><td colspan="6" class="empty-state">No hay centros comerciales registrados.</td></tr>';
            if (badgeTotal) badgeTotal.textContent = '0 sedes';
            if (selectCentroUsuario) {
                selectCentroUsuario.innerHTML = '<option value="al_oeste">Al Oeste Shopping</option>';
            }
            return;
        }

        if (badgeTotal) badgeTotal.textContent = `${snapshot.size} sede${snapshot.size === 1 ? '' : 's'}`;

        let html = '';
        let optionsHtml = '';

        snapshot.forEach(docSnap => {
            const c = docSnap.data();
            const id = docSnap.id;
            centrosCache.push({ id, ...c });
            const nombre = c.nombre || id;
            const direccion = c.direccion || '-';
            const irsaId = c.irsa_id || '-';
            const activo = c.activo !== false;

            const estadoBadge = activo 
                ? '<span class="badge badge-success"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:currentColor;margin-right:4px;"></span>Activo</span>'
                : '<span class="badge badge-danger"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:currentColor;margin-right:4px;"></span>Inactivo</span>';

            const toggleBtnText = activo ? 'Desactivar' : 'Activar';

            html += `
                <tr data-id="${id}">
                    <td><strong>${escapeHtml(nombre)}</strong></td>
                    <td><code>${escapeHtml(id)}</code></td>
                    <td>${escapeHtml(direccion)}</td>
                    <td>${escapeHtml(irsaId)}</td>
                    <td>${estadoBadge}</td>
                    <td style="text-align: right;">
                        <div style="display: inline-flex; gap: 6px; justify-content: flex-end;">
                            <button class="dash-btn" style="${activo ? 'background: rgba(255,255,255,0.08); border: 1px solid var(--dash-border); color: var(--dash-text);' : 'background: rgba(52, 211, 153, 0.15); color: var(--dash-success); border: 1px solid rgba(52, 211, 153, 0.3);'} font-size: 0.75rem; padding: 4px 8px;" onclick="window.toggleEstadoCentro('${id}', ${activo})">
                                ${toggleBtnText}
                            </button>
                            <button class="dash-btn" style="background: rgba(248, 113, 113, 0.15); color: var(--dash-danger); border: 1px solid rgba(248, 113, 113, 0.3); font-size: 0.75rem; padding: 4px 8px;" onclick="window.eliminarCentroDoc('${id}', '${escapeHtml(nombre)}')">
                                <span class="material-icons" style="font-size: 1rem;">delete</span>
                            </button>
                        </div>
                    </td>
                </tr>
            `;

            optionsHtml += `<option value="${id}">${escapeHtml(nombre)}</option>`;
        });

        bodyCentros.innerHTML = html;
        if (selectCentroUsuario) {
            selectCentroUsuario.innerHTML = optionsHtml;
        }

        // Si ya hay usuarios cargados, refrescar sus nombres de centro
        if (typeof renderUsuariosConCentros === 'function') {
            renderUsuariosConCentros();
        }
    }, (error) => {
        console.error("Error al cargar centros:", error);
        bodyCentros.innerHTML = `<tr><td colspan="6" class="empty-state" style="color:var(--dash-danger);">Error: ${error.message}</td></tr>`;
    });
}

window.autoGenerarCodigoCentro = function(nombre) {
    const idInput = document.getElementById('newCentroId');
    if (!idInput) return;
    const slug = nombre
        .toLowerCase()
        .trim()
        .replace(/[áàäâ]/g, 'a')
        .replace(/[éèëê]/g, 'e')
        .replace(/[íìïî]/g, 'i')
        .replace(/[óòöô]/g, 'o')
        .replace(/[úùüû]/g, 'u')
        .replace(/ñ/g, 'n')
        .replace(/[^a-z0-9]/g, '_')
        .replace(/_+/g, '_')
        .replace(/^_|_$/g, '');
    idInput.value = slug;
};

window.crearNuevoCentro = async function(event) {
    if (event) event.preventDefault();
    const nombreInput = document.getElementById('newCentroNombre');
    const idInput = document.getElementById('newCentroId');
    const dirInput = document.getElementById('newCentroDireccion');
    const irsaInput = document.getElementById('newCentroIrsaId');
    const msgEl = document.getElementById('msgCrearCentro');
    const btn = document.getElementById('btnCrearCentro');

    const nombre = nombreInput ? nombreInput.value.trim() : '';
    const id = idInput ? idInput.value.trim().toLowerCase().replace(/[^a-z0-9_-]/g, '_') : '';
    const direccion = dirInput ? dirInput.value.trim() : '';
    const irsaId = irsaInput ? irsaInput.value.trim() : '';

    if (!nombre || !id) {
        msgEl.style.color = 'var(--dash-danger)';
        msgEl.textContent = 'El nombre y código del centro son obligatorios.';
        return;
    }

    if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<span class="material-icons rotating">sync</span> Guardando...';
    }
    msgEl.style.color = 'var(--dash-info)';
    msgEl.textContent = 'Guardando centro comercial en la nube...';

    try {
        await setDoc(doc(db, "centros", id), {
            id: id,
            nombre: nombre,
            direccion: direccion,
            irsa_id: irsaId,
            activo: true,
            creado_el: serverTimestamp()
        }, { merge: true });

        msgEl.style.color = 'var(--dash-success)';
        msgEl.textContent = `¡Shopping "${nombre}" registrado con éxito!`;

        if (nombreInput) nombreInput.value = '';
        if (idInput) idInput.value = '';
        if (dirInput) dirInput.value = '';
        if (irsaInput) irsaInput.value = '';

        setTimeout(() => {
            if (msgEl && msgEl.textContent.includes('éxito')) {
                msgEl.textContent = '';
            }
        }, 5000);
    } catch(err) {
        console.error("Error al registrar centro:", err);
        msgEl.style.color = 'var(--dash-danger)';
        msgEl.textContent = 'Error: ' + err.message;
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = '<span class="material-icons">add_business</span> Registrar Shopping';
        }
    }
};

window.toggleEstadoCentro = async function(id, currentActivo) {
    try {
        await updateDoc(doc(db, "centros", id), {
            activo: !currentActivo
        });
    } catch(err) {
        alert("Error al actualizar estado del centro: " + err.message);
    }
};

window.eliminarCentroDoc = async function(id, nombre) {
    if (!confirm(`¿Está seguro de eliminar el shopping "${nombre}"? Esta acción no se puede deshacer.`)) {
        return;
    }
    try {
        await deleteDoc(doc(db, "centros", id));
    } catch(err) {
        alert("Error al eliminar shopping: " + err.message);
    }
};

// =============================================
// GESTIÓN DE USUARIOS
// =============================================
function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

let secondaryAuthInstance = null;
function getSecondaryAuth() {
    if (!secondaryAuthInstance) {
        const secondaryApp = initializeApp(firebaseConfig, "SecondaryAuthApp");
        secondaryAuthInstance = getAuth(secondaryApp);
    }
    return secondaryAuthInstance;
}

let rawUsuariosDocs = [];

function renderUsuariosConCentros() {
    const bodyUsuarios = document.getElementById('bodyUsuarios');
    const badgeTotal = document.getElementById('badgeTotalUsuarios');
    if (!bodyUsuarios) return;

    if (rawUsuariosDocs.length === 0) {
        bodyUsuarios.innerHTML = '<tr><td colspan="6" class="empty-state">No hay usuarios registrados.</td></tr>';
        if (badgeTotal) badgeTotal.textContent = '0 usuarios';
        return;
    }

    if (badgeTotal) badgeTotal.textContent = `${rawUsuariosDocs.length} usuario${rawUsuariosDocs.length === 1 ? '' : 's'}`;

    let html = '';
    rawUsuariosDocs.forEach(docSnap => {
        const u = docSnap.data();
        const uid = docSnap.id;
        const email = u.email || 'Sin correo';
        const nombre = u.nombre || '-';
        const activo = u.activo !== false;
        const rol = u.rol || 'acceso_principal';
        const centroId = u.centro_id || 'al_oeste';

        // Buscar nombre amigable del shopping
        const centroEncontrado = centrosCache.find(c => c.id === centroId);
        const centroNombre = centroEncontrado ? centroEncontrado.nombre : (centroId === 'al_oeste' ? 'Al Oeste Shopping' : centroId);

        let rolBadge = '';
        switch(rol) {
            case 'super_admin':
                rolBadge = '<span class="badge badge-warning">Super Admin</span>';
                break;
            case 'supervisor':
                rolBadge = '<span class="badge" style="background: var(--dash-purple-glow); color: var(--dash-purple); border: 1px solid rgba(167, 139, 250, 0.3);">Supervisor</span>';
                break;
            case 'acceso_principal':
                rolBadge = '<span class="badge badge-info">Garita Principal</span>';
                break;
            case 'acceso_satelite':
                rolBadge = '<span class="badge badge-info" style="opacity: 0.8;">Garita Satélite</span>';
                break;
            default:
                rolBadge = `<span class="badge badge-info">${rol}</span>`;
        }

        const estadoBadge = activo 
            ? '<span class="badge badge-success"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:currentColor;margin-right:4px;"></span>Activo</span>'
            : '<span class="badge badge-danger"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:currentColor;margin-right:4px;"></span>Inactivo</span>';

        const toggleBtnText = activo ? 'Desactivar' : 'Activar';

        html += `
            <tr data-uid="${uid}">
                <td><strong>${escapeHtml(email)}</strong></td>
                <td>${escapeHtml(nombre)}</td>
                <td>${rolBadge}</td>
                <td><span class="badge" style="background: rgba(255,255,255,0.06); border: 1px solid var(--dash-border); color: var(--dash-text); text-transform:none;">${escapeHtml(centroNombre)}</span></td>
                <td>${estadoBadge}</td>
                <td style="text-align: right;">
                    <div style="display: inline-flex; gap: 6px; justify-content: flex-end;">
                        <button class="dash-btn" style="${activo ? 'background: rgba(255,255,255,0.08); border: 1px solid var(--dash-border); color: var(--dash-text);' : 'background: rgba(52, 211, 153, 0.15); color: var(--dash-success); border: 1px solid rgba(52, 211, 153, 0.3);'} font-size: 0.75rem; padding: 4px 8px;" onclick="window.toggleEstadoUsuario('${uid}', ${activo})">
                            ${toggleBtnText}
                        </button>
                        <button class="dash-btn" style="background: rgba(248, 113, 113, 0.15); color: var(--dash-danger); border: 1px solid rgba(248, 113, 113, 0.3); font-size: 0.75rem; padding: 4px 8px;" onclick="window.eliminarUsuarioDoc('${uid}', '${escapeHtml(email)}')">
                            <span class="material-icons" style="font-size: 1rem;">delete</span>
                        </button>
                    </div>
                </td>
            </tr>
        `;
    });

    bodyUsuarios.innerHTML = html;
}

function initUsuarios() {
    const bodyUsuarios = document.getElementById('bodyUsuarios');
    if (!bodyUsuarios) return;

    onSnapshot(collection(db, "usuarios"), (snapshot) => {
        rawUsuariosDocs = snapshot.docs;
        renderUsuariosConCentros();
    }, (error) => {
        console.error("Error al cargar usuarios:", error);
        bodyUsuarios.innerHTML = `<tr><td colspan="6" class="empty-state" style="color:var(--dash-danger);">Error al cargar usuarios: ${error.message}</td></tr>`;
    });
}

window.toggleEstadoUsuario = async function(uid, currentActivo) {
    try {
        await updateDoc(doc(db, "usuarios", uid), {
            activo: !currentActivo
        });
    } catch(err) {
        alert("Error al actualizar estado del usuario: " + err.message);
    }
};

window.eliminarUsuarioDoc = async function(uid, email) {
    if (!confirm(`¿Está seguro de eliminar el usuario "${email}" del sistema? Esta acción no se puede deshacer.`)) {
        return;
    }
    try {
        await deleteDoc(doc(db, "usuarios", uid));
    } catch(err) {
        alert("Error al eliminar usuario: " + err.message);
    }
};

window.crearNuevoUsuario = async function(event) {
    if (event) event.preventDefault();
    const nombreInput = document.getElementById('newUserName');
    const emailInput = document.getElementById('newUserEmail');
    const passInput = document.getElementById('newUserPassword');
    const roleInput = document.getElementById('newUserRole');
    const centroInput = document.getElementById('newUserCentro');
    const msgEl = document.getElementById('msgCrearUsuario');
    const btn = document.getElementById('btnCrearUsuario');

    const nombre = nombreInput ? nombreInput.value.trim() : '';
    const email = emailInput ? emailInput.value.trim() : '';
    const password = passInput ? passInput.value : '';
    const rol = roleInput ? roleInput.value : 'acceso_principal';
    const centroId = centroInput ? centroInput.value : 'al_oeste';

    if (!email || !password || !rol) {
        msgEl.style.color = 'var(--dash-danger)';
        msgEl.textContent = 'Por favor complete todos los campos obligatorios.';
        return;
    }

    if (password.length < 6) {
        msgEl.style.color = 'var(--dash-danger)';
        msgEl.textContent = 'La contraseña debe tener al menos 6 caracteres.';
        return;
    }

    if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<span class="material-icons rotating">sync</span> Creando usuario...';
    }
    msgEl.style.color = 'var(--dash-info)';
    msgEl.textContent = 'Creando usuario y credenciales en Firebase...';

    try {
        const secAuth = getSecondaryAuth();
        const userCredential = await createUserWithEmailAndPassword(secAuth, email, password);
        const newUid = userCredential.user.uid;

        // Inmediatamente desconectar la sesión secundaria para mantener limpia la sesión principal
        await signOut(secAuth);

        // Guardar documento de rol y perfil en Firestore con centro_id
        await setDoc(doc(db, "usuarios", newUid), {
            email: email,
            nombre: nombre || email.split('@')[0],
            rol: rol,
            activo: true,
            centro_id: centroId,
            creado_el: serverTimestamp()
        });

        msgEl.style.color = 'var(--dash-success)';
        msgEl.textContent = `¡Usuario "${email}" creado exitosamente!`;

        if (nombreInput) nombreInput.value = '';
        if (emailInput) emailInput.value = '';
        if (passInput) passInput.value = '';

        setTimeout(() => {
            if (msgEl && msgEl.textContent.includes('exitosamente')) {
                msgEl.textContent = '';
            }
        }, 5000);

    } catch (e) {
        console.error("Error al crear usuario:", e);
        msgEl.style.color = 'var(--dash-danger)';
        if (e.code === 'auth/email-already-in-use') {
            msgEl.textContent = 'Error: El correo ya se encuentra en uso por otra cuenta.';
        } else if (e.code === 'auth/invalid-email') {
            msgEl.textContent = 'Error: El formato de correo es inválido.';
        } else if (e.code === 'auth/weak-password') {
            msgEl.textContent = 'Error: La contraseña es demasiado débil (mínimo 6 caracteres).';
        } else {
            msgEl.textContent = 'Error al crear usuario: ' + e.message;
        }
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = '<span class="material-icons">add_circle</span> Crear Usuario';
        }
    }
};

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
    let desde = document.getElementById('auditDesde').value;
    let hasta = document.getElementById('auditHasta').value;
    const tipo = document.getElementById('auditTipo').value;

    if (!desde || !hasta) {
        alert('Seleccioná ambas fechas');
        return;
    }

    if (desde > hasta) {
        const tmp = desde;
        desde = hasta;
        hasta = tmp;
        document.getElementById('auditDesde').value = desde;
        document.getElementById('auditHasta').value = hasta;
    }

    const btn = document.getElementById('btnBuscarAudit');
    btn.disabled = true;
    btn.innerHTML = '<span class="material-icons">hourglass_top</span> Buscando...';

    try {
        let records = [];

        if (tipo === 'fichajes') {
            try {
                const qFichajes = query(
                    collection(db, "fichajes"),
                    where("Fecha", ">=", desde),
                    where("Fecha", "<=", hasta)
                );
                const snap = await getDocs(qFichajes);
                snap.forEach(d => records.push({ id: d.id, ...d.data() }));
                records.sort((a, b) => {
                    const dtA = (a.Fecha || '') + ' ' + (a.Hora_Entrada || a.hora_entrada || '00:00:00');
                    const dtB = (b.Fecha || '') + ' ' + (b.Hora_Entrada || b.hora_entrada || '00:00:00');
                    return dtB.localeCompare(dtA);
                });
            } catch (errFich) {
                console.warn('Colección fichajes no disponible en Firestore:', errFich);
                records = [];
            }
        } else {
            // Accesos en Firestore
            const qAccesos = query(
                collection(db, "accesos"),
                where("Fecha", ">=", desde),
                where("Fecha", "<=", hasta)
            );
            const snap = await getDocs(qAccesos);
            snap.forEach(d => records.push({ id: d.id, ...d.data() }));

            records.sort((a, b) => {
                const dtA = (a.Fecha || '') + ' ' + (a.Hora_Ingreso || a.Hora_Salida || a.hora || '00:00:00');
                const dtB = (b.Fecha || '') + ' ' + (b.Hora_Ingreso || b.Hora_Salida || b.hora || '00:00:00');
                return dtB.localeCompare(dtA);
            });
        }

        auditoriaData = records;
        renderAuditoria(tipo);

        const info = document.getElementById('auditResultInfo');
        if (info) {
            info.style.display = 'flex';
            info.innerHTML = `<span class="material-icons">info</span> ${records.length} registros encontrados — Período: ${desde} al ${hasta}`;
        }
    } catch (e) {
        console.error('Error en auditoría Firestore:', e);
        alert('Error al buscar registros en Firebase: ' + (e.message || e));
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<span class="material-icons">search</span> Buscar';
    }
}

function renderAuditoria(tipo) {
    const thead = document.getElementById('headAuditoria');
    const tbody = document.getElementById('bodyAuditoria');

    if (auditoriaData.length === 0) {
        thead.innerHTML = '<tr><th colspan="9" class="empty-state">No se encontraron registros para el período seleccionado</th></tr>';
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
            <td>${r['Nombre y Apellido'] || r.nombre || ''}</td>
            <td>${r.Hora_Entrada || r.hora_entrada || ''}</td>
            <td>${r.Hora_Salida || r.hora_salida || '—'}</td>
        </tr>`).join('');
    } else {
        thead.innerHTML = `<tr>
            <th>Fecha</th><th>Hora</th><th>DNI</th><th>Nombre</th>
            <th>Evento</th><th>Tipo Permiso</th><th>Local/Destino</th><th>Resultado</th><th>Observaciones / Nota</th>
        </tr>`;
        tbody.innerHTML = auditoriaData.map(r => {
            const resultado = (r.Resultado || '').toUpperCase();
            let badgeClass = 'badge-info';
            if (resultado.includes('AUTORIZADO') || resultado === 'VERDE') badgeClass = 'badge-success';
            else if (resultado.includes('DENEGADO') || resultado === 'ROJO') badgeClass = 'badge-danger';

            const comentario = r.Comentario || r.comentario || r.observacion || r.Observacion || r.motivo_rechazo || r.Motivo_Rechazo || '';
            const comentarioHtml = comentario 
                ? `<span style="color:#fbbf24; font-weight:500; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px; background:rgba(251,191,36,0.12); padding:3px 8px; border-radius:4px; border:1px solid rgba(251,191,36,0.3);">📝 ${escapeHtml(comentario)}</span>`
                : '<span style="color:var(--dash-text-muted); opacity:0.4;">—</span>';

            return `<tr>
                <td>${r.Fecha || ''}</td>
                <td>${r.Hora_Ingreso || r.Hora_Salida || r.hora || ''}</td>
                <td><strong>${r.DNI || ''}</strong></td>
                <td>${escapeHtml(r['Nombre y Apellido'] || r.nombre || '')}</td>
                <td>${r.Evento || r.evento || ''}</td>
                <td>${r.Tipo_Permiso || r.tipo_permiso || ''}</td>
                <td>${escapeHtml(r.Local || r.local || '—')}</td>
                <td><span class="badge ${badgeClass}">${r.Resultado || ''}</span></td>
                <td>${comentarioHtml}</td>
            </tr>`;
        }).join('');
    }
}

function exportarAuditoria() {
    const desde = document.getElementById('auditDesde').value;
    const hasta = document.getElementById('auditHasta').value;
    const tipo = document.getElementById('auditTipo').value;

    if (!auditoriaData || auditoriaData.length === 0) {
        alert('No hay registros cargados para exportar. Por favor seleccioná un período y hacé clic en "Buscar".');
        return;
    }

    let csvContent = '\uFEFF'; // BOM UTF-8 para abrir directo en Excel en Windows con tildes y caracteres en español

    if (tipo === 'fichajes') {
        csvContent += 'Fecha;DNI;Nombre y Apellido;Hora Entrada;Hora Salida\r\n';
        auditoriaData.forEach(r => {
            const fecha = `"${String(r.Fecha || '').replace(/"/g, '""')}"`;
            const dni = `"${String(r.DNI || '').replace(/"/g, '""')}"`;
            const nombre = `"${String(r['Nombre y Apellido'] || r.nombre || '').replace(/"/g, '""')}"`;
            const entrada = `"${String(r.Hora_Entrada || r.hora_entrada || '').replace(/"/g, '""')}"`;
            const salida = `"${String(r.Hora_Salida || r.hora_salida || '').replace(/"/g, '""')}"`;
            csvContent += `${fecha};${dni};${nombre};${entrada};${salida}\r\n`;
        });
    } else {
        csvContent += 'Fecha;Hora;DNI;Nombre y Apellido;Evento;Tipo Permiso;Local;Resultado;Comentario\r\n';
        auditoriaData.forEach(r => {
            const fecha = `"${String(r.Fecha || '').replace(/"/g, '""')}"`;
            const hora = `"${String(r.Hora_Ingreso || r.Hora_Salida || r.hora || '').replace(/"/g, '""')}"`;
            const dni = `"${String(r.DNI || '').replace(/"/g, '""')}"`;
            const nombre = `"${String(r['Nombre y Apellido'] || r.nombre || '').replace(/"/g, '""')}"`;
            const evento = `"${String(r.Evento || r.evento || '').replace(/"/g, '""')}"`;
            const tipoPermiso = `"${String(r.Tipo_Permiso || r.tipo_permiso || '').replace(/"/g, '""')}"`;
            const local = `"${String(r.Local || r.local || '').replace(/"/g, '""')}"`;
            const resultado = `"${String(r.Resultado || '').replace(/"/g, '""')}"`;
            const comentario = `"${String(r.Comentario || r.comentario || '').replace(/"/g, '""')}"`;
            csvContent += `${fecha};${hora};${dni};${nombre};${evento};${tipoPermiso};${local};${resultado};${comentario}\r\n`;
        });
    }

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `auditoria_${tipo}_${desde}_al_${hasta}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
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
        let esVencido = false;
        if (item.vence) {
            const parts = item.vence.split('/');
            if (parts.length === 3) {
                const fVence = new Date(parts[2], parts[1]-1, parts[0], 23, 59, 59);
                const umbral = new Date(hoy);
                if (window.diasAlertaVencimiento) {
                    umbral.setDate(umbral.getDate() + window.diasAlertaVencimiento);
                }

                if (fVence < hoy) {
                    vencimientoClass = 'vencido-text';
                    esVencido = true;
                } else if (fVence <= umbral && window.diasAlertaVencimiento > 0) {
                    vencimientoClass = 'por-vencer-text';
                }
            }
        }
        
        const badgeClass = esVencido ? 'status-denied' : 'status-ok';
        const badgeText = esVencido ? 'Vencido' : item.tipo;

        html += `
            <tr>
                <td>${item.dni}</td>
                <td>${item.nombre}</td>
                <td>${item.empresa}</td>
                <td class="${vencimientoClass}">${item.vence || 'N/A'}</td>
                <td><span class="status-badge ${badgeClass}">${badgeText}</span></td>
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

    const hoyStr = formatDateToLocalString();

    let filtered = tramitesData.filter(item => {
        const matchesText = item.empresa.toLowerCase().includes(filterText) || item.id.toString().includes(filterText);
        const fFin = item.fecha_fin ? item.fecha_fin.split('T')[0] : '';
        const esVencido = Boolean(fFin && fFin < hoyStr);

        let matchesEstado = true;
        if (filterEstado === 'VIGENTES') {
            matchesEstado = !esVencido && (item.estado_codigo === '6' || item.estado_codigo === '4');
        } else if (filterEstado === 'VENCIDOS') {
            matchesEstado = esVencido;
        } else if (filterEstado !== 'TODOS') {
            matchesEstado = item.estado_codigo === filterEstado;
        }
        return matchesText && matchesEstado;
    });

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="empty-state">No se encontraron trámites.</td></tr>';
        return;
    }

    filtered.sort((a, b) => b.id - a.id);

    let html = '';
    filtered.forEach(item => {
        const fFin = item.fecha_fin ? item.fecha_fin.split('T')[0] : '';
        const esVencido = Boolean(fFin && fFin < hoyStr);

        let badgeClass = 'status-denied';
        let estadoStr = 'Desconocido';
        if (esVencido) {
            badgeClass = 'status-denied';
            estadoStr = (item.estado_codigo === '6' || item.estado_codigo === '4') ? 'Vencido' : 'Rechazado / Vencido';
        } else if (item.estado_codigo === '6') { 
            badgeClass = 'status-ok'; 
            estadoStr = 'Aprobado'; 
        } else if (item.estado_codigo === '4') { 
            badgeClass = 'status-ok'; 
            estadoStr = 'Aprobado Parcial'; 
        } else if (item.estado_codigo === '3') { 
            badgeClass = ''; 
            estadoStr = 'Enviado'; 
        } else if (item.estado_codigo === '5') { 
            badgeClass = 'status-denied'; 
            estadoStr = 'Requiere Cambios'; 
        }
        
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
                <td><span class="status-badge ${esVencido ? 'status-denied' : (item.tipo === 'FAO' ? 'status-ok' : 'status-denied')}">${item.tipo}</span></td>
                <td style="font-weight: 500;">#${item.id}</td>
                <td>${item.empresa || '-'}</td>
                <td>${item.fecha_inicio || '-'}</td>
                <td class="${esVencido ? 'vencido-text' : ''}">${item.fecha_fin || '-'}</td>
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
    
    async function callSyncIrsaCloud(payload = {}) {
        const endpoints = [
            'https://sync-irsa-http-sbm3jcejya-uc.a.run.app',
            'https://us-central1-control-acceso-c2f82.cloudfunctions.net/sync_irsa_http'
        ];
        let lastError = null;
        for (const url of endpoints) {
            try {
                const res = await fetch(url, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const data = await res.json();
                return data;
            } catch (err) {
                lastError = err;
                console.warn(`Fallo al llamar a ${url}, intentando alternativo:`, err);
            }
        }
        throw lastError;
    }

    if (btnIrsa) {
        btnIrsa.addEventListener('click', async () => {
            const orig = '<span class="material-icons">sync</span> Extraer Nóminas IRSA Ahora';
            btnIrsa.disabled = true;
            btnIrsa.innerHTML = '<span class="material-icons">hourglass_empty</span> Sincronizando en la nube...';
            statusMsg.textContent = 'Conectando con la Cloud Function en Google Cloud...';
            statusMsg.style.color = 'var(--dash-primary)';

            try {
                const data = await callSyncIrsaCloud({});
                if (data.success) {
                    statusMsg.textContent = data.result || 'Sincronización completada con éxito en la nube.';
                    statusMsg.style.color = 'var(--dash-success)';
                    loadDashboardData();
                } else {
                    statusMsg.textContent = 'Error: ' + (data.error || 'Falló la sincronización en la nube.');
                    statusMsg.style.color = 'var(--dash-danger)';
                }
            } catch (err) {
                console.error("Error llamando Cloud Function:", err);
                statusMsg.textContent = 'Error de conexión con la Cloud Function: ' + (err.message || err);
                statusMsg.style.color = 'var(--dash-danger)';
            } finally {
                btnIrsa.disabled = false;
                btnIrsa.innerHTML = orig;
            }
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
        btnManual.addEventListener('click', async () => {
            const val = manualInput.value.trim();
            if(!val) {
                alert("Por favor ingrese el número de Trámite (FAO o FAP).");
                return;
            }
            const orig = '<span class="material-icons">download</span> Descarga Forzada';
            btnManual.disabled = true;
            btnManual.innerHTML = '<span class="material-icons rotating">sync</span> Forzando...';
            statusMsg.textContent = `Consultando portal de IRSA para el trámite N° ${val}...`;
            statusMsg.style.color = 'var(--dash-primary)';

            try {
                const data = await callSyncIrsaCloud({ tramite_id: val });
                if (data.success) {
                    statusMsg.textContent = data.result || `¡Trámite N° ${val} descargado y personas autorizadas con éxito!`;
                    statusMsg.style.color = 'var(--dash-success)';
                    manualInput.value = '';
                    loadDashboardData();
                } else {
                    statusMsg.textContent = 'Error: ' + (data.error || 'No se pudo forzar el trámite en IRSA.');
                    statusMsg.style.color = 'var(--dash-danger)';
                }
            } catch (err) {
                console.error("Error forzando trámite en Cloud Function:", err);
                statusMsg.textContent = 'Error de conexión con el servidor en la nube: ' + (err.message || err);
                statusMsg.style.color = 'var(--dash-danger)';
            } finally {
                btnManual.disabled = false;
                btnManual.innerHTML = orig;
            }
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

// =============================================
// GESTIÓN Y LISTADO DE EXCEPCIONES Y NÓMINAS ESPECIALES
// =============================================
let excepcionesCache = [];

function initExcepciones() {
    const body = document.getElementById('bodyExcepciones');
    const filterInput = document.getElementById('filterExcepcionesInput');
    const filterEstado = document.getElementById('filterExcepcionesEstado');
    const btnRefresh = document.getElementById('btnRecargarExcepciones');
    if (!body) return;

    // Escuchar nóminas en tiempo real para filtrar las excepciones
    onSnapshot(collection(db, "nominas"), (snapshot) => {
        excepcionesCache = [];
        snapshot.forEach(docSnap => {
            const data = docSnap.data();
            const cat = String(data.categoria || data.tipo_permiso || '').toLowerCase();
            // Filtrar solo las que corresponden a excepciones / ART / especiales
            if (cat.includes('excep') || cat.includes('art') || cat.includes('especial')) {
                excepcionesCache.push({
                    id: docSnap.id,
                    ...data
                });
            }
        });

        // Ordenar por fecha_vencimiento descendente (las más nuevas primero)
        excepcionesCache.sort((a, b) => {
            const vA = a.fecha_vencimiento || '';
            const vB = b.fecha_vencimiento || '';
            if (vB !== vA) return vB.localeCompare(vA);
            return String(a.nombre || '').localeCompare(String(b.nombre || ''));
        });

        renderTablaExcepciones();
    }, (err) => {
        console.error("Error al escuchar excepciones en tiempo real:", err);
        body.innerHTML = `<tr><td colspan="6" class="empty-state" style="color:var(--dash-danger);">Error al cargar excepciones: ${err.message}</td></tr>`;
    });

    if (filterInput) {
        filterInput.addEventListener('input', renderTablaExcepciones);
    }
    if (filterEstado) {
        filterEstado.addEventListener('change', renderTablaExcepciones);
    }
    if (btnRefresh) {
        btnRefresh.addEventListener('click', () => {
            btnRefresh.classList.add('rotating');
            renderTablaExcepciones();
            setTimeout(() => btnRefresh.classList.remove('rotating'), 600);
        });
    }
}

function renderTablaExcepciones() {
    const body = document.getElementById('bodyExcepciones');
    const badgeTotal = document.getElementById('badgeTotalExcepciones');
    const filterInput = document.getElementById('filterExcepcionesInput');
    const filterEstado = document.getElementById('filterExcepcionesEstado');
    if (!body) return;

    const queryText = (filterInput ? filterInput.value.toLowerCase().trim() : '');
    const estadoFiltro = (filterEstado ? filterEstado.value : 'VIGENTES');
    const hoy = formatDateToLocalString();

    let totalVigentes = 0;
    let totalVencidas = 0;

    excepcionesCache.forEach(item => {
        const venc = item.fecha_vencimiento || '';
        const esVigente = venc >= hoy && item.activo !== false;
        if (esVigente) totalVigentes++;
        else totalVencidas++;
    });

    if (badgeTotal) {
        badgeTotal.innerHTML = `Total: <strong>${excepcionesCache.length}</strong> (${totalVigentes} vigentes / ${totalVencidas} vencidas)`;
    }

    const filtrados = excepcionesCache.filter(item => {
        const venc = item.fecha_vencimiento || '';
        const esVigente = venc >= hoy && item.activo !== false;

        if (estadoFiltro === 'VIGENTES' && !esVigente) return false;
        if (estadoFiltro === 'VENCIDAS' && esVigente) return false;

        if (queryText) {
            const dni = String(item.dni || '').toLowerCase();
            const nom = String(item.nombre || '').toLowerCase();
            const emp = String(item.empresa || item.motivo || item.puesto_especifico || '').toLowerCase();
            if (!dni.includes(queryText) && !nom.includes(queryText) && !emp.includes(queryText)) {
                return false;
            }
        }
        return true;
    });

    if (filtrados.length === 0) {
        body.innerHTML = `<tr><td colspan="6" class="empty-state" style="text-align: center; padding: 2rem;">No se encontraron excepciones ${queryText || estadoFiltro !== 'TODAS' ? 'con los filtros seleccionados' : 'registradas en el sistema'}.</td></tr>`;
        return;
    }

    body.innerHTML = filtrados.map(item => {
        const venc = item.fecha_vencimiento || '-';
        const esVigente = venc >= hoy && item.activo !== false;
        const esHoy = venc === hoy;

        let estadoBadge = '';
        if (esVigente) {
            if (esHoy) {
                estadoBadge = '<span class="badge" style="background: rgba(245, 158, 11, 0.15); color: #f59e0b; border: 1px solid rgba(245, 158, 11, 0.3); font-weight: 600;"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:currentColor;margin-right:4px;"></span>Vence Hoy</span>';
            } else {
                estadoBadge = '<span class="badge" style="background: rgba(52, 211, 153, 0.15); color: var(--dash-success); border: 1px solid rgba(52, 211, 153, 0.3); font-weight: 600;"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:currentColor;margin-right:4px;"></span>Vigente</span>';
            }
        } else {
            estadoBadge = '<span class="badge" style="background: rgba(239, 68, 68, 0.15); color: var(--dash-danger); border: 1px solid rgba(239, 68, 68, 0.3); font-weight: 600;"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:currentColor;margin-right:4px;"></span>Vencida</span>';
        }

        const empMotivo = item.empresa || item.motivo || item.puesto_especifico || '-';
        const nombre = item.nombre || '-';
        const dni = item.dni || '-';

        return `
            <tr data-id="${item.id}">
                <td><strong>${escapeHtml(dni)}</strong></td>
                <td>${escapeHtml(nombre)}</td>
                <td>${escapeHtml(empMotivo)}</td>
                <td><code style="font-size:0.9rem;">${escapeHtml(venc)}</code></td>
                <td>${estadoBadge}</td>
                <td style="text-align: right;">
                    <div style="display: inline-flex; gap: 6px; justify-content: flex-end;">
                        <button class="dash-btn" style="background: rgba(59, 130, 246, 0.15); color: var(--dash-primary); border: 1px solid rgba(59, 130, 246, 0.3); font-size: 0.75rem; padding: 5px 10px;" onclick="window.abrirModalEditarExcepcion('${item.id}')" title="Modificar fecha de vigencia o datos">
                            <span class="material-icons" style="font-size: 0.95rem;">edit</span> Modificar
                        </button>
                        <button class="dash-btn" style="background: rgba(239, 68, 68, 0.15); color: var(--dash-danger); border: 1px solid rgba(239, 68, 68, 0.3); font-size: 0.75rem; padding: 5px 10px;" onclick="window.eliminarExcepcionDoc('${item.id}', '${escapeHtml(nombre)}', '${escapeHtml(dni)}')" title="Eliminar excepción">
                            <span class="material-icons" style="font-size: 0.95rem;">delete</span> Borrar
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

window.abrirModalEditarExcepcion = function(docId) {
    const item = excepcionesCache.find(x => x.id === docId);
    if (!item) return;

    document.getElementById('editExcepcionDocId').value = docId;
    document.getElementById('editExcepcionDni').value = item.dni || '';
    document.getElementById('editExcepcionNombre').value = item.nombre || '';
    document.getElementById('editExcepcionEmpresa').value = item.empresa || item.motivo || item.puesto_especifico || '';
    document.getElementById('editExcepcionVencimiento').value = item.fecha_vencimiento || '';

    const statusEl = document.getElementById('editExcepcionStatus');
    if (statusEl) {
        statusEl.style.display = 'none';
        statusEl.textContent = '';
    }

    const modal = document.getElementById('modalEditarExcepcion');
    if (modal) modal.style.display = 'flex';
};

window.cerrarModalEditarExcepcion = function() {
    const modal = document.getElementById('modalEditarExcepcion');
    if (modal) modal.style.display = 'none';
};

window.guardarEdicionExcepcion = async function() {
    const docId = document.getElementById('editExcepcionDocId').value;
    const nombre = document.getElementById('editExcepcionNombre').value.trim();
    const empresa = document.getElementById('editExcepcionEmpresa').value.trim();
    const fechaVenc = document.getElementById('editExcepcionVencimiento').value;
    const statusEl = document.getElementById('editExcepcionStatus');
    const btnGuardar = document.getElementById('btnGuardarEditExcepcion');

    if (!docId) return;
    if (!nombre || !empresa || !fechaVenc) {
        if (statusEl) {
            statusEl.style.display = 'block';
            statusEl.style.color = 'var(--dash-danger)';
            statusEl.textContent = 'Por favor complete todos los campos.';
        }
        return;
    }

    if (btnGuardar) {
        btnGuardar.disabled = true;
        btnGuardar.innerHTML = '<span class="material-icons rotating">sync</span> Guardando...';
    }

    try {
        await updateDoc(doc(db, "nominas", docId), {
            nombre: nombre,
            empresa: empresa,
            fecha_vencimiento: fechaVenc,
            activo: true
        });

        if (statusEl) {
            statusEl.style.display = 'block';
            statusEl.style.color = 'var(--dash-success)';
            statusEl.textContent = '¡Excepción modificada con éxito!';
        }

        setTimeout(() => {
            window.cerrarModalEditarExcepcion();
            if (btnGuardar) {
                btnGuardar.disabled = false;
                btnGuardar.innerHTML = '<span class="material-icons">save</span> Guardar Cambios';
            }
        }, 600);
    } catch (e) {
        console.error("Error al modificar excepción:", e);
        if (statusEl) {
            statusEl.style.display = 'block';
            statusEl.style.color = 'var(--dash-danger)';
            statusEl.textContent = 'Error: ' + e.message;
        }
        if (btnGuardar) {
            btnGuardar.disabled = false;
            btnGuardar.innerHTML = '<span class="material-icons">save</span> Guardar Cambios';
        }
    }
};

window.eliminarExcepcionDoc = async function(docId, nombre, dni) {
    if (!confirm(`¿Está seguro de eliminar la excepción para:\n\n${nombre} (DNI ${dni})?\n\nEsta persona dejará de estar autorizada para ingresar.`)) {
        return;
    }

    try {
        await deleteDoc(doc(db, "nominas", docId));
    } catch (e) {
        console.error("Error al eliminar excepción:", e);
        alert("Error al eliminar excepción: " + e.message);
    }
};

