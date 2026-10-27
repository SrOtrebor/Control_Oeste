import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getAuth, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-auth.js";
import { getFirestore, collection, addDoc, query, where, getDocs, onSnapshot, serverTimestamp, doc, getDoc, updateDoc, setDoc, deleteDoc, enableIndexedDbPersistence } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-firestore.js";

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

// Activar Persistencia Offline de Firebase (Soporte sin Internet)
enableIndexedDbPersistence(db).catch((err) => {
    if (err.code == 'failed-precondition') {
        console.warn("Múltiples pestañas abiertas, la persistencia offline solo funciona en una.");
    } else if (err.code == 'unimplemented') {
        console.warn("El navegador no soporta persistencia offline.");
    }
});

// Función principal de inicialización del escáner
function initScanner() {

    // --- PROTECCIÓN DE LOGIN Y ROLES ---
    let currentUserRole = null;

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
                document.getElementById('header-title').textContent = `Acceso ${currentUserRole === 'acceso_satelite' ? 'Satélite' : 'Principal'}`;
                
                // Si es satélite, ocultamos botones de funciones avanzadas
                if (currentUserRole === 'acceso_satelite') {
                    const btnFichar = document.getElementById('btnFicharHuella');
                    if(btnFichar) btnFichar.style.display = 'none';
                    // Aquí podrías ocultar otras cosas si querés
                }
            } else {
                console.error("Usuario sin rol asignado.");
                signOut(auth);
            }
        } catch(e) {
            console.error("Error verificando rol", e);
        }
    });

    // Logout
    const btnLogout = document.getElementById('btnLogout');
    if (btnLogout) {
        btnLogout.addEventListener('click', () => {
            signOut(auth);
        });
    }

    // --- ESTADO GLOBAL ---
    let currentMode = 'entrada'; // 'entrada', 'visita', 'salida'

    // --- CACHE LOCAL PARA VELOCIDAD ---
    // localNominas almacena una lista (Array) de permisos por cada DNI para soportar múltiples permisos por persona
    const localNominas = new Map();
    const localListaNegra = new Map();

    function setupCacheListeners() {
        onSnapshot(collection(db, "nominas"), (snapshot) => {
            localNominas.clear();
            snapshot.forEach(docSnap => {
                const data = docSnap.data();
                let rawDni = String(data.dni || docSnap.id || '').trim();
                if (rawDni.endsWith('.0')) rawDni = rawDni.slice(0, -2);
                rawDni = rawDni.replace(/\D/g, '').trim();
                if (rawDni) {
                    if (!localNominas.has(rawDni)) {
                        localNominas.set(rawDni, []);
                    }
                    localNominas.get(rawDni).push({ ...data, docId: docSnap.id });
                }
            });
            console.log(`[Cache] Nóminas sincronizadas en memoria: ${snapshot.size} docs para ${localNominas.size} personas.`);
        }, (error) => {
            console.error("[Cache] Error en snapshot nominas:", error);
        });

        onSnapshot(collection(db, "lista_negra"), (snapshot) => {
            localListaNegra.clear();
            snapshot.forEach(doc => {
                const data = doc.data();
                const key = String(data.dni || doc.id || '').trim();
                if (key) {
                    localListaNegra.set(key, data);
                }
            });
            console.log(`[Cache] Lista negra sincronizada: ${localListaNegra.size}`);
        }, (error) => {
            console.error("[Cache] Error en snapshot lista negra:", error);
        });
    }

    // Formatear Fecha Hoy (YYYY-MM-DD)
    function getTodayString() {
        const now = new Date();
        const y = now.getFullYear();
        const m = String(now.getMonth() + 1).padStart(2, '0');
        const d = String(now.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }

    // Formatear Hora Local (HH:MM:SS)
    function getTimeString() {
        return new Date().toLocaleTimeString('es-AR', { hour12: false });
    }

    // Normalizar fechas variadas (YYYY-MM-DD o DD/MM/YYYY) a formato ISO YYYY-MM-DD
    function parseDateToISO(val) {
        if (!val) return null;
        const s = String(val).trim();
        const mISO = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (mISO) {
            return `${mISO[1]}-${mISO[2].padStart(2, '0')}-${mISO[3].padStart(2, '0')}`;
        }
        const mDMY = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
        if (mDMY) {
            return `${mDMY[3]}-${mDMY[2].padStart(2, '0')}-${mDMY[1].padStart(2, '0')}`;
        }
        return null;
    }

    // Formatear YYYY-MM-DD a DD/MM/YYYY para visualización clara
    function formatISODateToDisplay(isoStr) {
        if (!isoStr) return '';
        const m = isoStr.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) {
            return `${m[3]}/${m[2]}/${m[1]}`;
        }
        return isoStr;
    }

    // --- EVALUACIÓN Y RESOLUCIÓN DE PERMISOS MÚLTIPLES ---
    function evaluarPermisoIndividual(cand) {
        const rawVenc = cand.fecha_vencimiento || cand.fecha_fin || cand.vence || cand.vencimiento || '';
        const vencISO = parseDateToISO(rawVenc);
        const hoyIso = getTodayString();
        
        const esActivo = cand.activo !== false;
        let esVencido = false;
        let diasParaVencer = null;
        let porVencer = false;

        if (vencISO) {
            if (vencISO < hoyIso) {
                esVencido = true;
            } else {
                try {
                    const [yH, mH, dH] = hoyIso.split('-').map(Number);
                    const [yV, mV, dV] = vencISO.split('-').map(Number);
                    const dateHoy = new Date(yH, mH - 1, dH);
                    const dateVenc = new Date(yV, mV - 1, dV);
                    const diffDays = Math.round((dateVenc.getTime() - dateHoy.getTime()) / (1000 * 60 * 60 * 24));
                    diasParaVencer = diffDays;
                    if (diffDays >= 0 && diffDays <= 5) {
                        porVencer = true;
                    }
                } catch (e) {
                    console.warn("Error calculando días para vencer:", e);
                }
            }
        }

        return {
            ...cand,
            rawVenc,
            vencISO,
            vencimientoDisplay: vencISO ? formatISODateToDisplay(vencISO) : (rawVenc || 'Indefinido'),
            esActivo,
            esVencido,
            porVencer,
            diasParaVencer,
            esIndefinido: !vencISO
        };
    }

    function resolverPermisosPersona(candList) {
        if (!candList || candList.length === 0) return null;

        const evaluados = candList.map(evaluarPermisoIndividual);

        // Separar en vigentes activos vs vencidos vs inactivos
        const vigentes = evaluados.filter(p => p.esActivo && !p.esVencido);
        const vencidos = evaluados.filter(p => p.esVencido);
        const inactivos = evaluados.filter(p => !p.esActivo && !p.esVencido);

        if (vigentes.length > 0) {
            // La persona TIENE al menos un permiso vigente y activo!
            // Ordenar:
            // 1. Indefinidos primero
            // 2. Permiso con fecha más lejana en el futuro (más favorable)
            vigentes.sort((a, b) => {
                if (a.esIndefinido && !b.esIndefinido) return -1;
                if (!a.esIndefinido && b.esIndefinido) return 1;
                if (a.esIndefinido && b.esIndefinido) return 0;
                return (b.vencISO || '').localeCompare(a.vencISO || '');
            });

            const mejorVigente = vigentes[0];
            return {
                estado: 'VIGENTE',
                permiso: mejorVigente,
                esVencido: false,
                porVencer: mejorVigente.porVencer,
                diasParaVencer: mejorVigente.diasParaVencer,
                vencimientoDisplay: mejorVigente.vencimientoDisplay,
                todosVigentes: vigentes
            };
        }

        if (vencidos.length > 0) {
            // No tiene ningún permiso vigente, pero sí tiene permisos vencidos.
            // Ordenar vencidos por fecha más reciente (el que venció más tarde)
            vencidos.sort((a, b) => {
                return (b.vencISO || '').localeCompare(a.vencISO || '');
            });
            const ultimoVencido = vencidos[0];
            return {
                estado: 'VENCIDO',
                permiso: ultimoVencido,
                esVencido: true,
                porVencer: false,
                diasParaVencer: null,
                vencimientoDisplay: ultimoVencido.vencimientoDisplay,
                todosVencidos: vencidos
            };
        }

        if (inactivos.length > 0) {
            return {
                estado: 'INACTIVO',
                permiso: inactivos[0],
                esVencido: false,
                porVencer: false,
                diasParaVencer: null,
                vencimientoDisplay: inactivos[0].vencimientoDisplay
            };
        }

        return null;
    }

    // Parsear DNI (soporta TODOS los formatos de PDF417 y 1D)
    function parseDNI(rawData) {
        if (!rawData) return null;
        let dni = rawData, apellido = 'No Encontrado', nombre = '';
        
        if (rawData.includes('@') || rawData.includes('"')) {
            const parts = rawData.split(/[@"]/);
            
            if (parts[0] === "") {
                // Formato RENAPER Antiguo (El DNI está al principio)
                dni = (parts[1] || '').trim();
                apellido = (parts[4] || '').trim();
                nombre = (parts[5] || '').trim();
            } else {
                // Formatos RENAPER Nuevos (El trámite está al principio)
                apellido = (parts[1] || '').trim();
                nombre = (parts[2] || '').trim();
                
                // Buscar el DNI inteligentemente (suele ser la posición 3 o 4)
                for (let i = 3; i < parts.length; i++) {
                    const p = parts[i].trim();
                    if (/^\d{7,8}$/.test(p)) {
                        dni = p;
                        break;
                    }
                }
            }
        } else if (rawData.length === 38 && /^\d+$/.test(rawData)) {
            // Formato de código de barras 1D (38 dígitos). El DNI está entre el índice 14 y 22.
            dni = rawData.substring(14, 22).replace(/^0+/, '');
            apellido = 'No Encontrado';
            nombre = '';
        }
        
        // Limpiar para asegurar que solo devuelva números
        dni = dni.replace(/\D/g, '');
        if (!dni) return null;
        return { dni, apellido, nombre };
    }

    function setMode(newMode) {
        currentMode = newMode;
        const dniInput = document.getElementById('dniInput');
        const headerTitle = document.getElementById('header-title');

        ['btnEntrada', 'btnRegVisita', 'btnSalida'].forEach(id => {
            document.getElementById(id)?.classList.remove('active-mode');
        });
        
        const activeBtnId = `btn${newMode.charAt(0).toUpperCase() + newMode.slice(1)}`;
        const activeBtn = document.getElementById(activeBtnId === 'btnRegVisita' ? 'btnRegVisita' : activeBtnId);
        activeBtn?.classList.add('active-mode');

        const header = document.querySelector('.header');
        document.body.classList.remove('control-acceso-mode', 'registrar-visita-mode', 'salida-visita-mode');
        header.classList.remove('control-acceso-mode', 'registrar-visita-mode', 'salida-visita-mode');

        if (newMode === 'entrada') {
            document.body.classList.add('control-acceso-mode');
            header.classList.add('control-acceso-mode');
            headerTitle.textContent = 'Control de Acceso';
        } else if (newMode === 'visita') {
            document.body.classList.add('registrar-visita-mode');
            header.classList.add('registrar-visita-mode');
            headerTitle.textContent = 'Registrar Visita';
        } else if (newMode === 'salida') {
            document.body.classList.add('salida-visita-mode');
            header.classList.add('salida-visita-mode');
            headerTitle.textContent = 'Registrar Salida';
        }

        if (dniInput) dniInput.focus();
    }

    // Cache de registros diarios
    let dailyRecordsCache = [];

    function renderDailyRecords() {
        const ingresosListBody = document.getElementById('ingresosListBody');
        if (!ingresosListBody) return;
        ingresosListBody.innerHTML = '';

        const searchInput = document.getElementById('searchIngresos');
        const filterSelect = document.getElementById('filterTipo');
        const searchTerm = searchInput ? searchInput.value.trim().toLowerCase() : '';
        const filterVal = filterSelect ? filterSelect.value.trim().toLowerCase() : '';

        const filtered = dailyRecordsCache.filter(record => {
            if (searchTerm) {
                const dni = String(record.DNI || '').toLowerCase();
                const nombre = String(record['Nombre y Apellido'] || '').toLowerCase();
                const local = String(record.Local || '').toLowerCase();
                const permiso = String(record.Tipo_Permiso || '').toLowerCase();
                const evento = String(record.Evento || '').toLowerCase();
                const matches = dni.includes(searchTerm) || nombre.includes(searchTerm) || local.includes(searchTerm) || permiso.includes(searchTerm) || evento.includes(searchTerm);
                if (!matches) return false;
            }
            if (filterVal) {
                const tipo = String(record.Tipo_Permiso || '').toLowerCase();
                if (!tipo.includes(filterVal)) return false;
            }
            return true;
        });

        if (filtered.length === 0) {
            const row = ingresosListBody.insertRow();
            const cell = row.insertCell(0);
            cell.colSpan = 7;
            cell.textContent = dailyRecordsCache.length === 0 ? 'Sin registros hoy.' : 'No se encontraron resultados para la búsqueda.';
            cell.style.textAlign = 'center';
            cell.style.color = '#888';
            cell.style.padding = '15px';
            return;
        }

        filtered.forEach(record => {
            const res = (record.Resultado || '').toUpperCase();
            const ev = (record.Evento || '').toUpperCase();
            const row = ingresosListBody.insertRow();
            const esAlertaVenc = Boolean(record.Alerta_Vencimiento);

            if (ev === 'SALIDA') {
                row.classList.add('fila-gris');
            } else if (res.includes('AUTORIZADO') || res.includes('PERMITIDO')) {
                if (esAlertaVenc) {
                    row.classList.add('fila-amarilla');
                } else {
                    const tipo = record.Tipo_Permiso || '';
                    if (tipo.includes('FAO')) row.classList.add('fila-verde', 'fila-fao');
                    else if (tipo.includes('FAP')) row.classList.add('fila-verde');
                    else if (tipo.includes('PERSONAL')) row.classList.add('fila-gris');
                    else if (tipo.includes('VISITA')) row.classList.add('fila-gris');
                    else if (tipo.includes('FICHAJE')) row.classList.add('fila-gris');
                    else row.classList.add('fila-verde');
                }
            } else {
                row.classList.add('fila-roja');
            }

            row.style.cursor = 'pointer';
            row.title = 'Clic para agregar comentario/observación';
            row.onclick = (e) => openComentarioModal(record, e);

            const timeString = record.Hora_Ingreso || record.Hora_Salida || '';
            row.insertCell().textContent = timeString;
            row.insertCell().textContent = record.DNI || '';
            
            let commentHtml = record.Comentario ? `<br><small style="font-style: italic; color: #fbbf24;">📝 ${record.Comentario}</small>` : '';
            row.insertCell().innerHTML = (record['Nombre y Apellido'] || '') + commentHtml;
            
            row.insertCell().textContent = record.Evento || '';
            row.insertCell().textContent = record.Tipo_Permiso || '';
            row.insertCell().textContent = record.Local || '';

            // Columna 7: Fecha de Vencimiento
            const cellVenc = row.insertCell();
            let vencVal = record.Vence || record.Fecha_Vencimiento || '';
            if (!vencVal && record.DNI) {
                let cleanDni = String(record.DNI).trim();
                if (cleanDni.endsWith('.0')) cleanDni = cleanDni.slice(0, -2);
                cleanDni = cleanDni.replace(/\D/g, '');
                const candList = localNominas.get(cleanDni) || [];
                const resP = resolverPermisosPersona(candList);
                if (resP && resP.permiso) {
                    vencVal = resP.vencimientoDisplay;
                }
            }

            if (vencVal) {
                const isoV = parseDateToISO(vencVal);
                const hoyIso = getTodayString();
                if (isoV && isoV < hoyIso) {
                    cellVenc.innerHTML = `<span style="color: #ef4444; font-weight: 600;">${vencVal} (Vencido)</span>`;
                } else if (isoV) {
                    try {
                        const [yH, mH, dH] = hoyIso.split('-').map(Number);
                        const [yV, mV, dV] = isoV.split('-').map(Number);
                        const diffDays = Math.round((new Date(yV, mV - 1, dV) - new Date(yH, mH - 1, dH)) / 86400000);
                        if (diffDays >= 0 && diffDays <= 5) {
                            const txt = diffDays === 0 ? 'HOY' : (diffDays === 1 ? 'MAÑANA' : `${diffDays}d`);
                            cellVenc.innerHTML = `<span style="color: #d97706; font-weight: 600;" title="Permiso próximo a vencer">⚠️ ${vencVal} (${txt})</span>`;
                        } else {
                            cellVenc.textContent = vencVal;
                        }
                    } catch (e) {
                        cellVenc.textContent = vencVal;
                    }
                } else {
                    cellVenc.textContent = vencVal;
                }
            } else {
                cellVenc.textContent = '-';
                cellVenc.style.color = '#94a3b8';
            }
        });
    }

    // Configurar Listeners en tiempo real para Stats y Tabla Diaria
    function setupRealtimeListeners() {
        // Escuchar estado_adentro directamente
        onSnapshot(collection(db, "estado_adentro"), (snapshot) => {
            const el = document.getElementById('statTotalAdentro');
            if (el) el.textContent = snapshot.size;
        });

        const today = getTodayString();
        const q = query(collection(db, "accesos"), where("Fecha", "==", today));
        
        onSnapshot(q, (snapshot) => {
            let entradas = 0;
            let salidas = 0;
            let rechazados = 0;
            dailyRecordsCache = [];

            snapshot.forEach(doc => {
                let d = doc.data();
                d.id = doc.id;
                dailyRecordsCache.push(d);
                const res = (d.Resultado || '').toUpperCase();
                const ev = (d.Evento || '').toUpperCase();

                if (res.includes('AUTORIZADO') || res.includes('PERMITIDO') || res.includes('REGISTRADO')) {
                    if (ev === 'SALIDA') {
                        salidas++;
                    } else {
                        entradas++;
                    }
                } else if (res.includes('DENEGADO')) {
                    rechazados++;
                }
            });
            
            // Ordenar descendente por hora
            dailyRecordsCache.sort((a, b) => {
                const timeA = a.Hora_Ingreso || a.Hora_Salida || '00:00:00';
                const timeB = b.Hora_Ingreso || b.Hora_Salida || '00:00:00';
                return timeB.localeCompare(timeA);
            });

            const statPerm = document.getElementById('statPermitidos');
            const statSalidas = document.getElementById('statSalidas');
            const statRech = document.getElementById('statRechazados');
            if (statPerm) statPerm.textContent = entradas;
            if (statSalidas) statSalidas.textContent = salidas;
            if (statRech) statRech.textContent = rechazados;

            renderDailyRecords();
        });

        // Buscador de DNI / Nombre en la tabla diaria
        const searchInput = document.getElementById('searchIngresos');
        if (searchInput) {
            searchInput.addEventListener('input', renderDailyRecords);
            searchInput.addEventListener('keyup', renderDailyRecords);
        }
        const filterSelect = document.getElementById('filterTipo');
        if (filterSelect) {
            filterSelect.addEventListener('change', renderDailyRecords);
        }
    }

    async function verificarDNI() {
        const dniInput = document.getElementById('dniInput');
        if (!dniInput) return;
        const rawData = dniInput.value.trim();
        if (!rawData) {
            dniInput.focus();
            return;
        }
        dniInput.value = ''; // limpiar rapido para el proximo escaneo

        try {
            const parsed = parseDNI(rawData);
            if (!parsed || !parsed.dni) {
                mostrarResultado('roja', 'DNI Inválido', 'Ingrese números de documento', '', '', '');
                dniInput.focus();
                return;
            }
            
            let acceso = 'DENEGADO';
            let mensaje = 'Acceso Denegado';
            let tipoPermiso = 'DESCONOCIDO';
            let colorClass = 'roja';
            let localInfo = 'GENERAL';
            let permisoVencido = false;
            let permisoPorVencer = false;
            let diasParaVencer = null;
            let vencimientoDisplay = '';
            
            // 1. Consultar Caché Local Lista Negra (por DNI exacto o CUIL)
            let lnData = localListaNegra.get(parsed.dni);
            if (!lnData) {
                for (const [key, value] of localListaNegra.entries()) {
                    if (key.length >= 10 && key.includes(parsed.dni)) {
                        lnData = value;
                        break;
                    }
                }
            }
            
            if (lnData) {
                acceso = 'DENEGADO';
                mensaje = `⛔ ALERTA: ${lnData.motivo || 'LISTA NEGRA'}`;
                tipoPermiso = 'DENEGADO';
                colorClass = 'roja';
            } else {
                // 2. Chequear Nómina (FAP/FAO/Excepciones) en caché local buscando TODOS los permisos del DNI
                let empList = localNominas.get(parsed.dni) || [];
                if (empList.length === 0) {
                    for (const [key, list] of localNominas.entries()) {
                        if (key === parsed.dni || key.includes(parsed.dni) || (parsed.dni.length >= 7 && key.endsWith(parsed.dni))) {
                            empList = list;
                            break;
                        }
                    }
                }

                // 3. Fallback directo a Firestore si aún no sincronizó la caché local
                if (empList.length === 0 && parsed.dni) {
                    try {
                        const directDocs = [];
                        const docDirect = await getDoc(doc(db, "nominas", String(parsed.dni)));
                        if (docDirect.exists()) {
                            directDocs.push({ ...docDirect.data(), docId: docDirect.id });
                        }
                        const qDni = query(collection(db, "nominas"), where("dni", "==", String(parsed.dni)));
                        const snapDni = await getDocs(qDni);
                        snapDni.forEach(d => {
                            if (!directDocs.some(x => x.docId === d.id)) {
                                directDocs.push({ ...d.data(), docId: d.id });
                            }
                        });
                        if (directDocs.length > 0) {
                            localNominas.set(parsed.dni, directDocs);
                            empList = directDocs;
                        }
                    } catch (e) {
                        console.warn("[Verificar] Fallback directo a Firestore falló:", e);
                    }
                }
                
                permisoVencido = false;
                permisoPorVencer = false;
                diasParaVencer = null;
                vencimientoDisplay = '';

                // Evaluar todos los permisos del DNI (si tiene uno vigente y otro vencido, se prioriza el vigente)
                const resPermisos = resolverPermisosPersona(empList);

                if (resPermisos) {
                    const empData = resPermisos.permiso;
                    vencimientoDisplay = resPermisos.vencimientoDisplay;
                    permisoVencido = resPermisos.esVencido;
                    permisoPorVencer = resPermisos.porVencer;
                    diasParaVencer = resPermisos.diasParaVencer;

                    if (resPermisos.estado === 'VENCIDO') {
                        acceso = 'DENEGADO';
                        const cat = empData.categoria || 'PERMISO';
                        const nroTram = empData.nro_tramite ? ` #${empData.nro_tramite}` : '';
                        mensaje = `ACCESO DENEGADO: ${cat}${nroTram} VENCIDO el ${vencimientoDisplay}`;
                        tipoPermiso = `${cat}${nroTram} (VENCIDO)`;
                        colorClass = 'roja';
                        localInfo = empData.puesto_especifico || empData.empresa_nombre || empData.empresa || 'GENERAL';
                        if (empData.nombre) parsed.nombre = empData.nombre;
                    } else if (resPermisos.estado === 'INACTIVO') {
                        acceso = 'DENEGADO';
                        mensaje = 'Empleado Bloqueado / Inactivo';
                        tipoPermiso = empData.categoria || 'N/A';
                        colorClass = 'roja';
                        localInfo = empData.puesto_especifico || empData.empresa_nombre || empData.empresa || 'GENERAL';
                        if (empData.nombre) parsed.nombre = empData.nombre;
                    } else if (resPermisos.estado === 'VIGENTE') {
                        acceso = 'PERMITIDO';
                        if (currentMode === 'salida') {
                            mensaje = 'Salida Registrada';
                            tipoPermiso = 'SALIDA';
                            colorClass = 'gris';
                        } else {
                            const cat = empData.categoria || 'PERSONAL';
                            const nroTram = empData.nro_tramite ? ` #${empData.nro_tramite}` : '';
                            tipoPermiso = `${cat}${nroTram}`;

                            if (permisoPorVencer) {
                                colorClass = 'amarilla';
                                const textoDias = diasParaVencer === 0 ? 'HOY' : (diasParaVencer === 1 ? 'MAÑANA' : `EN ${diasParaVencer} DÍAS`);
                                mensaje = `⚠️ ACCESO PERMITIDO - VENCE ${textoDias}`;
                            } else {
                                colorClass = 'verde';
                                mensaje = `Acceso Personal OK`;
                            }
                        }
                        localInfo = empData.puesto_especifico || empData.empresa_nombre || empData.empresa || 'GENERAL';
                        if (empData.nombre) parsed.nombre = empData.nombre;
                    }
                } else {
                    // No está en nómina. ¿Es visita o salida?
                    if (currentMode === 'visita') {
                        acceso = 'PERMITIDO';
                        mensaje = 'Visita Autorizada';
                        tipoPermiso = 'VISITA';
                        colorClass = 'verde';
                    } else if (currentMode === 'salida') {
                        acceso = 'PERMITIDO';
                        mensaje = 'Salida Registrada';
                        tipoPermiso = 'SALIDA';
                        colorClass = 'gris';
                    } else {
                        acceso = 'DENEGADO';
                        mensaje = 'DNI No Registrado en Nómina';
                        tipoPermiso = 'DESCONOCIDO';
                        colorClass = 'roja';
                    }
                }
            }

            const nowTime = getTimeString();
            const nowDate = getTodayString();
            const fullName = `${parsed.apellido || ''} ${parsed.nombre || ''}`.trim() || 'No Encontrado';
            const evento = currentMode === 'salida' ? 'SALIDA' : 'ENTRADA';

            // Escribir a Firestore (en segundo plano y sanitizado)
            addDoc(collection(db, "accesos"), {
                DNI: String(parsed.dni || ''),
                "Nombre y Apellido": fullName,
                Evento: evento,
                Fecha: nowDate,
                Hora_Ingreso: currentMode === 'salida' ? '' : nowTime,
                Hora_Salida: currentMode === 'salida' ? nowTime : '',
                Resultado: acceso === 'PERMITIDO' ? 'AUTORIZADO' : 'DENEGADO',
                Tipo_Permiso: String(tipoPermiso || 'DESCONOCIDO'),
                Local: String(localInfo || 'GENERAL'),
                Vence: String(vencimientoDisplay || ''),
                Alerta_Vencimiento: Boolean(permisoPorVencer),
                Dias_Vencimiento: (diasParaVencer !== null && diasParaVencer !== undefined) ? diasParaVencer : null,
                Motivo_Rechazo: permisoVencido ? `Permiso vencido el ${vencimientoDisplay}` : '',
                createdAt: serverTimestamp()
            }).catch(error => {
                console.error('Error escribiendo en Firestore:', error);
            });

            // Mantener conteo en vivo
            if (acceso === 'PERMITIDO') {
                if (evento === 'ENTRADA') {
                    setDoc(doc(db, "estado_adentro", String(parsed.dni)), {
                        DNI: String(parsed.dni),
                        nombre: fullName,
                        hora: nowTime,
                        fecha: nowDate,
                        tipo_permiso: String(tipoPermiso || 'DESCONOCIDO'),
                        local: String(localInfo || 'GENERAL')
                    }).catch(e => console.error(e));
                } else if (evento === 'SALIDA') {
                    deleteDoc(doc(db, "estado_adentro", String(parsed.dni))).catch(e => console.error(e));
                }
            }

            const venceInfoTexto = permisoVencido 
                ? `VENCIDO EL ${vencimientoDisplay}` 
                : (permisoPorVencer 
                    ? `⚠️ ATENCIÓN: Vence el ${vencimientoDisplay} (${diasParaVencer === 0 ? 'HOY' : (diasParaVencer === 1 ? 'MAÑANA' : `en ${diasParaVencer} días`)})` 
                    : (vencimientoDisplay ? `Vence: ${vencimientoDisplay}` : ''));

            mostrarResultado(colorClass, mensaje, fullName, localInfo, tipoPermiso, venceInfoTexto);
        } catch (err) {
            console.error("Error al verificar DNI:", err);
            mostrarResultado('roja', 'Error de Verificación', err.message || 'Error inesperado', '', '', '');
        } finally {
            dniInput.focus();
        }
    }


    function mostrarResultado(colorClass, mensaje, nombre, localInfo, tareaInfo, venceInfo) {
        const resultadoDiv = document.getElementById('resultado');
        const mensajeResultado = document.getElementById('mensajeResultado');
        const nombrePersona = document.getElementById('nombrePersona');
        const infoLocal = document.getElementById('infoLocal');
        const infoTarea = document.getElementById('infoTarea');
        const infoVence = document.getElementById('infoVence');

        if (!resultadoDiv || !mensajeResultado) return;

        resultadoDiv.className = `result-display luz-${colorClass || 'roja'}`;
        mensajeResultado.textContent = mensaje || '';
        if (nombrePersona) nombrePersona.textContent = nombre && nombre !== 'No Encontrado' ? `Nombre: ${nombre}` : '';
        if (infoLocal) infoLocal.textContent = '';
        if (infoTarea) infoTarea.textContent = '';
        if (infoVence) infoVence.textContent = '';

        if (colorClass === 'verde' || colorClass === 'amarilla' || colorClass === 'gris' || (nombre && nombre !== 'No Encontrado')) {
            if (infoLocal) infoLocal.textContent = localInfo && localInfo !== 'N/A' ? `${localInfo}` : '';
            if (infoTarea) infoTarea.textContent = tareaInfo && tareaInfo !== 'N/A' ? `${tareaInfo}` : '';
            if (infoVence) infoVence.textContent = venceInfo || '';
        }
    }

    // Listeners principales
    const dniInput = document.getElementById('dniInput');
    if (dniInput) {
        dniInput.focus();

        document.getElementById('btnEntrada')?.addEventListener('click', () => setMode('entrada'));
        document.getElementById('btnRegVisita')?.addEventListener('click', () => setMode('visita'));
        document.getElementById('btnSalida')?.addEventListener('click', () => setMode('salida'));
        
        dniInput.addEventListener('keydown', (event) => {
            if (event.key === 'Enter' || event.keyCode === 13) {
                event.preventDefault();
                verificarDNI();
            }
        });
        
        setMode('entrada');
        setupCacheListeners();
        setupRealtimeListeners();
    }

    function isAnyModalOpen() {
        const overlays = document.querySelectorAll('.modal-overlay, .modal');
        for (const el of overlays) {
            const style = window.getComputedStyle(el);
            if (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0') {
                return true;
            }
        }
        return false;
    }

    // Mantener siempre el foco en el campo DNI para escaneos continuos,
    // EXCEPTO cuando un modal esté abierto en pantalla
    document.addEventListener('click', (e) => {
        if (isAnyModalOpen()) return;
        const tag = e.target ? e.target.tagName.toLowerCase() : '';
        if (tag !== 'input' && tag !== 'textarea' && tag !== 'select' && tag !== 'button' && 
            !e.target.closest('.modal') && !e.target.closest('.modal-content') && !e.target.closest('.modal-overlay')) {
            const input = document.getElementById('dniInput');
            if (input) input.focus();
        }
    });

    document.addEventListener('keydown', (e) => {
        if (isAnyModalOpen()) return;
        const active = document.activeElement;
        const tag = active ? active.tagName.toLowerCase() : '';
        if (tag !== 'input' && tag !== 'textarea' && tag !== 'select') {
            const input = document.getElementById('dniInput');
            if (input) {
                input.focus();
            }
        }
    });

    // --- Modal de Comentarios / Notas ---
    const modalComentario = document.getElementById('modalComentario');
    const closeComentarioModal = document.getElementById('closeComentarioModal');
    const btnGuardarComentario = document.getElementById('btnGuardarComentario');
    const comentarioTexto = document.getElementById('comentarioTexto');
    
    window.openComentarioModal = function(record, evt) {
        if (evt && typeof evt.stopPropagation === 'function') {
            evt.stopPropagation();
        }
        if (!record || !record.id) return;
        document.getElementById('comentarioContexto').textContent = `DNI: ${record.DNI} - ${record['Nombre y Apellido'] || ''}`;
        document.getElementById('comentarioAccesoId').value = record.id;
        comentarioTexto.value = record.Comentario || '';
        modalComentario.style.display = 'flex';
        // Posicionar el foco garantizado en el textarea sin que dniInput lo robe
        setTimeout(() => {
            if (comentarioTexto) {
                comentarioTexto.focus();
                comentarioTexto.select();
            }
        }, 60);
    };

    if (closeComentarioModal) {
        closeComentarioModal.onclick = () => {
            modalComentario.style.display = 'none';
            const input = document.getElementById('dniInput');
            if (input) input.focus();
        };
    }

    window.addEventListener('click', (e) => {
        if (e.target == modalComentario) {
            modalComentario.style.display = 'none';
            const input = document.getElementById('dniInput');
            if (input) input.focus();
        }
    });
    
    if (btnGuardarComentario) {
        btnGuardarComentario.onclick = async () => {
            const docId = document.getElementById('comentarioAccesoId').value;
            const nuevoComentario = comentarioTexto.value.trim();
            if (!docId) return;
            
            btnGuardarComentario.disabled = true;
            btnGuardarComentario.textContent = 'Guardando...';
            
            try {
                const docRef = doc(db, 'accesos', docId);
                await updateDoc(docRef, { Comentario: nuevoComentario });
                modalComentario.style.display = 'none';
                const input = document.getElementById('dniInput');
                if (input) input.focus();
            } catch (e) {
                console.error("Error guardando comentario:", e);
                alert("Error al guardar el comentario.");
            } finally {
                btnGuardarComentario.disabled = false;
                btnGuardarComentario.textContent = 'Guardar Comentario';
            }
        };
    }

    // --- LOGICA CIERRE DE JORNADA / TURNO ---
    const btnCierreJornada = document.getElementById('btnCierreJornada');
    const modalCierreJornada = document.getElementById('modalCierreJornada');
    const closeCierreModal = document.getElementById('closeCierreModal');
    const btnCancelarCierre = document.getElementById('btnCancelarCierre');
    const btnConfirmarCierre = document.getElementById('btnConfirmarCierre');
    const cierreCantAdentro = document.getElementById('cierreCantAdentro');
    const cierreObservacion = document.getElementById('cierreObservacion');

    if (btnCierreJornada && modalCierreJornada) {
        btnCierreJornada.onclick = async () => {
            // Consultar cuántos hay y abrir modal
            try {
                const snap = await getDocs(collection(db, "estado_adentro"));
                if (cierreCantAdentro) cierreCantAdentro.textContent = snap.size;
            } catch (e) {
                console.error("Error consultando estado_adentro", e);
            }
            if (cierreObservacion) cierreObservacion.value = '';
            modalCierreJornada.style.display = 'flex';
        };

        const cerrarModalCierre = () => {
            modalCierreJornada.style.display = 'none';
        };

        if (closeCierreModal) closeCierreModal.onclick = cerrarModalCierre;
        if (btnCancelarCierre) btnCancelarCierre.onclick = cerrarModalCierre;

        window.addEventListener('click', (e) => {
            if (e.target == modalCierreJornada) cerrarModalCierre();
        });

        if (btnConfirmarCierre) {
            btnConfirmarCierre.onclick = async () => {
                btnConfirmarCierre.disabled = true;
                btnConfirmarCierre.textContent = 'Purgando y guardando...';

                try {
                    // 1. Obtener todas las personas actualmente adentro
                    const snap = await getDocs(collection(db, "estado_adentro"));
                    const personasAdentro = [];
                    const deletePromises = [];

                    snap.forEach(docSnap => {
                        const d = docSnap.data();
                        personasAdentro.push({
                            dni: d.DNI || docSnap.id,
                            nombre: d.nombre || 'Desconocido',
                            hora_ingreso: d.hora || '',
                            fecha_ingreso: d.fecha || '',
                            tipo_permiso: d.tipo_permiso || '',
                            local: d.local || ''
                        });
                        deletePromises.push(deleteDoc(doc(db, "estado_adentro", docSnap.id)));
                    });

                    // 2. Guardar el reporte de auditoría en la colección cierres_jornada
                    const currentUser = auth.currentUser;
                    await addDoc(collection(db, "cierres_jornada"), {
                        fecha: getTodayString(),
                        hora: getTimeString(),
                        operario_email: currentUser ? currentUser.email : 'Garita Operador',
                        operario_uid: currentUser ? currentUser.uid : '',
                        total_quedaron_adentro: personasAdentro.length,
                        observacion: cierreObservacion ? cierreObservacion.value.trim() : '',
                        personas_adentro: personasAdentro,
                        createdAt: serverTimestamp()
                    });

                    // 3. Purgar la colección estado_adentro
                    await Promise.all(deletePromises);

                    alert(`✅ Cierre de jornada completado con éxito.\nQuedaron registradas ${personasAdentro.length} personas para auditoría.`);
                    cerrarModalCierre();
                } catch (e) {
                    console.error("Error al procesar cierre de jornada:", e);
                    alert("Error al procesar el cierre de jornada. Verifique la conexión.");
                } finally {
                    btnConfirmarCierre.disabled = false;
                    btnConfirmarCierre.textContent = 'Confirmar y Purgar';
                }
            };
        }
    }
}

// Inicialización robusta para ES Modules (evita condiciones de carrera si DOMContentLoaded ya disparó)
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initScanner);
} else {
    initScanner();
}
// --- PWA INSTALL LOGIC ---
let deferredPrompt;
const btnInstalar = document.getElementById('btnInstalarPWA');

window.addEventListener('beforeinstallprompt', (e) => {
    // Prevent Chrome 67 and earlier from automatically showing the prompt
    e.preventDefault();
    // Stash the event so it can be triggered later.
    deferredPrompt = e;
    // Update UI to notify the user they can add to home screen
    if(btnInstalar) btnInstalar.style.display = 'block';
});

if(btnInstalar) {
    btnInstalar.addEventListener('click', async () => {
        if (deferredPrompt) {
            // Show the install prompt
            deferredPrompt.prompt();
            // Wait for the user to respond to the prompt
            const { outcome } = await deferredPrompt.userChoice;
            console.log(`User response to the install prompt: ${outcome}`);
            // We've used the prompt, and can't use it again, throw it away
            deferredPrompt = null;
            btnInstalar.style.display = 'none';
        }
    });
}
window.addEventListener('appinstalled', () => {
    // Hide the app-provided install promotion
    if(btnInstalar) btnInstalar.style.display = 'none';
    deferredPrompt = null;
    console.log('PWA was installed');
});
