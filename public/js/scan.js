import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getAuth, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-auth.js";
import { getFirestore, collection, addDoc, query, where, getDocs, onSnapshot, serverTimestamp, doc, getDoc, updateDoc, enableIndexedDbPersistence } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-firestore.js";

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

document.addEventListener('DOMContentLoaded', () => {

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
    const localNominas = new Map();
    const localListaNegra = new Map();

    function setupCacheListeners() {
        onSnapshot(collection(db, "nominas"), (snapshot) => {
            localNominas.clear();
            snapshot.forEach(doc => {
                const data = doc.data();
                localNominas.set(data.dni || doc.id, data);
            });
        });
        onSnapshot(collection(db, "lista_negra"), (snapshot) => {
            localListaNegra.clear();
            snapshot.forEach(doc => localListaNegra.set(doc.id, doc.data()));
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
            headerTitle.textContent = 'Salida de Visita';
        }

        if (dniInput) dniInput.focus();
    }

    // Configurar Listeners en tiempo real para Stats y Tabla Diaria
    function setupRealtimeListeners() {
        // Escuchar estadísticas en vivo para Total Adentro
        onSnapshot(doc(db, "config", "live_stats"), (docSnap) => {
            if (docSnap.exists()) {
                const data = docSnap.data();
                if (data && data.total_adentro !== undefined) {
                    const el = document.getElementById('statTotalAdentro');
                    if (el) el.textContent = data.total_adentro;
                }
            }
        });

        const today = getTodayString();
        const q = query(collection(db, "accesos"), where("Fecha", "==", today));
        
        onSnapshot(q, (snapshot) => {
            const ingresosListBody = document.getElementById('ingresosListBody');
            if (ingresosListBody) ingresosListBody.innerHTML = '';
            
            let permitidos = 0;
            let rechazados = 0;
            const records = [];

            snapshot.forEach(doc => {
                let d = doc.data();
                d.id = doc.id;
                records.push(d);
            });
            
            // Ordenar descendente por hora
            records.sort((a, b) => {
                const timeA = a.Hora_Ingreso || a.Hora_Salida || '00:00:00';
                const timeB = b.Hora_Ingreso || b.Hora_Salida || '00:00:00';
                return timeB.localeCompare(timeA);
            });

            records.forEach(record => {
                const res = (record.Resultado || '').toUpperCase();
                if (res.includes('AUTORIZADO')) permitidos++;
                else if (res.includes('DENEGADO')) rechazados++;

                // Llenar tabla visual
                if (ingresosListBody) {
                    const row = ingresosListBody.insertRow();
                    if (res.includes('AUTORIZADO') || res.includes('PERMITIDO')) {
                        if (record.Tipo_Permiso === 'FAO') row.classList.add('fila-verde', 'fila-fao');
                        else if (record.Tipo_Permiso === 'FAP') row.classList.add('fila-verde');
                        else if (record.Tipo_Permiso === 'PERSONAL') row.classList.add('fila-gris');
                        else if (record.Tipo_Permiso === 'VISITA') row.classList.add('fila-gris');
                        else if (record.Tipo_Permiso === 'FICHAJE') row.classList.add('fila-gris');
                        else row.classList.add('fila-verde');
                    } else {
                        row.classList.add('fila-roja');
                    }

                    row.style.cursor = 'pointer';
                    row.title = 'Clic para agregar comentario/observación';
                    row.onclick = () => openComentarioModal(record);

                    const timeString = record.Hora_Ingreso || record.Hora_Salida || '';
                    row.insertCell().textContent = timeString;
                    row.insertCell().textContent = record.DNI || '';
                    
                    let commentHtml = record.Comentario ? `<br><small style="font-style: italic; color: #fbbf24;">📝 ${record.Comentario}</small>` : '';
                    row.insertCell().innerHTML = (record['Nombre y Apellido'] || '') + commentHtml;
                    
                    row.insertCell().textContent = record.Evento || '';
                    row.insertCell().textContent = record.Tipo_Permiso || '';
                    row.insertCell().textContent = record.Local || '';
                }
            });

            const statPerm = document.getElementById('statPermitidos');
            const statRech = document.getElementById('statRechazados');
            if(statPerm) statPerm.textContent = permitidos;
            if(statRech) statRech.textContent = rechazados;
        });
    }

    async function verificarDNI() {
        const dniInput = document.getElementById('dniInput');
        const rawData = dniInput.value.trim();
        if (!rawData) return;
        dniInput.value = ''; // limpiar rapido

        const parsed = parseDNI(rawData);
        if (!parsed) return;
        
        let acceso = 'DENEGADO';
        let mensaje = 'Acceso Denegado';
        let tipoPermiso = 'DESCONOCIDO';
        let colorClass = 'roja';
        let localInfo = 'GENERAL';
        
        // Consultar Caché Local primero (por DNI exacto o CUIL)
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
            // Si no está en lista negra, chequear Nómina (FAP/FAO)
            let empData = localNominas.get(parsed.dni);
            if (!empData) {
                for (const [key, value] of localNominas.entries()) {
                    if ((key && key.includes(parsed.dni)) || (value.dni && value.dni.includes(parsed.dni))) {
                        empData = value;
                        break;
                    }
                }
            }
            
            if (empData) {
                if (empData.activo === false) {
                    acceso = 'DENEGADO';
                    mensaje = 'Empleado Bloqueado / Inactivo';
                    tipoPermiso = empData.categoria || 'N/A';
                    colorClass = 'roja';
                    localInfo = empData.puesto_especifico || empData.empresa_nombre || 'GENERAL';
                } else {
                    acceso = 'PERMITIDO';
                    mensaje = `Acceso Personal OK`;
                    tipoPermiso = empData.categoria || 'PERSONAL';
                    colorClass = 'verde';
                    localInfo = empData.puesto_especifico || empData.empresa_nombre || 'GENERAL';
                    // Reemplazamos el nombre parseado por el oficial de la base
                    if (empData.nombre) parsed.nombre = empData.nombre;
                }
            } else {
                // No está en nómina. ¿Es visita?
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
        const fullName = `${parsed.apellido} ${parsed.nombre}`.trim();
        const evento = currentMode === 'salida' ? 'SALIDA' : 'ENTRADA';

        // Escribir a Firestore (Background, sin await para que UI responda instantáneamente)
        addDoc(collection(db, "accesos"), {
            DNI: parsed.dni,
            "Nombre y Apellido": fullName || 'No Encontrado',
            Evento: evento,
            Fecha: nowDate,
            Hora_Ingreso: currentMode === 'salida' ? '' : nowTime,
            Hora_Salida: currentMode === 'salida' ? nowTime : '',
            Resultado: acceso === 'PERMITIDO' ? 'AUTORIZADO' : 'DENEGADO',
            Tipo_Permiso: tipoPermiso,
            Local: localInfo,
            createdAt: serverTimestamp()
        }).catch(error => {
            console.error('Error escribiendo en Firestore:', error);
        });

        mostrarResultado(colorClass, mensaje, fullName, localInfo, tipoPermiso, '');
    }

    function mostrarResultado(colorClass, mensaje, nombre, localInfo, tareaInfo, venceInfo) {
        const resultadoDiv = document.getElementById('resultado');
        const mensajeResultado = document.getElementById('mensajeResultado');
        const nombrePersona = document.getElementById('nombrePersona');
        const infoLocal = document.getElementById('infoLocal');
        const infoTarea = document.getElementById('infoTarea');
        const infoVence = document.getElementById('infoVence');

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

    // Listeners principales
    const dniInput = document.getElementById('dniInput');
    if (dniInput) {
        document.getElementById('btnEntrada')?.addEventListener('click', () => setMode('entrada'));
        document.getElementById('btnRegVisita')?.addEventListener('click', () => setMode('visita'));
        document.getElementById('btnSalida')?.addEventListener('click', () => setMode('salida'));
        
        dniInput.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                verificarDNI();
            }
        });
        
        setMode('entrada');
        setupCacheListeners();
        setupRealtimeListeners();
    }

    // --- Modal de Comentarios ---
    const modalComentario = document.getElementById('modalComentario');
    const closeComentarioModal = document.getElementById('closeComentarioModal');
    const btnGuardarComentario = document.getElementById('btnGuardarComentario');
    const comentarioTexto = document.getElementById('comentarioTexto');
    
    window.openComentarioModal = function(record) {
        if (!record || !record.id) return;
        document.getElementById('comentarioContexto').textContent = `DNI: ${record.DNI} - ${record['Nombre y Apellido'] || ''}`;
        document.getElementById('comentarioAccesoId').value = record.id;
        comentarioTexto.value = record.Comentario || '';
        modalComentario.style.display = 'flex';
        comentarioTexto.focus();
    };

    if (closeComentarioModal) {
        closeComentarioModal.onclick = () => {
            modalComentario.style.display = 'none';
        };
    }

    window.addEventListener('click', (e) => {
        if (e.target == modalComentario) {
            modalComentario.style.display = 'none';
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
            } catch (e) {
                console.error("Error guardando comentario:", e);
                alert("Error al guardar el comentario.");
            } finally {
                btnGuardarComentario.disabled = false;
                btnGuardarComentario.textContent = 'Guardar Comentario';
            }
        };
    }
});
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
