import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getFirestore, collection, addDoc, doc, setDoc } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-firestore.js";

const firebaseConfig = {
    apiKey: "AIzaSyBg0ht9KWgGrhGmHkT7mzCRgJ2CS6e4-lQ",
    authDomain: "control-acceso-c2f82.firebaseapp.com",
    projectId: "control-acceso-c2f82",
    storageBucket: "control-acceso-c2f82.firebasestorage.app",
    messagingSenderId: "654089830879",
    appId: "1:654089830879:web:96b5a40c8fc6292422446a"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

const logEl = document.getElementById('log');

function log(msg, type='info') {
    const span = document.createElement('span');
    span.className = type;
    span.textContent = msg;
    logEl.appendChild(span);
    logEl.appendChild(document.createElement('br'));
    logEl.scrollTop = logEl.scrollHeight;
}

document.getElementById('btnMigrarListaNegra').addEventListener('click', async () => {
    try {
        log('Obteniendo Lista Negra del JSON local...');
        const res = await fetch('/migracion.json');
        const data = await res.json();
        
        const registros = data.lista_negra || [];
        log(`Se encontraron ${registros.length} registros. Subiendo a Firestore...`);
        
        for (const r of registros) {
            const docRef = doc(db, 'lista_negra', String(r.dni));
            await setDoc(docRef, {
                dni: String(r.dni),
                nombre: r.nombre || '',
                motivo: r.motivo || 'Ingresado por sistema viejo'
            });
            log(`+ Subido DNI ${r.dni}`, 'success');
        }
        log('✅ Migración de Lista Negra completa.', 'success');
    } catch (e) {
        log('❌ Error: ' + e.message, 'error');
    }
});

document.getElementById('btnMigrarNominas').addEventListener('click', async () => {
    try {
        log('Obteniendo Nóminas (FAO/FAP) del JSON local...');
        const res = await fetch('/migracion.json');
        const data = await res.json();
        
        const registros = data.nominas || [];
        log(`Se encontraron ${registros.length} personas en nómina. Subiendo a Firestore...`);
        
        for (const r of registros) {
            const docRef = doc(db, 'nominas', String(r.dni));
            await setDoc(docRef, {
                dni: String(r.dni),
                nombre: r.nombre || '',
                categoria: r.categoria || 'N/A',
                puesto_especifico: r.puesto_especifico || '',
                activo: r.activo === undefined ? true : r.activo
            });
            log(`+ Subido DNI ${r.dni} (${r.categoria})`, 'success');
        }
        log('✅ Migración de Nóminas completa.', 'success');
    } catch (e) {
        log('❌ Error: ' + e.message, 'error');
    }
});
