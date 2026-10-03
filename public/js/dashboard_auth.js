import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getAuth, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-auth.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-firestore.js";

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

// Proteger ruta
onAuthStateChanged(auth, async (user) => {
    if (!user) {
        window.location.href = '/';
        return;
    }
    
    // Verificar que sea supervisor o admin
    const docRef = doc(db, "usuarios", user.uid);
    const docSnap = await getDoc(docRef);
    if (!docSnap.exists()) {
        window.location.href = '/';
        return;
    }
    
    const rol = docSnap.data().rol;
    if (rol !== 'supervisor' && rol !== 'super_admin' && rol !== 'coi') {
        window.location.href = '/scan.html';
        return;
    }
});

// Logout listener
document.addEventListener('DOMContentLoaded', () => {
    const btnSalir = document.querySelector('a[href="/logout"]');
    if (btnSalir) {
        btnSalir.addEventListener('click', (e) => {
            e.preventDefault();
            signOut(auth).then(() => {
                window.location.href = '/';
            });
        });
    }
});
