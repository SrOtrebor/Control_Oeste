import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-app.js";
import { getAuth, signInWithEmailAndPassword, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-auth.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/10.8.1/firebase-firestore.js";

// Tu configuración de Firebase
const firebaseConfig = {
  apiKey: "AIzaSyBg0ht9KWgGrhGmHkT7mzCRgJ2CS6e4-lQ",
  authDomain: "control-acceso-c2f82.firebaseapp.com",
  projectId: "control-acceso-c2f82",
  storageBucket: "control-acceso-c2f82.firebasestorage.app",
  messagingSenderId: "654089830879",
  appId: "1:654089830879:web:96b5a40c8fc6292422446a"
};

// Inicializar Firebase
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

// Elementos del DOM
const loginForm = document.getElementById('loginForm');
const emailInput = document.getElementById('email');
const passwordInput = document.getElementById('password');
const btnText = document.getElementById('btnText');
const btnLoader = document.getElementById('btnLoader');
const loginBtn = document.getElementById('loginBtn');
const errorMessage = document.getElementById('errorMessage');

// Verificar si ya está logueado
onAuthStateChanged(auth, async (user) => {
    if (user) {
        // Usuario logueado, buscar su rol y redirigir
        await checkRoleAndRedirect(user.uid);
    }
});

// Manejar Login
loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    // UI - Cargando
    btnText.style.display = 'none';
    btnLoader.style.display = 'block';
    loginBtn.disabled = true;
    errorMessage.style.display = 'none';
    
    try {
        const userCredential = await signInWithEmailAndPassword(auth, emailInput.value, passwordInput.value);
        await checkRoleAndRedirect(userCredential.user.uid);
    } catch (error) {
        console.error("Error logging in:", error);
        
        let msg = "Error al iniciar sesión.";
        if (error.code === 'auth/invalid-credential' || error.code === 'auth/wrong-password' || error.code === 'auth/user-not-found') {
            msg = "Correo o contraseña incorrectos.";
        } else if (error.code === 'auth/too-many-requests') {
            msg = "Demasiados intentos. Intente más tarde.";
        }
        
        errorMessage.textContent = msg;
        errorMessage.style.display = 'block';
        
        // Restaurar UI
        btnText.style.display = 'block';
        btnLoader.style.display = 'none';
        loginBtn.disabled = false;
    }
});

// Función para leer el rol en Firestore y redirigir
async function checkRoleAndRedirect(uid) {
    try {
        const docRef = doc(db, "usuarios", uid);
        const docSnap = await getDoc(docRef);
        
        if (docSnap.exists()) {
            const userData = docSnap.data();
            
            if (!userData.activo) {
                errorMessage.textContent = "Su usuario ha sido bloqueado. Contacte al administrador.";
                errorMessage.style.display = 'block';
                auth.signOut();
                restoreBtn();
                return;
            }

            // Redirección según rol
            switch(userData.rol) {
                case 'super_admin':
                    window.location.href = '/maestro.html';
                    break;
                case 'supervisor':
                    window.location.href = '/dashboard.html';
                    break;
                case 'acceso_principal':
                case 'acceso_satelite':
                    window.location.href = '/scan.html';
                    break;
                default:
                    errorMessage.textContent = "Rol no reconocido.";
                    errorMessage.style.display = 'block';
                    auth.signOut();
                    restoreBtn();
            }
        } else {
            // El usuario existe en Auth pero no tiene documento de rol
            errorMessage.textContent = "Error: Usuario sin rol asignado en la base de datos.";
            errorMessage.style.display = 'block';
            auth.signOut();
            restoreBtn();
        }
    } catch (error) {
        console.error("Error buscando rol:", error);
        errorMessage.textContent = "Error al verificar permisos. Revise la consola.";
        errorMessage.style.display = 'block';
        auth.signOut();
        restoreBtn();
    }
}

function restoreBtn() {
    btnText.style.display = 'block';
    btnLoader.style.display = 'none';
    loginBtn.disabled = false;
}
