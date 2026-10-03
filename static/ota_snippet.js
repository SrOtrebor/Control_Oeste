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
