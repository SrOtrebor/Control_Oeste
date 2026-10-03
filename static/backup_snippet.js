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
