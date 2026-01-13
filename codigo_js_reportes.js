// Código JavaScript para agregar al final de script.js
// Manejo de descarga de reportes con fechas

// Descargar Reporte de Accesos
if (document.getElementById('descargarReporteBtn')) {
    document.getElementById('descargarReporteBtn').addEventListener('click', function() {
        const fechaDesde = document.getElementById('fechaDesdeReporte').value;
        const fechaHasta = document.getElementById('fechaHastaReporte').value;
        
        let url = '/descargar_reporte_fechas?';
        if (fechaDesde) url += `desde=${fechaDesde}&`;
        if (fechaHasta) url += `hasta=${fechaHasta}`;
        
        // Abrir en nueva ventana para descargar
        window.open(url, '_blank');
        
        const statusEl = document.getElementById('reporteStatus');
        if (statusEl) {
            if (fechaDesde && fechaHasta && fechaDesde !== fechaHasta) {
                statusEl.textContent = `Descargando reporte del ${fechaDesde} al ${fechaHasta}...`;
            } else if (fechaDesde || fechaHasta) {
                const fecha = fechaDesde || fechaHasta;
                statusEl.textContent = `Descargando reporte del ${fecha}...`;
            } else {
                statusEl.textContent = 'Descargando reporte de hoy...';
            }
            statusEl.style.color = '#4CAF50';
        }
    });
}

// Descargar Reporte de Fichajes
if (document.getElementById('descargarFichajesBtn')) {
    document.getElementById('descargarFichajesBtn').addEventListener('click', function() {
        const fechaDesde = document.getElementById('fechaDesdeReporte').value;
        const fechaHasta = document.getElementById('fechaHastaReporte').value;
        
        let url = '/descargar_fichajes_fechas?';
        if (fechaDesde) url += `desde=${fechaDesde}&`;
        if (fechaHasta) url += `hasta=${fechaHasta}`;
        
        // Abrir en nueva ventana para descargar
        window.open(url, '_blank');
        
        const statusEl = document.getElementById('reporteStatus');
        if (statusEl) {
            if (fechaDesde && fechaHasta && fechaDesde !== fechaHasta) {
                statusEl.textContent = `Descargando fichajes del ${fechaDesde} al ${fechaHasta}...`;
            } else if (fechaDesde || fechaHasta) {
                const fecha = fechaDesde || fechaHasta;
                statusEl.textContent = `Descargando fichajes del ${fecha}...`;
            } else {
                statusEl.textContent = 'Descargando fichajes de hoy...';
            }
            statusEl.style.color = '#4CAF50';
        }
    });
}
