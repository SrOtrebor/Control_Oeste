// --- SSE ALERT LISTENER ---
window.activeAlerts = [];

function initSSEAlerts() {
    const eventSource = new EventSource('/stream/alerts');
    const audio = document.getElementById('alertSound');
    const statAlertas = document.getElementById('statAlertas');
    const cardAlertas = document.getElementById('cardAlertas');

    eventSource.onmessage = function(event) {
        try {
            const data = JSON.parse(event.data);
            console.log("ALERTA RECIBIDA:", data);
            
            data.timestamp = new Date();
            window.activeAlerts.push(data);
            
            // Actualizar UI
            if (statAlertas) {
                statAlertas.textContent = window.activeAlerts.length;
            }
            if (cardAlertas) {
                cardAlertas.classList.add('alert-blinking');
            }
            
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
}

document.addEventListener('DOMContentLoaded', () => {
    setTimeout(initSSEAlerts, 1000);
});

window.mostrarHistorialAlertas = function() {
    const cardAlertas = document.getElementById('cardAlertas');
    const statAlertas = document.getElementById('statAlertas');
    const audio = document.getElementById('alertSound');

    if (window.activeAlerts.length === 0) {
        alert("No hay alertas activas.");
        return;
    }

    // Detener parpadeo y sonido
    if (cardAlertas) cardAlertas.classList.remove('alert-blinking');
    if (audio) {
        audio.pause();
        audio.currentTime = 0;
    }
    
    // Armar lista
    let msg = "ÚLTIMAS ALERTAS:\n\n";
    window.activeAlerts.forEach((a, idx) => {
        msg += `${idx + 1}. ${a.timestamp.toLocaleTimeString()} - DNI: ${a.dni} - Nombre: ${a.nombre || 'Desconocido'}\nMotivo: ${a.motivo}\n\n`;
    });
    
    alert(msg);
    
    // Marcar como leídas (limpiar)
    window.activeAlerts = [];
    if (statAlertas) statAlertas.textContent = "0";
};
