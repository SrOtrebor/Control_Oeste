import queue
import time
import json
import logging
from flask import Response

logger = logging.getLogger(__name__)

class AlertManager:
    def __init__(self):
        self.listeners = []

    def listen(self):
        q = queue.Queue(maxsize=5)
        self.listeners.append(q)
        return q

    def announce(self, msg_dict):
        """Envía una alerta a todos los clientes conectados"""
        msg = json.dumps(msg_dict)
        for i in reversed(range(len(self.listeners))):
            try:
                self.listeners[i].put_nowait(msg)
            except queue.Full:
                del self.listeners[i]

    def remove_listener(self, q):
        try:
            self.listeners.remove(q)
        except ValueError:
            pass

alert_manager = AlertManager()

def send_alert(dni, nombre, motivo, nivel="danger"):
    """
    Despacha la alerta.
    nivel: 'danger' (Rojo), 'warning' (Amarillo)
    """
    logger.info(f"🚨 ALERTA DISPARADA: {dni} - {nombre} - {motivo}")
    alert_manager.announce({
        'dni': dni,
        'nombre': nombre,
        'motivo': motivo,
        'nivel': nivel,
        'timestamp': time.time()
    })

def stream_alerts():
    """Generador para Server-Sent Events (SSE)"""
    q = alert_manager.listen()
    try:
        while True:
            msg = q.get(timeout=10) # Envia ping cada 10s para mantener conexión viva
            yield f"data: {msg}\n\n"
    except queue.Empty:
        # Keep-alive
        yield ": keep-alive\n\n"
        yield from stream_alerts()
    except GeneratorExit:
        alert_manager.remove_listener(q)
