import socket
import threading
import logging
import time

logger = logging.getLogger(__name__)

UDP_IP = "0.0.0.0"
UDP_PORT = 5005
MAGIC_REQUEST = b"WHERE_IS_GARITA"
MAGIC_RESPONSE = b"I_AM_GARITA"

def start_udp_beacon(port=5000):
    """
    Inicia un servidor UDP que escucha solicitudes broadcast en la red local.
    Cuando recibe MAGIC_REQUEST, responde con su IP y el puerto del servidor web.
    """
    def _listen():
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        # En Windows a veces necesitamos esto para broadcast:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
        
        try:
            sock.bind((UDP_IP, UDP_PORT))
            logger.info(f"Faro UDP iniciado en puerto {UDP_PORT}")
        except Exception as e:
            logger.error(f"Falla al iniciar Faro UDP: {e}")
            return

        while True:
            try:
                data, addr = sock.recvfrom(1024)
                if data.strip() == MAGIC_REQUEST:
                    # Enviar respuesta con el puerto del servidor Flask
                    response = f"{MAGIC_RESPONSE.decode()}:{port}".encode()
                    sock.sendto(response, addr)
            except Exception as e:
                logger.error(f"Error en Faro UDP: {e}")
                time.sleep(1)

    thread = threading.Thread(target=_listen, daemon=True)
    thread.start()
