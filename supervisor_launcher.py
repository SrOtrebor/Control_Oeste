import socket
import webbrowser
import time
import sys
import os
import tkinter as tk
from tkinter import messagebox

UDP_PORT = 5005
MAGIC_REQUEST = b"WHERE_IS_GARITA"
MAGIC_RESPONSE = b"I_AM_GARITA"

def find_garita():
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
    # Timeout de 3 segundos para no quedarse esperando por siempre
    sock.settimeout(3.0)
    
    try:
        # Enviar broadcast al puerto 5005
        sock.sendto(MAGIC_REQUEST, ('<broadcast>', UDP_PORT))
        # O como alternativa en Windows: sock.sendto(MAGIC_REQUEST, ('255.255.255.255', UDP_PORT))
        
        while True:
            data, addr = sock.recvfrom(1024)
            if data.startswith(MAGIC_RESPONSE):
                parts = data.decode().split(':')
                if len(parts) >= 2:
                    port = parts[1]
                    return f"http://{addr[0]}:{port}/dashboard"
    except socket.timeout:
        return None
    except Exception as e:
        return None
    finally:
        sock.close()

def main():
    # Ocultar ventana principal de tkinter
    root = tk.Tk()
    root.withdraw()
    
    url = find_garita()
    if url:
        webbrowser.open(url)
        sys.exit(0)
    
    # Intento 2 con 255.255.255.255
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
    sock.settimeout(3.0)
    try:
        sock.sendto(MAGIC_REQUEST, ('255.255.255.255', UDP_PORT))
        data, addr = sock.recvfrom(1024)
        if data.startswith(MAGIC_RESPONSE):
            parts = data.decode().split(':')
            if len(parts) >= 2:
                port = parts[1]
                url = f"http://{addr[0]}:{port}/dashboard"
                webbrowser.open(url)
                sys.exit(0)
    except:
        pass
    finally:
        sock.close()

    messagebox.showerror("Error de Conexión", 
        "No se pudo encontrar el servidor de la Garita en la red local.\n\n"
        "Verifique que la computadora de la garita esté encendida y la aplicación 'Control de Acceso' esté abierta.")
    sys.exit(1)

if __name__ == "__main__":
    main()
