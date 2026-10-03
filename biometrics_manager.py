import logging
import base64
import uuid

logger = logging.getLogger(__name__)

class BiometricsManager:
    def __init__(self):
        self.device_connected = False
        self.scanner_model = "SecuGen Hamster Plus (HSDU03P)"
        self.dpi = 500

    def init_device(self):
        """Inicializa la conexión con el driver de SecuGen usando ctypes.
        Actualmente es un STUB (simulador)."""
        logger.info(f"Conectando a lector biométrico: {self.scanner_model}")
        self.device_connected = True
        return True

    def capture_fingerprint(self):
        """Captura una huella, extrae las minucias y retorna el template base64."""
        if not self.device_connected:
            self.init_device()
            
        logger.info("Esperando que el usuario coloque el dedo en el sensor...")
        
        # --- STUB: Simulación de escaneo exitoso ---
        # En el futuro, aquí usaremos `sgfplib.dll` para `GetImage` y `CreateTemplate`.
        
        # Simulamos un delay de captura y procesamiento
        import time
        time.sleep(1.5)
        
        # Generamos un ID falso que actua como "Template" de la huella
        fake_minutiae_data = f"SECUGEN_TEMPLATE_V1_{uuid.uuid4().hex}".encode('utf-8')
        template_base64 = base64.b64encode(fake_minutiae_data).decode('utf-8')
        
        return {
            'success': True,
            'template': template_base64,
            'quality': 98,
            'message': 'Huella capturada y procesada con éxito'
        }

    def verify_fingerprint(self, capture_template, db_template):
        """Compara la huella actual con una guardada en BD.
        Retorna (match_score, is_match)."""
        # --- STUB ---
        # A futuro: Usar `MatchTemplate` de la API de SecuGen.
        return capture_template == db_template

biometrics_manager = BiometricsManager()
