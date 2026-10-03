import os
import json
import base64
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from dotenv import load_dotenv

load_dotenv()

IRSA_CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'irsa_config.json')

def _get_fernet():
    """Deriva una clave de encriptación a partir del SECRET_KEY de Flask"""
    secret = os.environ.get('SECRET_KEY', 'default-dev-secret-key').encode()
    salt = b'irsa_salt_1234'
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )
    key = base64.urlsafe_b64encode(kdf.derive(secret))
    return Fernet(key)

def save_irsa_credentials(username, password, shopping_id="61709", dias_alerta_vencimiento=0):
    """Guarda las credenciales encriptadas localmente"""
    f = _get_fernet()
    enc_password = f.encrypt(password.encode()).decode()
    
    config = {
        'username': username,
        'password_enc': enc_password,
        'shopping_id': shopping_id,
        'dias_alerta_vencimiento': dias_alerta_vencimiento
    }
    
    with open(IRSA_CONFIG_FILE, 'w') as file:
        json.dump(config, file)
        
    return True

def get_irsa_credentials():
    """Devuelve las credenciales desencriptadas. Retorna None si no existen."""
    if not os.path.exists(IRSA_CONFIG_FILE):
        return None
        
    try:
        with open(IRSA_CONFIG_FILE, 'r') as file:
            config = json.load(file)
            
        f = _get_fernet()
        dec_password = f.decrypt(config['password_enc'].encode()).decode()
        
        return {
            'username': config.get('username'),
            'password': dec_password,
            'shopping_id': config.get('shopping_id', '61709'),
            'dias_alerta_vencimiento': int(config.get('dias_alerta_vencimiento', 0))
        }
    except Exception as e:
        print(f"Error reading IRSA config: {e}")
        return None

def has_irsa_credentials():
    """Devuelve True si hay credenciales configuradas"""
    return os.path.exists(IRSA_CONFIG_FILE)
