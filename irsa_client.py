import json
import os
import requests
import pandas as pd
from datetime import datetime, timedelta
import time
import urllib3
import logging

from config import BASE_DIR
from config import EXCEL_FAP, EXCEL_FAO
from irsa_config import get_irsa_credentials

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = logging.getLogger(__name__)

class IRSASyncError(Exception):
    pass

class IRSAAuthError(IRSASyncError):
    pass

class IRSAClient:
    BASE_URL = "https://locatarios.irsa.com.ar"
    
    def __init__(self):
        self.session = requests.Session()
        # Fake user agent
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'application/json, text/plain, */*'
        })
        
    def login(self, username, password):
        """Autenticarse en la API REST del portal de IRSA (SPA React)"""
        login_url = f"{self.BASE_URL}/api/login"
        
        payload = {
            'userName': username,
            'password': password
        }
        
        try:
            res = self.session.post(login_url, json=payload, verify=False, timeout=15)
            
            if res.ok:
                body = res.json()
                if body.get('success'):
                    logger.info(f"Login IRSA exitoso para usuario: {username}")
                    return True
                else:
                    raise IRSAAuthError("Credenciales inválidas o usuario deshabilitado")
            else:
                raise IRSAAuthError(f"Error HTTP {res.status_code} durante el login")
        except IRSAAuthError:
            raise
        except Exception as e:
            raise IRSASyncError(f"Error durante el login: {e}")

    def _listar_tramites(self, endpoint_listar, shopping_id):
        """Lista los IDs de trámites (FAOs o FAPs) activos"""
        url = f"{self.BASE_URL}{endpoint_listar}"
        
        # Pedimos desde hace 2 aos (730 das) para traer todo lo activo
        hoy = datetime.now()
        desde = (hoy - timedelta(days=730)).strftime('%Y-%m-%d')
        hasta = (hoy + timedelta(days=180)).strftime('%Y-%m-%d')

        payload = {
            "start": 0,
            "limit": 2000,
            "shopping": [shopping_id],
            "desde": desde,
            "hasta": hasta
        }
        
        try:
            res = self.session.post(url, json=payload, verify=False, timeout=45)
            res.raise_for_status()
            data = res.json()
            
            # La API nueva devuelve los resultados en la clave 'results'
            items = []
            if isinstance(data, dict):
                items = data.get('results', data.get('data', []))
            elif isinstance(data, list):
                items = data
            
            ids = [item['id'] for item in items if 'id' in item]
            return ids
        except Exception as e:
            logger.error(f"Error listando {endpoint_listar}: {e}")
            raise RuntimeError(f"Fallo al listar tramites en IRSA: {e}")

    def _obtener_detalle(self, endpoint_obtener, item_id):
        """Obtiene el detalle completo de un trámite"""
        url = f"{self.BASE_URL}{endpoint_obtener}"
        try:
            res = self.session.post(url, json={"id": str(item_id)}, verify=False, timeout=10)
            if res.ok:
                return res.json()
        except Exception as e:
            logger.warning(f"Error obteniendo detalle de {item_id}: {e}")
        return None

    def sincronizar_todo(self):
        """Flujo principal: Obtiene credenciales, sincroniza y actualiza Excels locales"""
        creds = get_irsa_credentials()
        if not creds or not creds.get('username') or not creds.get('password'):
            raise IRSAAuthError("No hay credenciales configuradas en el sistema")
            
        logger.info(f"Iniciando sincronización IRSA para shopping {creds['shopping_id']}...")
        self.login(creds['username'], creds['password'])
        
        shopping_id = str(creds.get('shopping_id', '61709'))
        
        # 1. Obtener FAOs
        fao_ids = self._listar_tramites("/api/faos/listar", shopping_id)
        faos_data = []
        for fid in fao_ids:
            detalle = self._obtener_detalle("/api/faos/fao", fid)
            if detalle:
                faos_data.append(detalle)
            time.sleep(0.1) # Pausa amigable
            
        # 2. Obtener FAPs
        fap_ids = self._listar_tramites("/api/faps/listar", shopping_id)
        faps_data = []
        for fid in fap_ids:
            detalle = self._obtener_detalle("/api/faps/fap", fid)
            if detalle:
                faps_data.append(detalle)
            time.sleep(0.1)

        hoy_str = datetime.now().strftime('%Y-%m-%d')
        def is_not_expired(item):
            fecha_fin = item.get('fechaFin')
            if not fecha_fin: return True
            return fecha_fin.split('T')[0] >= hoy_str
            
        faos_data = [f for f in faos_data if is_not_expired(f)]
        faps_data = [f for f in faps_data if is_not_expired(f)]

        # 3. Guardar metadatos crudos para el dashboard (incluye todos los estados)
        with open(os.path.join(BASE_DIR, "MetadatosFAOs.json"), "w", encoding="utf-8") as f:
            json.dump(faos_data, f, ensure_ascii=False, indent=2)
        
        with open(os.path.join(BASE_DIR, "MetadatosFAPs.json"), "w", encoding="utf-8") as f:
            json.dump(faps_data, f, ensure_ascii=False, indent=2)

        # 4. Filtrar solo los Aprobados (6) y Aprobados Parciales (4) para el Excel
        estados_validos = ["4", "6"]
        faos_validos = [f for f in faos_data if str(f.get("estado")) in estados_validos]
        faps_validos = [f for f in faps_data if str(f.get("estado")) in estados_validos]

        # 5. Procesar y guardar en Excel solo los válidos
        self._guardar_excel_faos(faos_validos)
        self._guardar_excel_faps(faps_validos)
        
        logger.info("Sincronización IRSA finalizada exitosamente")
        return {"faos_procesados": len(faos_validos), "faps_procesados": len(faps_validos)}

    def _guardar_excel_faos(self, faos_data):
        """Convierte los datos JSON de FAOs al formato Excel local"""
        # Columnas esperadas en ListadoFAOs.xlsx:
        # ['FAO', 'Marca', 'Local', 'Fecha Inicio', 'Fecha Fin', 'Shopping', 'Tipo', 'Numero', 'Nombre', 'Apellido', 'Horario', 'Lugar', 'Tarea/s', 'Tipo.1', 'Locatario/Proveedor', 'Area Solicitante', 'Aprobacion automatica']
        filas = []
        
        for fao in faos_data:
            fao_id = fao.get('id', '')
            marca = fao.get('marca', '')
            local = fao.get('local', '')
            # Asumimos que las fechas vienen en formato ISO
            fecha_inicio = fao.get('fechaDesde', '').split('T')[0] if fao.get('fechaDesde') else ''
            fecha_fin = fao.get('fechaHasta', '').split('T')[0] if fao.get('fechaHasta') else ''
            
            personal = fao.get('personal', [])
            for p in personal:
                if not p.get('activo', True):
                    continue
                    
                filas.append({
                    'FAO': fao_id,
                    'Marca': marca,
                    'Local': local,
                    'Fecha Inicio': fecha_inicio,
                    'Fecha Fin': fecha_fin,
                    'Shopping': fao.get('shoppingNombre', 'AL OESTE SHOPPING'),
                    'Tipo': fao.get('tipo', 'MANTENIMIENTO'),
                    'Numero': p.get('nroDocumento', ''),
                    'Nombre': p.get('nombre', ''),
                    'Apellido': p.get('apellido', ''),
                    'Horario': fao.get('horario', ''),
                    'Lugar': fao.get('lugar', ''),
                    'Tarea/s': fao.get('tareas', ''),
                    'Tipo.1': p.get('tipoDocumento', 'DNI'),
                    'Locatario/Proveedor': fao.get('proveedor', ''),
                    'Area Solicitante': fao.get('area', ''),
                    'Aprobacion automatica': 'SI'
                })
                
        df = pd.DataFrame(filas)
        # Escribimos el header extra requerido por el formato actual ("Listado de FAOs" en fila 1)
        # Creamos un MultiIndex o lo guardamos simple. Por ahora simple + skiprows=1 match.
        
        # Para replicar el formato exacto de ListadoFAOs.xlsx donde la linea 1 es título:
        writer = pd.ExcelWriter(EXCEL_FAO, engine='openpyxl')
        df.to_excel(writer, index=False, startrow=1)
        
        # Escribimos el título en A1
        worksheet = writer.sheets['Sheet1']
        worksheet.cell(row=1, column=1, value="Listado de FAOs")
        writer.close()

    def _guardar_excel_faps(self, faps_data):
        """Convierte los datos JSON de FAPs al formato Excel local"""
        # Columnas esperadas en ListadoFAPs.xlsx:
        # ['FAP', 'Marca', 'Local', 'Fecha Inicio', 'Fecha Fin', 'Shopping', 'Tipo', 'Numero', 'Nombre', 'Apellido']
        filas = []
        
        for fap in faps_data:
            fap_id = fap.get('id', '')
            marca = fap.get('marca', '')
            local = fap.get('local', '')
            fecha_inicio = fap.get('fechaDesde', '').split('T')[0] if fap.get('fechaDesde') else ''
            fecha_fin = fap.get('fechaHasta', '').split('T')[0] if fap.get('fechaHasta') else ''
            
            personal = fap.get('personal', [])
            for p in personal:
                if not p.get('activo', True):
                    continue
                    
                filas.append({
                    'FAP': fap_id,
                    'Marca': marca,
                    'Local': local,
                    'Fecha Inicio': fecha_inicio,
                    'Fecha Fin': fecha_fin,
                    'Shopping': fap.get('shoppingNombre', 'AL OESTE SHOPPING'),
                    'Tipo': 'INGRESO',
                    'Numero': p.get('nroDocumento', ''),
                    'Nombre': p.get('nombre', ''),
                    'Apellido': p.get('apellido', '')
                })
                
        df = pd.DataFrame(filas)
        writer = pd.ExcelWriter(EXCEL_FAP, engine='openpyxl')
        df.to_excel(writer, index=False, startrow=1)
        
        worksheet = writer.sheets['Sheet1']
        worksheet.cell(row=1, column=1, value="Listado de FAPs")
        writer.close()
