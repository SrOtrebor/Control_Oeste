"""
Tests de Integración para Control de Acceso AOS

Tests básicos para endpoints principales de la aplicación Flask.
Ejecutar con: pytest test_integration.py -v
"""

import pytest
import sys
import os

# Agregar el directorio padre al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app


@pytest.fixture
def client():
    """Fixture que proporciona un cliente de prueba de Flask"""
    app.config['TESTING'] = True
    app.config['SECRET_KEY'] = 'test_secret_key'
    
    with app.test_client() as client:
        yield client


class TestEndpointsPublicos:
    """Tests para endpoints públicos (sin autenticación)"""
    
    def test_home_redirect(self, client):
        """Test que la ruta raíz redirige correctamente"""
        response = client.get('/')
        
        # Debería redirigir o retornar 200
        assert response.status_code in [200, 302, 308]
    
    def test_login_page_exists(self, client):
        """Test que la página de login existe"""
        response = client.get('/login')
        
        assert response.status_code == 200


class TestEndpointsProtegidos:
    """Tests para endpoints que requieren autenticación"""
    
    def test_admin_sin_autenticacion(self, client):
        """Test acceso a admin sin autenticación"""
        response = client.get('/admin')
        
        # Debería redirigir al login
        assert response.status_code in [302, 308, 401, 403]
    
    def test_upload_excel_sin_autenticacion(self, client):
        """Test subida de archivos sin autenticación"""
        response = client.post('/upload_excel')
        
        # Debería denegar acceso
        assert response.status_code in [302, 308, 401, 403]


class TestAPIEndpoints:
    """Tests para endpoints de API"""
    
    def test_verificar_dni_sin_datos(self, client):
        """Test verificar DNI sin enviar datos"""
        response = client.post('/verificar_dni', json={})
        
        # Debería retornar error o respuesta válida
        assert response.status_code in [200, 400]
    
    def test_verificar_dni_con_dni_invalido(self, client):
        """Test verificar DNI con formato inválido"""
        response = client.post('/verificar_dni', json={
            'scanner_data': 'INVALIDO',
            'mode': 'entrada'
        })
        
        assert response.status_code == 200
        data = response.get_json()
        
        # Debería denegar acceso
        if data:
            assert data.get('acceso') == 'DENEGADO'
    
    def test_get_ingresos_diarios(self, client):
        """Test obtener ingresos diarios"""
        response = client.get('/get_daily_records')
        
        assert response.status_code == 200
        data = response.get_json()
        
        # Debería retornar diccionario con clave records
        assert isinstance(data, dict)
        assert 'records' in data
    
    def test_get_stats(self, client):
        """Test obtener estadísticas"""
        response = client.get('/get_dynamic_stats')
        
        assert response.status_code == 200
        data = response.get_json()
        
        # Debería retornar un diccionario con stats
        assert isinstance(data, dict)
        assert 'total_adentro' in data


class TestLogin:
    """Tests para sistema de login"""
    
    def test_login_con_credenciales_invalidas(self, client):
        """Test login con credenciales incorrectas"""
        response = client.post('/perform_login', json={
            'username': 'usuario_invalido',
            'password': 'password_invalido'
        })
        
        # Debería rechazar el login
        assert response.status_code == 200
        data = response.get_json()
        assert data.get('success') is False
    
    def test_logout(self, client):
        """Test logout"""
        response = client.get('/logout', follow_redirects=False)
        
        # Debería redirigir
        assert response.status_code in [302, 308]


class TestReportes:
    """Tests para generación de reportes"""
    
    def test_descargar_reporte_sin_autenticacion(self, client):
        """Test descargar reporte sin autenticación"""
        response = client.get('/descargar_reporte_diario')
        
        # Debería denegar acceso
        assert response.status_code in [302, 308, 401, 403]
    
    def test_descargar_fichajes_sin_autenticacion(self, client):
        """Test descargar fichajes sin autenticación"""
        response = client.get('/descargar_reporte_fichajes')
        
        # Debería denegar acceso
        assert response.status_code in [302, 308, 401, 403]


class TestNominas:
    """Tests para gestión de nóminas"""
    
    def test_get_nominas_sin_autenticacion(self, client):
        """Test obtener nóminas sin autenticación"""
        response = client.get('/get_nominas_guardadas')
        
        # Debería denegar acceso
        assert response.status_code in [302, 308, 401, 403]
    
    def test_preview_nomina_sin_datos(self, client):
        """Test preview de nómina sin datos"""
        response = client.post('/parse_nomina', json={})
        
        # Debería retornar error o denegar acceso
        assert response.status_code in [200, 400, 401, 403]


if __name__ == '__main__':
    # Ejecutar tests
    pytest.main([__file__, '-v', '--tb=short'])
