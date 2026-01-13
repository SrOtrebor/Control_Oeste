"""
Tests Unitarios para Control de Acceso AOS

Tests básicos para las funciones más críticas del sistema.
Ejecutar con: pytest test_access_manager.py -v
"""

import pytest
import sys
import os

# Agregar el directorio padre al path para importar módulos
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from access_manager import parsear_codigo_barra


class TestParsearCodigoBarra:
    """Tests para la función parsear_codigo_barra que maneja múltiples formatos de DNI"""
    
    def test_formato_dni_nuevo_con_comillas(self):
        """Test formato DNI nuevo con comillas (8+ campos)"""
        scanner_data = '"PEREZ"JUAN"M"12345678"A"19900101"20300101"'
        resultado = parsear_codigo_barra(scanner_data)
        
        assert resultado is not None
        assert resultado['dni'] == '12345678'
        assert resultado['nombre'] == 'JUAN'
        assert resultado['apellido'] == 'PEREZ'
        assert resultado['nombre_completo'] == 'JUAN PEREZ'
        assert resultado['sexo'] == 'M'
    
    def test_formato_con_arroba_y_guiones(self):
        """Test formato con @ y guiones bajos"""
        scanner_data = '@PEREZ_JUAN_M_12345678_'
        resultado = parsear_codigo_barra(scanner_data)
        
        assert resultado is not None
        assert resultado['dni'] == '12345678'
        assert resultado['nombre'] == 'JUAN'
        assert resultado['apellido'] == 'PEREZ'
        assert resultado['nombre_completo'] == 'JUAN PEREZ'
    
    def test_formato_dni_solo_numeros(self):
        """Test búsqueda de DNI de 7-8 dígitos en texto"""
        scanner_data = 'Algún texto con DNI 12345678 en el medio'
        resultado = parsear_codigo_barra(scanner_data)
        
        assert resultado is not None
        assert resultado['dni'] == '12345678'
    
    def test_formato_dni_7_digitos(self):
        """Test DNI de 7 dígitos"""
        scanner_data = '1234567'
        resultado = parsear_codigo_barra(scanner_data)
        
        assert resultado is not None
        assert resultado['dni'] == '1234567'
    
    def test_formato_dni_con_caracteres_extra(self):
        """Test extracción de dígitos de texto con caracteres extra"""
        scanner_data = 'ABC-12345678-XYZ'
        resultado = parsear_codigo_barra(scanner_data)
        
        assert resultado is not None
        assert resultado['dni'] == '12345678'
    
    def test_scanner_data_vacio(self):
        """Test con datos vacíos"""
        resultado = parsear_codigo_barra('')
        
        # Debería retornar un dict con dni vacío o None
        assert resultado is None or resultado.get('dni') == ''
    
    def test_scanner_data_sin_dni(self):
        """Test con datos que no contienen DNI"""
        scanner_data = 'ABCDEFGH'
        resultado = parsear_codigo_barra(scanner_data)
        
        # Puede retornar None o un dict con dni inválido
        if resultado:
            assert len(resultado.get('dni', '')) < 7 or not resultado.get('dni', '').isdigit()


class TestValidacionDNI:
    """Tests para validación de formato de DNI"""
    
    def test_dni_valido_8_digitos(self):
        """Test DNI válido de 8 dígitos"""
        from utils import validar_dni
        
        assert validar_dni('12345678') == True
    
    def test_dni_valido_7_digitos(self):
        """Test DNI válido de 7 dígitos"""
        from utils import validar_dni
        
        assert validar_dni('1234567') == True
    
    def test_dni_invalido_muy_corto(self):
        """Test DNI inválido (muy corto)"""
        from utils import validar_dni
        
        assert validar_dni('12345') == False
    
    def test_dni_invalido_muy_largo(self):
        """Test DNI inválido (muy largo)"""
        from utils import validar_dni
        
        assert validar_dni('123456789') == False
    
    def test_dni_invalido_con_letras(self):
        """Test DNI inválido con letras"""
        from utils import validar_dni
        
        assert validar_dni('1234567A') == False
    
    def test_dni_vacio(self):
        """Test DNI vacío"""
        from utils import validar_dni
        
        assert validar_dni('') == False
    
    def test_dni_none(self):
        """Test DNI None"""
        from utils import validar_dni
        
        assert validar_dni(None) == False


class TestBackups:
    """Tests para sistema de backups"""
    
    def test_crear_backup_archivo_inexistente(self):
        """Test crear backup de archivo que no existe"""
        from utils import crear_backup
        
        exito, mensaje = crear_backup('archivo_que_no_existe.xlsx', 'manual')
        
        assert exito == False
        assert 'no existe' in mensaje.lower()
    
    def test_limpiar_backups_antiguos(self):
        """Test limpieza de backups antiguos"""
        from utils import limpiar_backups_antiguos
        
        # Debería ejecutarse sin errores
        eliminados = limpiar_backups_antiguos(dias=30, tipo='auto')
        
        assert isinstance(eliminados, int)
        assert eliminados >= 0


class TestValidacionArchivos:
    """Tests para validación de archivos FAP/FAO"""
    
    def test_validar_archivo_fap_inexistente(self):
        """Test validar archivo FAP que no existe"""
        from utils import validar_archivo_fap
        
        valido, mensaje = validar_archivo_fap('archivo_inexistente.xlsx')
        
        assert valido == False
        assert 'error' in mensaje.lower() or 'no' in mensaje.lower()
    
    def test_validar_archivo_fao_inexistente(self):
        """Test validar archivo FAO que no existe"""
        from utils import validar_archivo_fao
        
        valido, mensaje = validar_archivo_fao('archivo_inexistente.xlsx')
        
        assert valido == False
        assert 'error' in mensaje.lower() or 'no' in mensaje.lower()


# Fixtures para tests que requieren datos de prueba
@pytest.fixture
def dni_valido():
    """Fixture con DNI válido de prueba"""
    return '12345678'


@pytest.fixture
def scanner_data_formato1():
    """Fixture con datos de scanner formato 1"""
    return '"PEREZ"JUAN"M"12345678"A"19900101"20300101"'


@pytest.fixture
def scanner_data_formato2():
    """Fixture con datos de scanner formato 2"""
    return '@PEREZ_JUAN_M_12345678_'


if __name__ == '__main__':
    # Ejecutar tests
    pytest.main([__file__, '-v', '--tb=short'])
