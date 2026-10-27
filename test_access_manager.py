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


class TestResolucionPermisosMultiples:
    """Tests para validación de fechas de vencimiento y resolución de permisos múltiples"""

    def test_es_fecha_vencida(self):
        from access_manager import es_fecha_vencida
        from datetime import datetime, timedelta
        
        ayer = (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')
        manana = (datetime.now() + timedelta(days=1)).strftime('%d/%m/%Y')
        
        assert es_fecha_vencida(ayer) is True
        assert es_fecha_vencida(manana) is False
        assert es_fecha_vencida('Indefinido') is False
        assert es_fecha_vencida('') is False
        assert es_fecha_vencida(None) is False

    def test_verificar_dni_prioriza_vigente_sobre_vencido(self, monkeypatch):
        import access_manager
        from datetime import datetime, timedelta
        
        ayer_str = (datetime.now() - timedelta(days=5)).strftime('%d/%m/%Y')
        futuro_str = (datetime.now() + timedelta(days=30)).strftime('%d/%m/%Y')
        
        # Simular que el directorio IRSA tiene dos permisos para el mismo DNI: uno vencido y uno vigente
        dni_test = "33444555"
        mock_irsa = [
            {
                'dni': dni_test,
                'nombre': 'GOMEZ JUAN',
                'empresa': 'EMPRESA A',
                'vence': ayer_str,
                'tipo': 'FAO',
                'numero': '101'
            },
            {
                'dni': dni_test,
                'nombre': 'GOMEZ JUAN',
                'empresa': 'EMPRESA B',
                'vence': futuro_str,
                'tipo': 'FAO',
                'numero': '202'
            }
        ]
        
        monkeypatch.setattr(access_manager.data_manager, 'obtener_directorio_irsa', lambda: mock_irsa)
        monkeypatch.setattr(access_manager, 'verificar_dni_sqlite', lambda dni: None)
        monkeypatch.setattr(access_manager, 'es_persona_de_interes', lambda dni: (False, None))
        monkeypatch.setattr(access_manager, 'log_access_event', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'registrar_evento', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'loggear_acceso_sqlite', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'guardar_estado_adentro_sqlite', lambda *args, **kwargs: None)
        
        res = access_manager.verificar_dni(dni_test, mode='entrada')
        
        # Debe ser PERMITIDO y elegir el permiso vigente (202), no el vencido (101)
        assert res['acceso'] == 'PERMITIDO'
        assert res['num_permiso'] == '202'
        assert res['vence'] == futuro_str
        assert res['alerta_vencimiento'] is False

    def test_prioriza_vigente_sqlite_sobre_vencido_sqlite(self, monkeypatch):
        import access_manager
        from datetime import datetime, timedelta
        import database
        
        ayer_str = (datetime.now() - timedelta(days=5)).strftime('%d/%m/%Y')
        futuro_str = (datetime.now() + timedelta(days=20)).strftime('%d/%m/%Y')
        dni_test = "77889901"

        with database.get_db_connection() as conn:
            conn.execute("DELETE FROM autorizaciones WHERE dni LIKE ?", (dni_test + "%",))
            conn.execute('''
                INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, num_permiso, local, tarea, fecha_fin, activo)
                VALUES (?, 'JUAN', 'TEST', 'FAO', 'VENC-01', 'LOCAL A', 'TAREA A', ?, 1)
            ''', (dni_test, ayer_str))
            conn.execute('''
                INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, num_permiso, local, tarea, fecha_fin, activo)
                VALUES (?, 'JUAN', 'TEST', 'NOMINA', 'VIG-02', 'LOCAL B', 'TAREA B', ?, 1)
            ''', (dni_test, futuro_str))

        monkeypatch.setattr(access_manager.data_manager, 'obtener_directorio_irsa', lambda: [])
        monkeypatch.setattr(access_manager, 'es_persona_de_interes', lambda dni: (False, None))
        monkeypatch.setattr(access_manager, 'log_access_event', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'registrar_evento', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'loggear_acceso_sqlite', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'guardar_estado_adentro_sqlite', lambda *args, **kwargs: None)

        try:
            res = access_manager.verificar_dni(dni_test, mode='entrada')
            assert res['acceso'] == 'PERMITIDO'
            assert res['tipo_permiso'] == 'NOMINA'
            assert res['num_permiso'] == 'VIG-02'
            assert res['vence'] == futuro_str
        finally:
            with database.get_db_connection() as conn:
                conn.execute("DELETE FROM autorizaciones WHERE dni LIKE ?", (dni_test + "%",))

    def test_prioriza_vigente_sqlite_sobre_vencido_irsa(self, monkeypatch):
        import access_manager
        from datetime import datetime, timedelta
        import database
        
        ayer_str = (datetime.now() - timedelta(days=10)).strftime('%d/%m/%Y')
        futuro_str = (datetime.now() + timedelta(days=45)).strftime('%d/%m/%Y')
        dni_test = "77889902"

        mock_irsa = [{
            'dni': dni_test,
            'nombre': 'CARLOS IRSA',
            'empresa': 'CONTRATISTA X',
            'vence': ayer_str,
            'tipo': 'FAO',
            'numero': 'IRSA-VENC'
        }]

        with database.get_db_connection() as conn:
            conn.execute("DELETE FROM autorizaciones WHERE dni LIKE ?", (dni_test + "%",))
            conn.execute('''
                INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, num_permiso, local, tarea, fecha_fin, activo)
                VALUES (?, 'CARLOS', 'SQL', 'EXCEPCION', 'SQL-VIG', 'LOCAL C', 'REPARACION', ?, 1)
            ''', (dni_test, futuro_str))

        monkeypatch.setattr(access_manager.data_manager, 'obtener_directorio_irsa', lambda: mock_irsa)
        monkeypatch.setattr(access_manager, 'es_persona_de_interes', lambda dni: (False, None))
        monkeypatch.setattr(access_manager, 'log_access_event', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'registrar_evento', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'loggear_acceso_sqlite', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'guardar_estado_adentro_sqlite', lambda *args, **kwargs: None)

        try:
            res = access_manager.verificar_dni(dni_test, mode='entrada')
            assert res['acceso'] == 'PERMITIDO'
            assert res['tipo_permiso'] == 'EXCEPCION'
            assert res['num_permiso'] == 'SQL-VIG'
            assert res['vence'] == futuro_str
        finally:
            with database.get_db_connection() as conn:
                conn.execute("DELETE FROM autorizaciones WHERE dni LIKE ?", (dni_test + "%",))

    def test_prioriza_vigente_irsa_sobre_vencido_sqlite(self, monkeypatch):
        import access_manager
        from datetime import datetime, timedelta
        import database
        
        ayer_str = (datetime.now() - timedelta(days=15)).strftime('%d/%m/%Y')
        futuro_str = (datetime.now() + timedelta(days=15)).strftime('%d/%m/%Y')
        dni_test = "77889903"

        mock_irsa = [{
            'dni': dni_test,
            'nombre': 'ANA IRSA',
            'empresa': 'EMPRESA Y',
            'tarea': 'PINTURA',
            'vence': futuro_str,
            'tipo': 'FAO',
            'numero': 'FAO-VIGENTE'
        }]

        with database.get_db_connection() as conn:
            conn.execute("DELETE FROM autorizaciones WHERE dni LIKE ?", (dni_test + "%",))
            conn.execute('''
                INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, num_permiso, local, tarea, fecha_fin, activo)
                VALUES (?, 'ANA', 'SQL', 'FAO', 'FAO-VIEJO', 'LOCAL D', 'ANTIGUO', ?, 1)
            ''', (dni_test, ayer_str))

        monkeypatch.setattr(access_manager.data_manager, 'obtener_directorio_irsa', lambda: mock_irsa)
        monkeypatch.setattr(access_manager, 'es_persona_de_interes', lambda dni: (False, None))
        monkeypatch.setattr(access_manager, 'log_access_event', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'registrar_evento', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'loggear_acceso_sqlite', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'guardar_estado_adentro_sqlite', lambda *args, **kwargs: None)

        try:
            res = access_manager.verificar_dni(dni_test, mode='entrada')
            assert res['acceso'] == 'PERMITIDO'
            assert res['num_permiso'] == 'FAO-VIGENTE'
            assert res['vence'] == futuro_str
        finally:
            with database.get_db_connection() as conn:
                conn.execute("DELETE FROM autorizaciones WHERE dni LIKE ?", (dni_test + "%",))

    def test_todos_vencidos_rechaza_con_el_mas_reciente(self, monkeypatch):
        import access_manager
        from datetime import datetime, timedelta
        import database
        
        hace_un_mes = (datetime.now() - timedelta(days=30)).strftime('%d/%m/%Y')
        hace_tres_dias = (datetime.now() - timedelta(days=3)).strftime('%d/%m/%Y')
        dni_test = "77889904"

        with database.get_db_connection() as conn:
            conn.execute("DELETE FROM autorizaciones WHERE dni LIKE ?", (dni_test + "%",))
            conn.execute('''
                INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, num_permiso, local, tarea, fecha_fin, activo)
                VALUES (?, 'PEDRO', 'TEST', 'FAO', 'VIEJO-30', 'LOCAL 1', 'T1', ?, 1)
            ''', (dni_test, hace_un_mes))
            conn.execute('''
                INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, num_permiso, local, tarea, fecha_fin, activo)
                VALUES (?, 'PEDRO', 'TEST', 'FAO', 'VIEJO-03', 'LOCAL 2', 'T2', ?, 1)
            ''', (dni_test, hace_tres_dias))

        monkeypatch.setattr(access_manager.data_manager, 'obtener_directorio_irsa', lambda: [])
        monkeypatch.setattr(access_manager, 'es_persona_de_interes', lambda dni: (False, None))
        monkeypatch.setattr(access_manager, 'log_access_event', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'registrar_evento', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'loggear_acceso_sqlite', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'guardar_estado_adentro_sqlite', lambda *args, **kwargs: None)

        try:
            res = access_manager.verificar_dni(dni_test, mode='entrada')
            assert res['acceso'] == 'DENEGADO'
            assert 'VENCIDO' in res['mensaje']
            assert res['num_permiso'] == 'VIEJO-03'
            assert res['vence'] == hace_tres_dias
        finally:
            with database.get_db_connection() as conn:
                conn.execute("DELETE FROM autorizaciones WHERE dni LIKE ?", (dni_test + "%",))

    def test_multiples_vigentes_prioriza_vencimiento_mas_lejano(self, monkeypatch):
        import access_manager
        from datetime import datetime, timedelta
        import database
        
        vence_pronto = (datetime.now() + timedelta(days=1)).strftime('%d/%m/%Y')
        vence_lejano = (datetime.now() + timedelta(days=60)).strftime('%d/%m/%Y')
        dni_test = "77889905"

        with database.get_db_connection() as conn:
            conn.execute("DELETE FROM autorizaciones WHERE dni LIKE ?", (dni_test + "%",))
            conn.execute('''
                INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, num_permiso, local, tarea, fecha_fin, activo)
                VALUES (?, 'MARIA', 'TEST', 'FAO', 'PRONTO', 'LOCAL A', 'T1', ?, 1)
            ''', (dni_test, vence_pronto))
            conn.execute('''
                INSERT INTO autorizaciones (dni, nombre, apellido, tipo_permiso, num_permiso, local, tarea, fecha_fin, activo)
                VALUES (?, 'MARIA', 'TEST', 'NOMINA', 'LEJANO', 'LOCAL B', 'T2', ?, 1)
            ''', (dni_test, vence_lejano))

        monkeypatch.setattr(access_manager.data_manager, 'obtener_directorio_irsa', lambda: [])
        monkeypatch.setattr(access_manager, 'es_persona_de_interes', lambda dni: (False, None))
        monkeypatch.setattr(access_manager, 'log_access_event', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'registrar_evento', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'loggear_acceso_sqlite', lambda *args, **kwargs: None)
        monkeypatch.setattr(access_manager, 'guardar_estado_adentro_sqlite', lambda *args, **kwargs: None)

        try:
            res = access_manager.verificar_dni(dni_test, mode='entrada')
            assert res['acceso'] == 'PERMITIDO'
            assert res['num_permiso'] == 'LEJANO'
            assert res['vence'] == vence_lejano
            assert res['alerta_vencimiento'] is False
        finally:
            with database.get_db_connection() as conn:
                conn.execute("DELETE FROM autorizaciones WHERE dni LIKE ?", (dni_test + "%",))


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
