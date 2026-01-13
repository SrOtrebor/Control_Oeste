# Control de Acceso AOS - Tests

## Ejecutar Tests

### Tests Unitarios
```bash
pytest test_access_manager.py -v
```

### Tests de Integración
```bash
pytest test_integration.py -v
```

### Todos los Tests
```bash
pytest -v
```

### Tests con Cobertura
```bash
pytest --cov=. --cov-report=html
```

## Tests Implementados

### test_access_manager.py (Tests Unitarios)

**TestParsearCodigoBarra** - 8 tests
- ✅ Formato DNI nuevo con comillas
- ✅ Formato con @ y guiones bajos
- ✅ DNI solo números (7-8 dígitos)
- ✅ DNI de 7 dígitos
- ✅ Extracción de DNI con caracteres extra
- ✅ Scanner data vacío
- ✅ Scanner data sin DNI

**TestValidacionDNI** - 7 tests
- ✅ DNI válido 8 dígitos
- ✅ DNI válido 7 dígitos
- ✅ DNI inválido (muy corto)
- ✅ DNI inválido (muy largo)
- ✅ DNI inválido con letras
- ✅ DNI vacío
- ✅ DNI None

**TestBackups** - 2 tests
- ✅ Crear backup de archivo inexistente
- ✅ Limpiar backups antiguos

**TestValidacionArchivos** - 2 tests
- ✅ Validar archivo FAP inexistente
- ✅ Validar archivo FAO inexistente

**Total: 19 tests unitarios**

### test_integration.py (Tests de Integración)

**TestEndpointsPublicos** - 2 tests
- ✅ Home redirect
- ✅ Login page exists

**TestEndpointsProtegidos** - 2 tests
- ✅ Admin sin autenticación
- ✅ Upload Excel sin autenticación

**TestAPIEndpoints** - 4 tests
- ✅ Verificar DNI sin datos
- ✅ Verificar DNI con DNI inválido
- ✅ Get ingresos diarios
- ✅ Get stats

**TestLogin** - 2 tests
- ✅ Login con credenciales inválidas
- ✅ Logout

**TestReportes** - 2 tests
- ✅ Descargar reporte sin autenticación
- ✅ Descargar fichajes sin autenticación

**TestNominas** - 2 tests
- ✅ Get nóminas sin autenticación
- ✅ Preview nómina sin datos

**Total: 14 tests de integración**

## Cobertura de Tests

- **Parseo de DNI**: 4 formatos diferentes
- **Validación**: DNI, archivos FAP/FAO
- **Backups**: Creación y limpieza
- **Endpoints**: Públicos, protegidos, API
- **Autenticación**: Login, logout, protección

## Notas

- Los tests están diseñados para ejecutarse sin necesidad de datos reales
- Los tests de integración usan el modo TESTING de Flask
- Algunos tests verifican comportamiento de error (archivos inexistentes, etc.)
