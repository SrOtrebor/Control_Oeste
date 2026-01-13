import re
import os

def replace_prints_in_file(filepath):
    """Reemplaza todos los prints por logging en un archivo"""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    
    # Reemplazos específicos por tipo de mensaje
    replacements = [
        # INFO
        (r'print\(f"INFO: ([^"]+)"\)', r'logger.info(f"\1")'),
        (r'print\("INFO: ([^"]+)"\)', r'logger.info("\1")'),
        
        # ADVERTENCIA
        (r'print\(f"ADVERTENCIA: ([^"]+)"\)', r'logger.warning(f"\1")'),
        (r'print\("ADVERTENCIA: ([^"]+)"\)', r'logger.warning("\1")'),
        
        # ERROR CRÍTICO
        (r'print\(f"Error Crítico ([^"]+)"\)', r'logger.error(f"Error crítico \1", exc_info=True)'),
        (r'print\(f"Error crítico ([^"]+)"\)', r'logger.error(f"Error crítico \1", exc_info=True)'),
        
        # ERROR general
        (r'print\(f"Error ([^"]+)"\)', r'logger.error(f"Error \1")'),
        (r'print\("Error ([^"]+)"\)', r'logger.error("Error \1")'),
        
        # DEBUG (flechas y guiones)
        (r'print\(f"-> ([^"]+)"\)', r'logger.debug(f"\1")'),
        (r'print\("-> ([^"]+)"\)', r'logger.debug("\1")'),
        (r'print\(f"--- ([^"]+)"\)', r'logger.debug(f"\1")'),
        (r'print\("--- ([^"]+)"\)', r'logger.debug("\1")'),
        
        # DEBUG (líneas de procesamiento)
        (r'print\(f"\\\\n\[Línea ([^"]+)"\)', r'logger.debug(f"Línea \1")'),
        (r'print\(f"   - ([^"]+)"\)', r'logger.debug(f"\1")'),
        
        # DEBUG general
        (r'print\(f"Texto ([^"]+)"\)', r'logger.debug(f"Texto \1")'),
        (r'print\(f"Procesando ([^"]+)"\)', r'logger.debug(f"Procesando \1")'),
        (r'print\(f"Coincide ([^"]+)"\)', r'logger.debug(f"Coincide \1")'),
        (r'print\(f"Datos ([^"]+)"\)', r'logger.debug(f"Datos \1")'),
        (r'print\(f"CUIL ([^"]+)"\)', r'logger.debug(f"CUIL \1")'),
        (r'print\(f"DNI ([^"]+)"\)', r'logger.debug(f"DNI \1")'),
        (r'print\(f"Nombre ([^"]+)"\)', r'logger.debug(f"Nombre \1")'),
        (r'print\(f"Apellido ([^"]+)"\)', r'logger.debug(f"Apellido \1")'),
        (r'print\(f"ÉXITO: ([^"]+)"\)', r'logger.debug(f"ÉXITO: \1")'),
    ]
    
    for pattern, replacement in replacements:
        content = re.sub(pattern, replacement, content)
    
    if content != original_content:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    return False

# Procesar archivos
files = ['data_manager.py', 'app.py']
for file in files:
    if os.path.exists(file):
        if replace_prints_in_file(file):
            print(f'✅ {file} actualizado')
        else:
            print(f'ℹ️  {file} sin cambios')
