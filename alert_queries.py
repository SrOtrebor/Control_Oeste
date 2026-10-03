from database import get_db_connection

def es_persona_de_interes(dni):
    """Verifica si un DNI está en la lista negra."""
    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT motivo FROM lista_negra WHERE dni = ?", (dni,))
            row = cursor.fetchone()
            if row:
                return True, row['motivo']
            return False, ""
    except Exception:
        return False, ""
