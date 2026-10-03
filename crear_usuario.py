import firebase_admin
from firebase_admin import credentials, auth, firestore

# Inicializar Firebase
try:
    cred = credentials.Certificate('serviceAccountKey.json')
    firebase_admin.initialize_app(cred)
    db = firestore.client()
except Exception as e:
    print(f"Error al inicializar Firebase. Asegúrate de tener serviceAccountKey.json en la carpeta. Error: {e}")
    exit(1)

def crear_usuario():
    print("=== CREADOR DE USUARIOS PARA CONTROL DE ACCESO ===")
    print("Este script creará el usuario en Authentication y le asignará su Rol en la Base de Datos.\n")
    
    email = input("✉️  Ingrese el CORREO del nuevo usuario (ej: garita1@oeste.com): ").strip()
    password = input("🔑 Ingrese la CONTRASEÑA (mínimo 6 caracteres): ").strip()
    
    print("\nSeleccione el ROL del usuario:")
    print("1) super_admin (Acceso total al Dashboard Maestro)")
    print("2) supervisor (Acceso total al Dashboard del Oeste)")
    print("3) acceso_principal (Garita principal con escáner y fichajes)")
    print("4) acceso_satelite (Garita satélite solo escaneo)")
    
    rol_opcion = input("👉 Ingrese el número del rol (1-4): ").strip()
    
    roles_map = {
        '1': 'super_admin',
        '2': 'supervisor',
        '3': 'acceso_principal',
        '4': 'acceso_satelite'
    }
    
    rol = roles_map.get(rol_opcion)
    
    if not rol:
        print("❌ Opción inválida. Cancelando.")
        return

    try:
        # 1. Crear en Auth
        user = auth.create_user(
            email=email,
            password=password
        )
        print(f"\n✅ Usuario '{email}' creado exitosamente en Authentication (UID: {user.uid}).")
        
        # 2. Guardar Rol en Firestore
        db.collection('usuarios').document(user.uid).set({
            'email': email,
            'rol': rol,
            'activo': True,
            'centro_id': 'al_oeste'
        })
        print(f"✅ Rol '{rol}' asignado correctamente en la Base de Datos.")
        
    except Exception as e:
        print(f"\n❌ Hubo un error al crear el usuario: {e}")

if __name__ == "__main__":
    crear_usuario()
    input("\nPresione ENTER para salir...")
