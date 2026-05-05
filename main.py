"""
main.py — Punto de entrada principal
Inicializa la base de datos, crea usuarios de prueba y lanza la aplicación.

Uso:
    python main.py

Para crear ejecutable .exe:
    pip install pyinstaller
    pyinstaller --onefile --windowed --name FirmaDigital main.py
"""

import sys
import os

# Asegura que el directorio raíz esté en el path (importante para el .exe)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules.db    import inicializar_db, crear_usuario
from modules.login import VentanaLogin
from modules.app   import AppFirmaDigital

# ─────────────────────────────────────────────
# USUARIOS DE DEMOSTRACIÓN
# ─────────────────────────────────────────────

USUARIOS_DEMO = [
    # (usuario,    contraseña,     nombre completo)
    ("jfUrrutia",    "C_ISO2026",   "Juan Francisco Ramirez Urrutia"),
    ("Rgarnica",   "A_ISO2026",   "Roberto Enriquez Garnica"),
    ("admin",     "Admin@999",    "Administrador"),
]


def crear_usuarios_demo():
    """
    Crea usuarios de demostración si no existen.
    Las contraseñas se hashean automáticamente; nunca se guardan en texto plano.
    """
    for usuario, password, nombre in USUARIOS_DEMO:
        creado = crear_usuario(usuario, password, nombre)
        if creado:
            print(f"  ✓ Usuario creado: {usuario}")
        else:
            print(f"  · Ya existe:      {usuario}")


# ─────────────────────────────────────────────
# PUNTO DE ENTRADA
# ─────────────────────────────────────────────

def main():
    print("=" * 45)
    print("  🔏  Sistema de Firma Digital v1.0")
    print("=" * 45)

    # 1. Inicializar base de datos y tablas
    print("\n[DB] Inicializando base de datos...")
    inicializar_db()
    print("[DB] ✓ Base de datos lista.")

    # 2. Crear usuarios demo (solo si no existen)
    print("\n[DB] Verificando usuarios de demostración:")
    crear_usuarios_demo()

    print("\n[APP] Lanzando interfaz gráfica...\n")

    # 3. Mostrar ventana de login
    login = VentanaLogin()
    usuario_data = login.mostrar()

    if not usuario_data:
        # El usuario cerró la ventana sin autenticarse
        print("[APP] Sesión cancelada por el usuario.")
        sys.exit(0)

    print(f"[APP] ✓ Sesión iniciada: {usuario_data['nombre']}")

    # 4. Lanzar ventana principal
    app = AppFirmaDigital(usuario_data)
    app.iniciar()

    print("[APP] Sesión cerrada.")


if __name__ == "__main__":
    main()
