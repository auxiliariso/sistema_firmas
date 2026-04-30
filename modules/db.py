"""
db.py — Módulo de base de datos
Gestiona la conexión SQLite, creación de tablas y operaciones CRUD.
"""

import sqlite3
import hashlib
import os
from datetime import datetime

# Ruta de la base de datos (junto al ejecutable/script)
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'firma_digital.db')

# ─────────────────────────────────────────────
# CONEXIÓN
# ─────────────────────────────────────────────

def get_connection():
    """Retorna una conexión activa a la base de datos SQLite."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row  # Acceso por nombre de columna
    return conn


# ─────────────────────────────────────────────
# INICIALIZACIÓN
# ─────────────────────────────────────────────

def inicializar_db():
    """
    Crea las tablas necesarias si no existen.
    Se ejecuta al arrancar la aplicación.
    """
    with get_connection() as conn:
        cursor = conn.cursor()

        # Tabla de usuarios
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                usuario          TEXT PRIMARY KEY,
                password_hash    TEXT NOT NULL,
                salt             TEXT NOT NULL,
                nombre           TEXT NOT NULL,
                bloqueado        BOOLEAN DEFAULT 0,
                intentos_fallidos INTEGER DEFAULT 0
            )
        """)

        # Tabla de log de firmas
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS log_firmas (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario          TEXT NOT NULL,
                fecha_hora       TEXT NOT NULL,
                id_firma         TEXT NOT NULL UNIQUE,
                documento        TEXT NOT NULL
            )
        """)

        conn.commit()


# ─────────────────────────────────────────────
# GESTIÓN DE USUARIOS
# ─────────────────────────────────────────────

def crear_usuario(usuario: str, password: str, nombre: str) -> bool:
    """
    Crea un nuevo usuario con contraseña hasheada + salt.
    Retorna True si se creó, False si ya existe.
    """
    # Generar salt aleatorio (16 bytes → 32 hex chars)
    salt = os.urandom(16).hex()
    password_hash = _hash_password(password, salt)

    try:
        with get_connection() as conn:
            conn.execute(
                "INSERT INTO usuarios (usuario, password_hash, salt, nombre) VALUES (?, ?, ?, ?)",
                (usuario, password_hash, salt, nombre)
            )
            conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False  # El usuario ya existe


def _hash_password(password: str, salt: str) -> str:
    """
    Genera un hash SHA-256 de la contraseña combinada con el salt.
    Se usa PBKDF2 con 260,000 iteraciones para mayor seguridad.
    """
    dk = hashlib.pbkdf2_hmac(
        'sha256',
        password.encode('utf-8'),
        salt.encode('utf-8'),
        iterations=260_000
    )
    return dk.hex()


def verificar_credenciales(usuario: str, password: str) -> dict | None:
    """
    Verifica usuario y contraseña.
    Retorna el registro del usuario si son correctas, None si no.
    Gestiona automáticamente el contador de intentos fallidos.
    """
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM usuarios WHERE usuario = ?", (usuario,)
        ).fetchone()

    if not row:
        return None  # Usuario no existe

    # Verificar si está bloqueado
    if row['bloqueado']:
        return {'bloqueado': True}

    # Verificar contraseña
    hash_calculado = _hash_password(password, row['salt'])
    if hash_calculado == row['password_hash']:
        # Éxito: resetear intentos fallidos
        _resetear_intentos(usuario)
        return dict(row)
    else:
        # Fallo: incrementar intentos
        _registrar_intento_fallido(usuario)
        return None


def _registrar_intento_fallido(usuario: str):
    """Incrementa el contador de intentos fallidos y bloquea si llega a 3."""
    with get_connection() as conn:
        conn.execute(
            """UPDATE usuarios
               SET intentos_fallidos = intentos_fallidos + 1,
                   bloqueado = CASE WHEN intentos_fallidos + 1 >= 3 THEN 1 ELSE 0 END
               WHERE usuario = ?""",
            (usuario,)
        )
        conn.commit()


def _resetear_intentos(usuario: str):
    """Resetea el contador de intentos fallidos tras login exitoso."""
    with get_connection() as conn:
        conn.execute(
            "UPDATE usuarios SET intentos_fallidos = 0, bloqueado = 0 WHERE usuario = ?",
            (usuario,)
        )
        conn.commit()


def obtener_intentos_restantes(usuario: str) -> int:
    """Retorna cuántos intentos le quedan al usuario antes de ser bloqueado."""
    with get_connection() as conn:
        row = conn.execute(
            "SELECT intentos_fallidos FROM usuarios WHERE usuario = ?", (usuario,)
        ).fetchone()
    if not row:
        return 3
    return max(0, 3 - row['intentos_fallidos'])


# ─────────────────────────────────────────────
# GESTIÓN DE LOG DE FIRMAS
# ─────────────────────────────────────────────

def guardar_firma(usuario: str, fecha_hora: str, id_firma: str, documento: str):
    """Registra una firma en el log."""
    with get_connection() as conn:
        conn.execute(
            "INSERT INTO log_firmas (usuario, fecha_hora, id_firma, documento) VALUES (?, ?, ?, ?)",
            (usuario, fecha_hora, id_firma, documento)
        )
        conn.commit()


def verificar_firma(id_firma: str) -> dict | None:
    """
    Busca un ID de firma en el log.
    Retorna el registro completo si existe, None si no.
    """
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM log_firmas WHERE id_firma = ?", (id_firma,)
        ).fetchone()
    return dict(row) if row else None


def obtener_firmas_usuario(usuario: str) -> list:
    """Retorna todas las firmas registradas por un usuario."""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM log_firmas WHERE usuario = ? ORDER BY fecha_hora DESC",
            (usuario,)
        ).fetchall()
    return [dict(r) for r in rows]
