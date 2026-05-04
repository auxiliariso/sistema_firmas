"""
api.py — API REST de Firma Digital
Expone los endpoints HTTP para autenticación, firma y verificación.
No requiere interfaz gráfica. Ideal para integración con otros sistemas.

Uso:
    python api.py                        ← arranca en http://localhost:5000
    python api.py --port 8080            ← puerto personalizado
    python api.py --host 0.0.0.0         ← accesible en red local

Endpoints:
    POST /login                          ← autenticar usuario
    POST /firmar                         ← generar firma (requiere token)
    GET  /verificar/<id_firma>           ← verificar si una firma existe
    GET  /historial/<usuario>            ← firmas de un usuario (requiere token)
    POST /usuarios/crear                 ← crear usuario (requiere token admin)
    GET  /health                         ← estado del servidor
"""

import sys
import os
import argparse
import secrets
import hashlib
from datetime import datetime, timedelta
from functools import wraps

# ── Fix de path ──────────────────────────────────────────────────────────────
_ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if _ROOT_DIR not in sys.path:
    sys.path.insert(0, _ROOT_DIR)
os.chdir(_ROOT_DIR)
# ─────────────────────────────────────────────────────────────────────────────

try:
    from flask import Flask, request, jsonify
except ImportError:
    print("Flask no está instalado. Ejecuta: pip install flask")
    sys.exit(1)

from modules.db    import inicializar_db, verificar_credenciales, crear_usuario, obtener_firmas_usuario
from modules.firma import crear_firma, validar_firma

app = Flask(__name__)

# ─────────────────────────────────────────────
# ALMACÉN DE TOKENS EN MEMORIA
# Token = clave generada al login, expira en 8 horas
# ─────────────────────────────────────────────

_tokens: dict[str, dict] = {}   # { token: { usuario_data, expira } }


def _generar_token(usuario_data: dict) -> str:
    """Crea un token seguro y lo almacena con expiración."""
    token = secrets.token_hex(32)
    _tokens[token] = {
        'usuario': usuario_data,
        'expira':  datetime.now() + timedelta(hours=8)
    }
    return token


def _validar_token(token: str) -> dict | None:
    """Retorna los datos del usuario si el token es válido y no expiró."""
    entry = _tokens.get(token)
    if not entry:
        return None
    if datetime.now() > entry['expira']:
        del _tokens[token]
        return None
    return entry['usuario']


# ─────────────────────────────────────────────
# DECORADOR: requiere token válido en header
# Header esperado:  Authorization: Bearer <token>
# ─────────────────────────────────────────────

def requiere_token(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        auth = request.headers.get('Authorization', '')
        if not auth.startswith('Bearer '):
            return jsonify({'error': 'Token requerido. Header: Authorization: Bearer <token>'}), 401
        token = auth.split(' ', 1)[1]
        usuario = _validar_token(token)
        if not usuario:
            return jsonify({'error': 'Token inválido o expirado. Vuelve a hacer login.'}), 401
        request.usuario_actual = usuario
        return f(*args, **kwargs)
    return wrapper


def requiere_admin(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        auth = request.headers.get('Authorization', '')
        if not auth.startswith('Bearer '):
            return jsonify({'error': 'Token requerido.'}), 401
        token = auth.split(' ', 1)[1]
        usuario = _validar_token(token)
        if not usuario:
            return jsonify({'error': 'Token inválido o expirado.'}), 401
        if usuario.get('usuario') != 'admin':
            return jsonify({'error': 'Acción reservada para el administrador.'}), 403
        request.usuario_actual = usuario
        return f(*args, **kwargs)
    return wrapper


# ─────────────────────────────────────────────
# ENDPOINTS
# ─────────────────────────────────────────────

@app.route('/health', methods=['GET'])
def health():
    """Verificación de estado del servidor."""
    return jsonify({
        'status':  'ok',
        'servicio': 'Firma Digital API',
        'version':  '1.0',
        'hora':     datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    })


# ── POST /login ───────────────────────────────────────────────────────────────
@app.route('/login', methods=['POST'])
def login():
    """
    Autentica al usuario y retorna un token de sesión.

    Body JSON:
        { "usuario": "jperez", "password": "Firma2026!" }

    Respuesta exitosa:
        { "token": "...", "nombre": "Juan Pérez", "expira_en": "8 horas" }
    """
    data = request.get_json(silent=True)
    if not data or not data.get('usuario') or not data.get('password'):
        return jsonify({'error': 'Se requieren los campos: usuario, password'}), 400

    resultado = verificar_credenciales(data['usuario'], data['password'])

    if resultado is None:
        return jsonify({'error': 'Credenciales incorrectas o usuario no existe.'}), 401

    if isinstance(resultado, dict) and resultado.get('bloqueado'):
        return jsonify({'error': 'Usuario bloqueado por múltiples intentos fallidos. Contacta al administrador.'}), 403

    token = _generar_token(resultado)
    return jsonify({
        'ok':        True,
        'token':     token,
        'nombre':    resultado['nombre'],
        'usuario':   resultado['usuario'],
        'expira_en': '8 horas'
    })


# ── POST /firmar ──────────────────────────────────────────────────────────────
@app.route('/firmar', methods=['POST'])
@requiere_token
def firmar():
    """
    Genera una firma digital para un documento.
    Requiere token en header: Authorization: Bearer <token>

    Body JSON:
        { "documento": "contrato_enero.xlsx" }

    Respuesta exitosa:
        {
          "ok": true,
          "firma": {
            "nombre":     "Juan Pérez",
            "fecha_hora": "2026-05-04 10:30",
            "id_firma":   "A82K91LX",
            "documento":  "contrato_enero.xlsx"
          },
          "texto": "──────...\nFirmado por: ..."
        }
    """
    data = request.get_json(silent=True)
    if not data or not data.get('documento'):
        return jsonify({'error': 'Se requiere el campo: documento'}), 400

    usuario_data = request.usuario_actual
    documento    = data['documento'].strip()

    try:
        from modules.firma import crear_firma, formatear_firma_texto
        firma = crear_firma(usuario_data, documento)
        texto = formatear_firma_texto(firma)
        return jsonify({
            'ok':     True,
            'firma':  {
                'nombre':     firma['nombre'],
                'fecha_hora': firma['fecha_hora'],
                'id_firma':   firma['id_firma'],
                'documento':  firma['documento'],
            },
            'texto_firma': texto
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── GET /verificar/<id_firma> ─────────────────────────────────────────────────
@app.route('/verificar/<string:id_firma>', methods=['GET'])
def verificar(id_firma: str):
    """
    Verifica si una firma existe en la base de datos.
    No requiere autenticación (es una consulta pública de trazabilidad).

    Respuesta si existe:
        { "valida": true, "datos": { usuario, fecha_hora, documento, id_firma } }

    Respuesta si NO existe:
        { "valida": false, "mensaje": "Firma no encontrada" }
    """
    resultado = validar_firma(id_firma.upper())
    if resultado['valida']:
        return jsonify({
            'valida': True,
            'datos':  resultado['datos']
        })
    return jsonify({
        'valida':  False,
        'mensaje': 'Firma no encontrada en la base de datos.'
    }), 404


# ── GET /historial/<usuario> ──────────────────────────────────────────────────
@app.route('/historial/<string:usuario>', methods=['GET'])
@requiere_token
def historial(usuario: str):
    """
    Retorna el historial de firmas de un usuario.
    Solo el propio usuario o el admin pueden ver el historial.
    Requiere token en header.
    """
    actual = request.usuario_actual
    if actual['usuario'] != usuario and actual['usuario'] != 'admin':
        return jsonify({'error': 'No autorizado para ver el historial de otro usuario.'}), 403

    firmas = obtener_firmas_usuario(usuario)
    return jsonify({
        'ok':     True,
        'usuario': usuario,
        'total':   len(firmas),
        'firmas':  firmas
    })


# ── POST /usuarios/crear ──────────────────────────────────────────────────────
@app.route('/usuarios/crear', methods=['POST'])
@requiere_admin
def crear_usuario_endpoint():
    """
    Crea un nuevo usuario. Solo accesible por el administrador.
    Requiere token de admin en header.

    Body JSON:
        { "usuario": "nuevo", "password": "Pass123!", "nombre": "Nombre Completo" }
    """
    data = request.get_json(silent=True)
    if not data or not all(k in data for k in ('usuario', 'password', 'nombre')):
        return jsonify({'error': 'Se requieren: usuario, password, nombre'}), 400

    creado = crear_usuario(data['usuario'], data['password'], data['nombre'])
    if creado:
        return jsonify({'ok': True, 'mensaje': f"Usuario '{data['usuario']}' creado correctamente."})
    return jsonify({'error': f"El usuario '{data['usuario']}' ya existe."}), 409


# ── POST /logout ──────────────────────────────────────────────────────────────
@app.route('/logout', methods=['POST'])
@requiere_token
def logout():
    """Invalida el token actual."""
    auth  = request.headers.get('Authorization', '')
    token = auth.split(' ', 1)[1]
    if token in _tokens:
        del _tokens[token]
    return jsonify({'ok': True, 'mensaje': 'Sesión cerrada.'})


# ─────────────────────────────────────────────
# ARRANQUE
# ─────────────────────────────────────────────

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='API de Firma Digital')
    parser.add_argument('--host', default='127.0.0.1', help='Host (default: 127.0.0.1)')
    parser.add_argument('--port', type=int, default=5000, help='Puerto (default: 5000)')
    parser.add_argument('--debug', action='store_true', help='Modo debug')
    args = parser.parse_args()

    print("=" * 50)
    print("  🔏  Firma Digital — API REST v1.0")
    print("=" * 50)
    print(f"\n  Servidor: http://{args.host}:{args.port}")
    print(f"  Endpoints disponibles:")
    print(f"    POST  /login")
    print(f"    POST  /firmar           (token)")
    print(f"    GET   /verificar/<id>")
    print(f"    GET   /historial/<usr>  (token)")
    print(f"    POST  /usuarios/crear   (admin)")
    print(f"    POST  /logout           (token)")
    print(f"    GET   /health")
    print(f"\n  Presiona Ctrl+C para detener\n")

    inicializar_db()
    app.run(host=args.host, port=args.port, debug=args.debug)
