"""
build_exe.py — Generador de ejecutables con diagnóstico completo
================================================================
Uso:
    python build_exe.py           <- construye los 3 .exe
    python build_exe.py --app     <- solo la app GUI
    python build_exe.py --api     <- solo la API
    python build_exe.py --celda   <- solo el firmador de celda
    python build_exe.py --debug   <- muestra salida completa de PyInstaller
"""

import subprocess
import sys
import os
import shutil
import argparse

# ─────────────────────────────────────────────
# RAÍZ DEL PROYECTO
# ─────────────────────────────────────────────
_ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(_ROOT)

# Separador de rutas para --add-data:  ';' en Windows, ':' en Linux/Mac
SEP = ';' if sys.platform == 'win32' else ':'

# ─────────────────────────────────────────────
# HIDDEN IMPORTS COMUNES A TODOS LOS .EXE
# ─────────────────────────────────────────────
HIDDEN = [
    'modules.db', 'modules.firma', 'modules.login',
    'modules.app', 'modules.utils',
    'openpyxl', 'openpyxl.styles', 'openpyxl.utils',
    'openpyxl.utils.cell', 'openpyxl.writer.excel',
    'openpyxl.reader.excel',
    'docx', 'docx.shared', 'docx.enum.text',
    'docx.oxml', 'docx.oxml.ns', 'docx.oxml.shared',
    'flask', 'flask.json', 'werkzeug', 'werkzeug.serving',
    'jinja2', 'click', 'itsdangerous',
    'sqlite3', 'hashlib', 'secrets', 'getpass',
    'tkinter', 'tkinter.ttk', 'tkinter.messagebox',
    'tkinter.filedialog', '_tkinter',
    'email.mime.text', 'email.mime.multipart',
]

# ─────────────────────────────────────────────
# DEFINICIÓN DE CADA EJECUTABLE
# ─────────────────────────────────────────────
BUILDS = {
    'app': {
        'script': 'main.py',
        'nombre': 'FirmaDigital',
        'desc':   'App de escritorio (GUI)',
        'extra':  ['--windowed'],
    },
    'api': {
        'script': 'api.py',
        'nombre': 'FirmaDigital_API',
        'desc':   'Servidor API REST',
        'extra':  [],
    },
    'celda': {
        'script': 'firma_celda.py',
        'nombre': 'FirmaDigital_Celda',
        'desc':   'Firmador de celda Excel',
        'extra':  [],
    },
}


# ─────────────────────────────────────────────
# VERIFICACIONES PREVIAS
# ─────────────────────────────────────────────

def verificar_entorno():
    """Verifica que todos los archivos y paquetes necesarios existen."""
    print("\n[DIAGNOSTICO] Verificando entorno...\n")
    ok = True

    # Archivos del proyecto
    archivos = [
        'main.py', 'api.py', 'firma_celda.py',
        os.path.join('modules', '__init__.py'),
        os.path.join('modules', 'db.py'),
        os.path.join('modules', 'firma.py'),
        os.path.join('modules', 'login.py'),
        os.path.join('modules', 'app.py'),
        os.path.join('modules', 'utils.py'),
    ]
    for f in archivos:
        ruta = os.path.join(_ROOT, f)
        if os.path.exists(ruta):
            print(f"  OK  {f}")
        else:
            print(f"  XX  {f}  <-- FALTA ESTE ARCHIVO")
            ok = False

    print()

    # Paquetes Python
    paquetes = {
        'PyInstaller': 'pyinstaller',
        'openpyxl':    'openpyxl',
        'docx':        'python-docx',
        'flask':       'flask',
    }
    for modulo, pip_nombre in paquetes.items():
        try:
            __import__(modulo)
            print(f"  OK  {modulo}")
        except ImportError:
            print(f"  XX  {modulo}  <-- ejecuta: pip install {pip_nombre}")
            ok = False

    print()
    return ok


def instalar_dependencias():
    """Instala paquetes faltantes automáticamente."""
    deps = ['openpyxl', 'python-docx', 'flask', 'pyinstaller']
    print("[INSTALANDO] Dependencias...\n")
    for dep in deps:
        print(f"  pip install {dep} ...")
        r = subprocess.run(
            [sys.executable, '-m', 'pip', 'install', dep, '-q'],
            capture_output=True, text=True
        )
        if r.returncode == 0:
            print(f"  OK  {dep}")
        else:
            print(f"  XX  {dep}")
            print(f"      {r.stderr.strip()[:200]}")
    print()


# ─────────────────────────────────────────────
# CONSTRUCCIÓN
# ─────────────────────────────────────────────

def construir(clave: str, mostrar_log: bool = False) -> bool:
    cfg = BUILDS[clave]

    print(f"\n{'='*54}")
    print(f"  Compilando: {cfg['nombre']}.exe  ({cfg['desc']})")
    print(f"{'='*54}")

    modules_src = os.path.join(_ROOT, 'modules')
    modules_dst = 'modules'

    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',
        f'--name={cfg["nombre"]}',
        f'--distpath={os.path.join(_ROOT, "dist")}',
        f'--workpath={os.path.join(_ROOT, "build_tmp")}',
        f'--specpath={os.path.join(_ROOT, "build_tmp")}',
        '--noconfirm',
        '--clean',
        f'--add-data={modules_src}{SEP}{modules_dst}',
    ]

    # Hidden imports
    for hi in HIDDEN:
        cmd.append(f'--hidden-import={hi}')

    # Argumentos específicos del ejecutable (ej: --windowed)
    cmd += cfg['extra']

    # Script de entrada
    cmd.append(os.path.join(_ROOT, cfg['script']))

    print(f"  Comando PyInstaller:")
    print(f"  {' '.join(cmd[2:4])} {cfg['nombre']} {cfg['script']}\n")

    # Ejecutar PyInstaller
    if mostrar_log:
        # Salida completa visible
        result = subprocess.run(cmd, cwd=_ROOT)
    else:
        # Filtrar solo líneas importantes
        result = subprocess.run(
            cmd, cwd=_ROOT,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace'
        )
        lineas_clave = []
        for linea in result.stdout.splitlines():
            l = linea.strip()
            if any(k in l for k in ('ERROR', 'WARNING: lib', 'Building EXE',
                                     'completed successfully', 'CRITICAL', 'ModuleNotFound')):
                lineas_clave.append(linea)
        if lineas_clave:
            print("  Log relevante:")
            for l in lineas_clave:
                color = '  !! ' if 'ERROR' in l or 'CRITICAL' in l else '  -- '
                print(f"{color}{l.strip()}")

        # Si falló, mostrar últimas 50 líneas para diagnóstico
        if result.returncode != 0:
            print("\n  [LOG COMPLETO - ultimas 50 lineas]")
            for l in result.stdout.splitlines()[-50:]:
                print(f"  {l}")
            # También indicar dónde está el log completo
            log_path = os.path.join(_ROOT, 'build_tmp', cfg['nombre'], 'warn-' + cfg['nombre'] + '.txt')
            if os.path.exists(log_path):
                print(f"\n  Log completo en: {log_path}")

    # En Windows genera .exe, en Linux/Mac sin extensión
    ext      = '.exe' if sys.platform == 'win32' else ''
    exe      = os.path.join(_ROOT, 'dist', cfg['nombre'] + ext)
    exe_show = f"dist/{cfg['nombre']}{ext}"

    if os.path.exists(exe):
        mb = os.path.getsize(exe) / (1024 * 1024)
        print(f"\n  OK  {exe_show}  ({mb:.1f} MB)")
        return True
    else:
        print(f"\n  XX  {exe_show}  -- no se genero")
        print(f"      Vuelve a ejecutar con --debug para ver el log completo.")
        return False


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Build de ejecutables - Firma Digital')
    parser.add_argument('--app',    action='store_true', help='Solo app GUI')
    parser.add_argument('--api',    action='store_true', help='Solo API REST')
    parser.add_argument('--celda',  action='store_true', help='Solo firmador de celda')
    parser.add_argument('--debug',  action='store_true', help='Mostrar log completo de PyInstaller')
    parser.add_argument('--instalar', action='store_true', help='Instalar dependencias antes de compilar')
    args = parser.parse_args()

    print("=" * 54)
    print("  Firma Digital — Build de Ejecutables")
    print(f"  Python: {sys.version.split()[0]}  |  Plataforma: {sys.platform}")
    print(f"  Raiz del proyecto: {_ROOT}")
    print("=" * 54)

    # Instalar dependencias si se pide
    if args.instalar:
        instalar_dependencias()

    # Diagnóstico del entorno
    if not verificar_entorno():
        print("  ERRORES encontrados arriba. Corrigelos antes de continuar.")
        print("  Consejo: ejecuta con --instalar para instalar paquetes faltantes.")
        sys.exit(1)

    # Determinar qué compilar
    todo = not (args.app or args.api or args.celda)
    targets = []
    if todo or args.app:   targets.append('app')
    if todo or args.api:   targets.append('api')
    if todo or args.celda: targets.append('celda')

    # Limpiar carpetas temporales
    for carpeta in ['build_tmp']:
        p = os.path.join(_ROOT, carpeta)
        if os.path.exists(p):
            shutil.rmtree(p)
            print(f"  Limpiado: {carpeta}/")

    os.makedirs(os.path.join(_ROOT, 'dist'), exist_ok=True)

    # Compilar
    resultados = {}
    for t in targets:
        resultados[t] = construir(t, mostrar_log=args.debug)

    # Resumen
    print(f"\n{'='*54}")
    print("  RESUMEN")
    print(f"{'='*54}")
    exitosos = sum(1 for v in resultados.values() if v)
    for t, ok in resultados.items():
        nombre = BUILDS[t]['nombre']
        ext    = '.exe' if sys.platform == 'win32' else ''
        estado = "OK " if ok else "XX "
        print(f"  {estado}  dist/{nombre}{ext}")

    print()
    if exitosos == len(resultados):
        print("  Todos los .exe generados en dist/")
        print()
        print("  PARA DISTRIBUIR:")
        print("    Copia la carpeta dist/ a los equipos destino.")
        print("    No requieren Python instalado.")
        print()
        print("  NOTA: La BD firma_digital.db se crea sola")
        print("        en la misma carpeta que el .exe al primer arranque.")
    else:
        fallidos = len(resultados) - exitosos
        print(f"  {fallidos} ejecutable(s) fallaron.")
        print()
        print("  PARA VER EL ERROR COMPLETO ejecuta:")
        print(f"    python build_exe.py --debug")
        print()
        print("  CAUSAS COMUNES:")
        print("    1. Falta algun archivo en modules/")
        print("    2. Falta instalar un paquete: python build_exe.py --instalar")
        print("    3. Antivirus bloqueando PyInstaller (desactivar temporalmente)")
        sys.exit(1)


if __name__ == '__main__':
    main()