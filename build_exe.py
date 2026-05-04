"""
build_exe.py — Generador automático de ejecutables
Crea los tres .exe del sistema con PyInstaller.

Uso:
    python build_exe.py           ← construye los 3 ejecutables
    python build_exe.py --app     ← solo la app de escritorio
    python build_exe.py --api     ← solo la API
    python build_exe.py --celda   ← solo el firmador de celda
"""

import subprocess
import sys
import os
import shutil
import argparse

_ROOT = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────
# DEFINICIÓN DE EJECUTABLES
# ─────────────────────────────────────────────

BUILDS = {
    'app': {
        'script':  'main.py',
        'nombre':  'FirmaDigital',
        'desc':    'Aplicación de escritorio (GUI)',
        'args':    ['--windowed'],       # Sin ventana de consola
    },
    'api': {
        'script':  'api.py',
        'nombre':  'FirmaDigital_API',
        'desc':    'Servidor API REST',
        'args':    [],                   # Con consola (muestra logs del servidor)
    },
    'celda': {
        'script':  'firma_celda.py',
        'nombre':  'FirmaDigital_Celda',
        'desc':    'Firmador de celda específica en Excel',
        'args':    [],                   # Con consola (interactivo)
    },
}


# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────

def verificar_pyinstaller():
    """Verifica que PyInstaller esté instalado, lo instala si no."""
    try:
        import PyInstaller
        print(f"  ✓ PyInstaller {PyInstaller.__version__} encontrado.")
        return True
    except ImportError:
        print("  ⚠ PyInstaller no encontrado. Instalando...")
        result = subprocess.run(
            [sys.executable, '-m', 'pip', 'install', 'pyinstaller'],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            print("  ✓ PyInstaller instalado correctamente.")
            return True
        else:
            print(f"  ❌ Error instalando PyInstaller:\n{result.stderr}")
            return False


def limpiar_build():
    """Elimina carpetas temporales de builds anteriores."""
    for carpeta in ('build', '__pycache__'):
        if os.path.exists(carpeta):
            shutil.rmtree(carpeta)
            print(f"  🗑  Limpiado: {carpeta}/")


def construir(clave: str) -> bool:
    """
    Construye un ejecutable específico.
    Retorna True si tuvo éxito.
    """
    cfg = BUILDS[clave]
    print(f"\n{'─'*50}")
    print(f"  🔨 Construyendo: {cfg['nombre']}  ({cfg['desc']})")
    print(f"{'─'*50}")

    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',                                    # Un solo .exe
        f"--name={cfg['nombre']}",                     # Nombre del ejecutable
        '--distpath=dist',                             # Carpeta de salida
        '--workpath=build',                            # Carpeta temporal
        '--specpath=build',                            # Specs en build/
        '--noconfirm',                                 # Sobreescribir sin preguntar
        '--clean',                                     # Limpiar caché antes
        # Incluir la carpeta modules/ completa
        '--add-data=modules;modules',
    ] + cfg['args'] + [cfg['script']]

    print(f"  Comando: {' '.join(cmd[2:])}\n")

    result = subprocess.run(cmd, cwd=_ROOT)

    if result.returncode == 0:
        exe_path = os.path.join(_ROOT, 'dist', cfg['nombre'] + '.exe')
        size_mb  = os.path.getsize(exe_path) / (1024 * 1024) if os.path.exists(exe_path) else 0
        print(f"\n  ✅ Creado: dist/{cfg['nombre']}.exe  ({size_mb:.1f} MB)")
        return True
    else:
        print(f"\n  ❌ Falló la construcción de {cfg['nombre']}")
        return False


# ─────────────────────────────────────────────
# PUNTO DE ENTRADA
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Constructor de ejecutables - Firma Digital')
    parser.add_argument('--app',   action='store_true', help='Construir solo la app GUI')
    parser.add_argument('--api',   action='store_true', help='Construir solo la API')
    parser.add_argument('--celda', action='store_true', help='Construir solo el firmador de celda')
    args = parser.parse_args()

    # Si no se especifica nada, construir todo
    construir_todo = not (args.app or args.api or args.celda)
    targets = []
    if construir_todo or args.app:   targets.append('app')
    if construir_todo or args.api:   targets.append('api')
    if construir_todo or args.celda: targets.append('celda')

    print("=" * 50)
    print("  🔏  Firma Digital — Build de Ejecutables")
    print("=" * 50)

    os.chdir(_ROOT)

    # Verificar PyInstaller
    if not verificar_pyinstaller():
        sys.exit(1)

    # Limpiar residuos anteriores
    print("\n[1/3] Limpiando builds anteriores...")
    limpiar_build()

    # Construir cada target
    resultados = {}
    for target in targets:
        resultados[target] = construir(target)

    # Resumen final
    print(f"\n{'═'*50}")
    print("  📦  Resumen del build")
    print(f"{'═'*50}")
    ok_count = 0
    for target, ok in resultados.items():
        estado = "✅" if ok else "❌"
        nombre = BUILDS[target]['nombre']
        print(f"  {estado}  dist/{nombre}.exe")
        if ok:
            ok_count += 1

    if ok_count == len(resultados):
        print(f"\n  ✅ Todos los ejecutables generados en dist/")
        print(f"\n  📋 INSTRUCCIONES DE DISTRIBUCIÓN:")
        print(f"     1. Copia la carpeta dist/ al equipo destino")
        print(f"     2. Los .exe son autocontenidos, no requieren Python")
        print(f"     3. La BD firma_digital.db se crea junto al .exe al primer arranque")
        print(f"     4. Para la API: FirmaDigital_API.exe --host 0.0.0.0 --port 5000")
    else:
        print(f"\n  ⚠  {len(resultados) - ok_count} ejecutable(s) fallaron.")
        sys.exit(1)


if __name__ == '__main__':
    main()
