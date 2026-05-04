"""
firma_celda.py — Inserción de firma en celda específica de Excel
================================================================
Uso independiente: inserta la firma de un usuario autenticado
en la celda exacta que tú indiques de un archivo Excel.

Modos de uso:
─────────────────────────────────────────────────────────────
1) Como script directo (interactivo):
       python firma_celda.py

2) Como módulo importado desde otro script:
       from firma_celda import firmar_celda_excel
       firmar_celda_excel(
           ruta_archivo = "mi_reporte.xlsx",
           celda        = "F10",
           usuario      = "jperez",
           password     = "Firma2026!",
           hoja         = "Hoja1",          # opcional
           modo         = "bloque"          # "bloque" | "linea" | "id"
       )

3) Desde línea de comandos con argumentos:
       python firma_celda.py --archivo reporte.xlsx --celda F10
                             --usuario jperez --hoja "Hoja1" --modo bloque
─────────────────────────────────────────────────────────────

Modos de inserción:
  bloque  → Escribe 3 celdas hacia abajo (Firmado por / Fecha / ID)   [default]
  linea   → Todo en una sola celda:  "Juan Pérez | 2026-05-04 | A82K91LX"
  id      → Solo el ID de firma:     "A82K91LX"
"""

import sys
import os
import argparse
import getpass

# ── Fix de path ───────────────────────────────────────────────────────────────
_ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if _ROOT_DIR not in sys.path:
    sys.path.insert(0, _ROOT_DIR)
os.chdir(_ROOT_DIR)
# ─────────────────────────────────────────────────────────────────────────────

from modules.db    import inicializar_db, verificar_credenciales
from modules.firma import crear_firma


# ═════════════════════════════════════════════
# FUNCIÓN PRINCIPAL EXPORTABLE
# ═════════════════════════════════════════════

def firmar_celda_excel(
    ruta_archivo: str,
    celda:        str,
    usuario:      str,
    password:     str,
    hoja:         str  = None,
    modo:         str  = "bloque",
    crear_si_no_existe: bool = True,
) -> dict:
    """
    Inserta una firma digital en una celda específica de un archivo Excel.

    Parámetros
    ──────────
    ruta_archivo        Ruta al .xlsx  (ej: "C:/docs/reporte.xlsx")
    celda               Celda destino  (ej: "F10", "B3", "A1")
    usuario             Nombre de usuario registrado en el sistema
    password            Contraseña del usuario (nunca se guarda)
    hoja                Nombre de la hoja. None = hoja activa
    modo                Estilo de inserción:
                          "bloque"  → 3 filas: nombre / fecha / ID  [recomendado]
                          "linea"   → 1 celda con todo
                          "id"      → solo el ID de firma
    crear_si_no_existe  Si True, crea el archivo si no existe

    Retorna
    ───────
    dict con:
      ok        → True/False
      mensaje   → descripción del resultado
      firma     → dict con los datos de la firma (si ok=True)
      celda     → celda donde se insertó (si ok=True)
      archivo   → ruta del archivo modificado (si ok=True)
      error     → descripción del error (si ok=False)
    """

    # ── 1. Validar modo ───────────────────────────────────────────────────
    modos_validos = ("bloque", "linea", "id")
    if modo not in modos_validos:
        return {'ok': False, 'error': f"Modo '{modo}' no válido. Usa: {modos_validos}"}

    # ── 2. Autenticar usuario ─────────────────────────────────────────────
    resultado_auth = verificar_credenciales(usuario, password)

    if resultado_auth is None:
        return {'ok': False, 'error': 'Credenciales incorrectas o usuario no existe.'}

    if isinstance(resultado_auth, dict) and resultado_auth.get('bloqueado'):
        return {'ok': False, 'error': 'Usuario bloqueado. Contacta al administrador.'}

    usuario_data = resultado_auth

    # ── 3. Generar firma y registrar en DB ────────────────────────────────
    nombre_doc = os.path.basename(ruta_archivo)
    firma = crear_firma(usuario_data, nombre_doc)

    # ── 4. Abrir / crear archivo Excel ────────────────────────────────────
    try:
        from openpyxl import load_workbook, Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    except ImportError:
        return {'ok': False, 'error': 'openpyxl no instalado. Ejecuta: pip install openpyxl'}

    if os.path.exists(ruta_archivo):
        wb = load_workbook(ruta_archivo)
    elif crear_si_no_existe:
        wb = Workbook()
    else:
        return {'ok': False, 'error': f"Archivo no encontrado: {ruta_archivo}"}

    # Seleccionar hoja
    if hoja and hoja in wb.sheetnames:
        ws = wb[hoja]
    else:
        ws = wb.active

    # ── 5. Parsear celda de destino ───────────────────────────────────────
    celda = celda.upper().strip()
    try:
        col_letra, fila_num = coordinate_from_string(celda)
        col_num = column_index_from_string(col_letra)
    except Exception:
        return {'ok': False, 'error': f"Celda '{celda}' no es válida. Usa formato Excel: A1, B3, F10..."}

    # ── 6. Insertar según el modo elegido ─────────────────────────────────
    if modo == "bloque":
        _insertar_bloque(ws, fila_num, col_num, firma)
    elif modo == "linea":
        _insertar_linea(ws, fila_num, col_num, firma)
    elif modo == "id":
        _insertar_solo_id(ws, fila_num, col_num, firma)

    # ── 7. Guardar ────────────────────────────────────────────────────────
    wb.save(ruta_archivo)

    return {
        'ok':      True,
        'mensaje': f"Firma insertada en {ws.title}!{celda} (modo: {modo})",
        'firma':   {
            'nombre':     firma['nombre'],
            'fecha_hora': firma['fecha_hora'],
            'id_firma':   firma['id_firma'],
            'documento':  firma['documento'],
        },
        'celda':   f"{ws.title}!{celda}",
        'archivo': ruta_archivo,
    }


# ═════════════════════════════════════════════
# MODOS DE INSERCIÓN
# ═════════════════════════════════════════════

def _estilos():
    """Retorna los estilos compartidos para todos los modos."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    COLOR_HEADER = "1F3864"
    COLOR_FILA   = "D6E4F0"

    return {
        'fuente_header': Font(name="Calibri", bold=True, color="FFFFFF", size=10),
        'fuente_label':  Font(name="Calibri", bold=True, size=10, color="1F3864"),
        'fuente_valor':  Font(name="Calibri", size=10),
        'relleno_header': PatternFill("solid", fgColor=COLOR_HEADER),
        'relleno_fila':   PatternFill("solid", fgColor=COLOR_FILA),
        'alineacion':     Alignment(horizontal="left", vertical="center", wrap_text=True),
        'borde': Border(
            left=Side(style='thin'),   right=Side(style='thin'),
            top=Side(style='thin'),    bottom=Side(style='thin')
        ),
    }


def _insertar_bloque(ws, fila: int, col: int, firma: dict):
    """
    Modo BLOQUE: inserta la firma en 4 filas (encabezado + 3 datos).

    Ejemplo en celda F10:
        F10  │ 🔏 FIRMA DIGITAL        │ (encabezado azul, merge F10:G10)
        F11  │ Firmado por  │ Juan Pérez
        F12  │ Fecha        │ 2026-05-04 10:30
        F13  │ ID Firma     │ A82K91LX
    """
    from openpyxl.styles import Alignment

    s = _estilos()

    # Fila 0: encabezado con merge de 2 columnas
    try:
        ws.merge_cells(
            start_row=fila,    start_column=col,
            end_row=fila,      end_column=col + 1
        )
    except Exception:
        pass  # Si ya hay merge previo, ignorar

    c_enc = ws.cell(row=fila, column=col, value="🔏 FIRMA DIGITAL")
    c_enc.font      = s['fuente_header']
    c_enc.fill      = s['relleno_header']
    c_enc.alignment = Alignment(horizontal="center", vertical="center")
    c_enc.border    = s['borde']
    ws.row_dimensions[fila].height = 18

    # Filas 1-3: etiqueta | valor
    datos = [
        ("Firmado por", firma['nombre']),
        ("Fecha",       firma['fecha_hora']),
        ("ID Firma",    firma['id_firma']),
    ]

    for offset, (etiqueta, valor) in enumerate(datos, start=1):
        f = fila + offset

        c_label = ws.cell(row=f, column=col,     value=etiqueta)
        c_valor = ws.cell(row=f, column=col + 1, value=valor)

        c_label.font      = s['fuente_label']
        c_label.fill      = s['relleno_fila']
        c_label.alignment = s['alineacion']
        c_label.border    = s['borde']

        c_valor.font      = s['fuente_valor']
        c_valor.fill      = s['relleno_fila']
        c_valor.alignment = s['alineacion']
        c_valor.border    = s['borde']

        ws.row_dimensions[f].height = 15

    # Ajustar anchos mínimos de columna
    from openpyxl.utils import get_column_letter
    col_letra_a = get_column_letter(col)
    col_letra_b = get_column_letter(col + 1)
    if ws.column_dimensions[col_letra_a].width < 14:
        ws.column_dimensions[col_letra_a].width = 14
    if ws.column_dimensions[col_letra_b].width < 26:
        ws.column_dimensions[col_letra_b].width = 26


def _insertar_linea(ws, fila: int, col: int, firma: dict):
    """
    Modo LINEA: todo en una sola celda, separado por " | ".

    Ejemplo en celda F10:
        F10  │ Juan Pérez | 2026-05-04 10:30 | A82K91LX
    """
    from openpyxl.styles import Alignment

    s = _estilos()

    texto = f"{firma['nombre']}  |  {firma['fecha_hora']}  |  ID: {firma['id_firma']}"

    c = ws.cell(row=fila, column=col, value=texto)
    c.font      = s['fuente_valor']
    c.fill      = s['relleno_fila']
    c.border    = s['borde']
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[fila].height = 16

    # Ampliar columna si es necesaria
    from openpyxl.utils import get_column_letter
    col_letra = get_column_letter(col)
    ancho_necesario = len(texto) * 0.9
    if ws.column_dimensions[col_letra].width < ancho_necesario:
        ws.column_dimensions[col_letra].width = min(ancho_necesario, 80)


def _insertar_solo_id(ws, fila: int, col: int, firma: dict):
    """
    Modo ID: solo el ID de firma en la celda destino.

    Ejemplo en celda F10:
        F10  │ A82K91LX
    """
    from openpyxl.styles import Alignment, Font

    c = ws.cell(row=fila, column=col, value=firma['id_firma'])
    c.font      = Font(name="Courier New", bold=True, size=10, color="1F3864")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border    = _estilos()['borde']
    ws.row_dimensions[fila].height = 15


# ═════════════════════════════════════════════
# MODO INTERACTIVO (cuando se ejecuta directo)
# ═════════════════════════════════════════════

def _modo_interactivo():
    """Guía al usuario paso a paso por la firma de una celda específica."""
    print("\n" + "═" * 52)
    print("  🔏  Firma Digital — Celda Específica en Excel")
    print("═" * 52)

    # Archivo
    print("\n📂 Ruta del archivo Excel (.xlsx):")
    print("   Ejemplo: C:\\Users\\HP\\Documents\\reporte.xlsx")
    ruta = input("   → ").strip().strip('"')
    if not ruta:
        print("❌ Ruta vacía. Saliendo.")
        return

    # Hoja
    print("\n📋 Nombre de la hoja (Enter = hoja activa):")
    hoja = input("   → ").strip() or None

    # Celda
    print("\n📌 Celda de destino (ej: F10, B3, A1):")
    celda = input("   → ").strip()
    if not celda:
        print("❌ Celda vacía. Saliendo.")
        return

    # Modo
    print("\n🎨 Modo de inserción:")
    print("   1  bloque  → encabezado + 3 filas [recomendado]")
    print("   2  linea   → todo en 1 celda")
    print("   3  id      → solo el ID de firma")
    opcion = input("   → ").strip()
    modos_map = {'1': 'bloque', '2': 'linea', '3': 'id',
                 'bloque': 'bloque', 'linea': 'linea', 'id': 'id'}
    modo = modos_map.get(opcion, 'bloque')
    print(f"   Modo seleccionado: {modo}")

    # Credenciales
    print("\n🔐 Credenciales de firma:")
    usuario  = input("   Usuario  → ").strip()
    password = getpass.getpass("   Contraseña → ")   # oculta en terminal

    # Ejecutar
    print("\n⏳ Procesando...")
    resultado = firmar_celda_excel(
        ruta_archivo = ruta,
        celda        = celda,
        usuario      = usuario,
        password     = password,
        hoja         = hoja,
        modo         = modo,
    )

    # Mostrar resultado
    print()
    if resultado['ok']:
        f = resultado['firma']
        print("✅ FIRMA INSERTADA EXITOSAMENTE")
        print(f"   Archivo:     {resultado['archivo']}")
        print(f"   Celda:       {resultado['celda']}")
        print(f"   Firmado por: {f['nombre']}")
        print(f"   Fecha:       {f['fecha_hora']}")
        print(f"   ID Firma:    {f['id_firma']}")
    else:
        print(f"❌ ERROR: {resultado['error']}")

    input("\nPresiona Enter para salir...")


# ═════════════════════════════════════════════
# LÍNEA DE COMANDOS
# ═════════════════════════════════════════════

def _modo_cli():
    parser = argparse.ArgumentParser(
        description='Inserta una firma digital en una celda específica de Excel.',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('--archivo',  required=True,  help='Ruta al archivo .xlsx')
    parser.add_argument('--celda',    required=True,  help='Celda destino (ej: F10, B3)')
    parser.add_argument('--usuario',  required=True,  help='Nombre de usuario del sistema')
    parser.add_argument('--hoja',     default=None,   help='Nombre de hoja (default: hoja activa)')
    parser.add_argument('--modo',     default='bloque',
                        choices=['bloque', 'linea', 'id'],
                        help='Modo de inserción (default: bloque)')
    args = parser.parse_args()

    # Contraseña siempre oculta, incluso en CLI
    password = getpass.getpass(f"Contraseña para '{args.usuario}': ")

    resultado = firmar_celda_excel(
        ruta_archivo = args.archivo,
        celda        = args.celda,
        usuario      = args.usuario,
        password     = password,
        hoja         = args.hoja,
        modo         = args.modo,
    )

    if resultado['ok']:
        print(f"✅ {resultado['mensaje']}")
        print(f"   ID Firma: {resultado['firma']['id_firma']}")
        sys.exit(0)
    else:
        print(f"❌ Error: {resultado['error']}")
        sys.exit(1)


# ═════════════════════════════════════════════
# PUNTO DE ENTRADA
# ═════════════════════════════════════════════

if __name__ == '__main__':
    inicializar_db()

    # Si recibe argumentos CLI, usar modo CLI
    if len(sys.argv) > 1:
        _modo_cli()
    else:
        # Sin argumentos → modo interactivo amigable
        _modo_interactivo()
