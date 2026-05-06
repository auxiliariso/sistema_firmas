"""
buscador.py — Búsqueda dinámica y reemplazo de marcadores de firma
===================================================================
LÓGICA DE MARCADORES:
─────────────────────────────────────────────────────────────────────
El documento puede contener estos marcadores (no distingue mayúsculas):
  <<Firma_Juan Pérez>>              → reemplaza con el ID de firma de Juan Pérez
  <<Firma_Juan Pérez,Ana López>>    → reemplaza con IDs de ambas firmas
  <<Nombre_Juan Pérez>>             → reemplaza con el nombre del usuario
  <<Nombre_Juan Pérez>><<Firma_Juan Pérez>> → nombre + ID en misma celda/párrafo

DETECCIÓN DINÁMICA:
  Si una celda/párrafo contiene el nombre completo de un usuario de la DB
  pero NO tiene marcador explícito, el sistema lo detecta automáticamente
  y lo reemplaza con el ID de firma correspondiente.

REGLAS DE DETECCIÓN:
  1. Se cargan TODOS los nombres de usuarios desde la base de datos.
  2. Se buscan los marcadores <<Firma_*>> y <<Nombre_*>> en el documento.
  3. Se buscan también nombres de usuarios escritos en texto libre (detección dinámica).
  4. Se valida que el nombre dentro del marcador exista en la DB.
  5. Se muestra vista previa con los reemplazos resaltados.
  6. Doble clic → reemplaza los marcadores con los valores reales.

FLUJO:
  1. Abrir buscador desde la app principal (botón "🔎 Buscar Firma")
  2. Cargar archivo(s) Excel o Word
  3. El sistema detecta automáticamente los marcadores y nombres dinámicos
  4. Vista previa muestra qué se reemplazará por qué
  5. Doble clic en resultado → se aplica el reemplazo en el archivo
"""

import os
import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

_ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT_DIR not in sys.path:
    sys.path.insert(0, _ROOT_DIR)

from modules.db    import get_connection, guardar_firma
from modules.firma import generar_id_firma, crear_firma

# ═══════════════════════════════════════════════════════════════════
# CONSTANTES
# ═══════════════════════════════════════════════════════════════════

TIPO_EXCEL = "Excel (.xlsx)"
TIPO_WORD  = "Word (.docx)"
TIPO_OTRO  = "Otro"

ICONOS_TIPO  = {TIPO_EXCEL: "📊", TIPO_WORD: "📝", TIPO_OTRO: "📁"}
COLORES_TIPO = {TIPO_EXCEL: "#1E6B3C", TIPO_WORD: "#1E3A6B", TIPO_OTRO: "#888"}

# Patrones de marcadores (case-insensitive)
PAT_FIRMA  = re.compile(r'<<\s*Firma_([^>]+?)>>',          re.IGNORECASE)
PAT_NOMBRE = re.compile(r'<<\s*Nombre_([^>]+?)>>',         re.IGNORECASE)
PAT_AMBOS  = re.compile(r'<<\s*(?:Firma|Nombre)_([^>]+?)>>', re.IGNORECASE)

# ═══════════════════════════════════════════════════════════════════
# ACCESO A LA BASE DE DATOS
# ═══════════════════════════════════════════════════════════════════

def _obtener_todos_usuarios() -> dict:
    """
    Retorna { nombre_completo_lower: {'usuario': ..., 'nombre': ...} }
    para comparar contra los marcadores del documento.
    """
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT usuario, nombre FROM usuarios"
        ).fetchall()
    return {row['nombre'].lower(): dict(row) for row in rows}


def _obtener_firma_activa_usuario(nombre_completo: str, usuarios_db: dict) -> dict | None:
    key = nombre_completo.strip().lower()
    return usuarios_db.get(key)

# ═══════════════════════════════════════════════════════════════════
# ANÁLISIS DE MARCADORES EN EL TEXTO
# ═══════════════════════════════════════════════════════════════════

def _analizar_marcadores(texto: str, usuarios_db: dict) -> list:
    """
    Busca todos los marcadores <<Firma_*>> y <<Nombre_*>> en el texto.
    Retorna lista de dicts con metadatos de cada coincidencia.
    """
    resultados = []

    for linea_num, linea in enumerate(texto.splitlines(), start=1):
        for m in re.finditer(r'<<\s*(Firma|Nombre)_([^>]+?)>>', linea, re.IGNORECASE):
            tipo_marcador = m.group(1).lower()
            contenido     = m.group(2).strip()

            nombres_raw = [n.strip() for n in contenido.split(',') if n.strip()]
            validos   = []
            invalidos = []

            for nombre in nombres_raw:
                u = usuarios_db.get(nombre.lower())
                if u:
                    validos.append(u)
                else:
                    invalidos.append(nombre)

            if tipo_marcador == 'firma':
                if validos:
                    reemplazo = ' | '.join(f'«ID_{u["usuario"]}»' for u in validos)
                else:
                    reemplazo = '«USUARIO NO ENCONTRADO»'
            else:
                if validos:
                    reemplazo = ' / '.join(u['nombre'] for u in validos)
                else:
                    reemplazo = '«USUARIO NO ENCONTRADO»'

            resultados.append({
                'marcador_original': m.group(0),
                'tipo':              tipo_marcador,
                'nombres':           nombres_raw,
                'usuarios_validos':  validos,
                'usuarios_invalidos': invalidos,
                'reemplazo_preview': reemplazo,
                'linea_num':         linea_num,
                'linea_raw':         linea,
                'match_start':       m.start(),
                'match_end':         m.end(),
            })

    return resultados


def _detectar_nombres_dinamicos(texto: str, usuarios_db: dict) -> list:
    """
    Busca líneas que contengan el nombre completo de un usuario de la DB
    pero que NO tengan marcador <<Firma_*>> ni <<Nombre_*>>.
    Sirve para documentos donde el nombre está escrito en texto libre.

    Retorna lista de dicts con la misma estructura que _analizar_marcadores.
    """
    resultados = []
    # Nombres ya procesados en esta pasada (evitar duplicados por línea)
    procesados = set()

    for linea_num, linea in enumerate(texto.splitlines(), start=1):
        # Si la línea ya tiene marcador explícito se ignora aquí
        if PAT_AMBOS.search(linea):
            continue

        linea_lower = linea.lower()

        for nombre_lower, u in usuarios_db.items():
            if not nombre_lower:
                continue
            if nombre_lower not in linea_lower:
                continue

            clave = (linea_num, nombre_lower)
            if clave in procesados:
                continue
            procesados.add(clave)

            pos = linea_lower.find(nombre_lower)
            resultados.append({
                'marcador_original': f'[detectado: {u["nombre"]}]',
                'tipo':              'dinamico',
                'nombres':           [u['nombre']],
                'usuarios_validos':  [u],
                'usuarios_invalidos': [],
                'reemplazo_preview': f'«ID_{u["usuario"]}»',
                'linea_num':         linea_num,
                'linea_raw':         linea,
                'match_start':       pos,
                'match_end':         pos + len(nombre_lower),
            })

    return resultados

# ═══════════════════════════════════════════════════════════════════
# LECTURA DE DOCUMENTOS
# ═══════════════════════════════════════════════════════════════════

def _detectar_tipo(ruta: str) -> str:
    ext = os.path.splitext(ruta)[1].lower()
    if ext == '.xlsx': return TIPO_EXCEL
    if ext == '.docx': return TIPO_WORD
    return TIPO_OTRO


def _leer_excel(ruta: str) -> tuple[list, dict]:
    import openpyxl
    wb    = openpyxl.load_workbook(ruta, data_only=True)
    hojas = []
    mapa  = {}

    for nombre_hoja in wb.sheetnames:
        ws    = wb[nombre_hoja]
        filas = []
        col_w = {}

        for idx_f, fila in enumerate(ws.iter_rows(values_only=True), start=1):
            celdas = [str(c) if c is not None else '' for c in fila]
            filas.append((idx_f, celdas))
            for ci, v in enumerate(celdas):
                col_w[ci] = max(col_w.get(ci, 4), min(len(v), 28))
                mapa[(nombre_hoja, idx_f, ci + 1)] = v

        hojas.append({'nombre': nombre_hoja, 'filas': filas, 'col_widths': col_w})

    return hojas, mapa


def _leer_word(ruta: str) -> list:
    from docx import Document
    doc     = Document(ruta)
    parrafos = []

    for i, p in enumerate(doc.paragraphs):
        parrafos.append({'num': i, 'texto': p.text})

    for t_i, tabla in enumerate(doc.tables):
        for r_i, fila in enumerate(tabla.rows):
            texto = ' | '.join(c.text for c in fila.cells)
            parrafos.append({'num': f't{t_i}r{r_i}', 'texto': texto})

    return parrafos


def _texto_plano_excel(hojas_tabla: list) -> str:
    lineas = []
    for hoja in hojas_tabla:
        lineas.append(f'── Hoja: {hoja["nombre"]} ──')
        for idx_f, celdas in hoja['filas']:
            if any(c.strip() for c in celdas):
                lineas.append(f'[F{idx_f}|{hoja["nombre"]}]\t' + '\t'.join(celdas))
    return '\n'.join(lineas)


def _texto_plano_word(parrafos: list) -> str:
    return '\n'.join(f'[P{p["num"]}] {p["texto"]}' for p in parrafos)

# ═══════════════════════════════════════════════════════════════════
# REEMPLAZO EN DOCUMENTOS
# ═══════════════════════════════════════════════════════════════════

def _reemplazar_en_excel(ruta: str, usuario_activo: dict, usuarios_db: dict, hojas_tabla: list) -> str:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb     = openpyxl.load_workbook(ruta)
    cambios = 0

    for ws in wb.worksheets:
        for fila in ws.iter_rows():
            for cell in fila:
                if not cell.value or not isinstance(cell.value, str):
                    continue
                val       = cell.value
                nuevo_val, n = _reemplazar_marcadores_en_texto(
                    val, usuarios_db, ruta, usuario_activo)
                if n > 0:
                    cell.value = nuevo_val
                    cell.fill  = PatternFill('solid', fgColor='D6E4F0')
                    cell.font  = Font(name='Calibri', bold=True, color='1F3864', size=10)
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'),  bottom=Side(style='thin'))
                    cambios += n

    if cambios:
        wb.save(ruta)
        return f'✅ {cambios} marcador(es) reemplazado(s) en {os.path.basename(ruta)}'
    return '⚠ No se encontraron marcadores que reemplazar.'


def _reemplazar_en_word(ruta: str, usuario_activo: dict, usuarios_db: dict) -> str:
    from docx import Document
    from docx.shared import Pt, RGBColor

    doc     = Document(ruta)
    cambios = 0

    def _aplicar_reemplazo_parrafo(p):
        nonlocal cambios
        if not p.text:
            return
        nuevo_texto, n = _reemplazar_marcadores_en_texto(
            p.text, usuarios_db, ruta, usuario_activo)
        if n > 0:
            cambios += n
            for run in p.runs:
                run.text = ''
            if p.runs:
                p.runs[0].text = nuevo_texto
                p.runs[0].font.bold  = True
                p.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
                p.runs[0].font.size  = Pt(10)
            else:
                run = p.add_run(nuevo_texto)
                run.font.bold  = True
                run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
                run.font.size  = Pt(10)

    for p in doc.paragraphs:
        _aplicar_reemplazo_parrafo(p)

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    _aplicar_reemplazo_parrafo(p)

    if cambios:
        doc.save(ruta)
        return f'✅ {cambios} marcador(es) reemplazado(s) en {os.path.basename(ruta)}'
    return '⚠ No se encontraron marcadores que reemplazar.'


def _reemplazar_marcadores_en_texto(
        texto: str,
        usuarios_db: dict,
        ruta_doc: str,
        usuario_activo: dict) -> tuple[str, int]:
    """
    Reemplaza todos los marcadores <<Firma_*>>, <<Nombre_*>> y nombres dinámicos
    en una cadena de texto. Genera y guarda firmas reales en la DB.
    Retorna (texto_nuevo, cantidad_reemplazos).
    """
    resultado  = texto
    reemplazos = 0
    nombre_doc = os.path.basename(ruta_doc)
    _cache_ids: dict = {}

    def _obtener_id_firma(usuario_data: dict) -> str:
        key = usuario_data['usuario']
        if key not in _cache_ids:
            firma = crear_firma(usuario_data, nombre_doc)
            _cache_ids[key] = firma['id_firma']
        return _cache_ids[key]

    # ── Reemplazar <<Firma_...>> ───────────────────────────────────
    def repl_firma(m):
        nonlocal reemplazos
        contenido = m.group(1).strip()
        nombres   = [n.strip() for n in contenido.split(',') if n.strip()]
        ids = []
        for nombre in nombres:
            u = usuarios_db.get(nombre.lower())
            if u:
                ids.append(_obtener_id_firma(u))
            else:
                ids.append(f'«{nombre}:NO_ENCONTRADO»')
        reemplazos += 1
        return ' | '.join(ids)

    resultado = PAT_FIRMA.sub(repl_firma, resultado)

    # ── Reemplazar <<Nombre_...>> ──────────────────────────────────
    def repl_nombre(m):
        nonlocal reemplazos
        contenido    = m.group(1).strip()
        nombres      = [n.strip() for n in contenido.split(',') if n.strip()]
        nombres_reales = []
        for nombre in nombres:
            u = usuarios_db.get(nombre.lower())
            nombres_reales.append(u['nombre'] if u else f'«{nombre}:NO_ENCONTRADO»')
        reemplazos += 1
        return ' / '.join(nombres_reales)

    resultado = PAT_NOMBRE.sub(repl_nombre, resultado)

    # ── Reemplazar nombres detectados dinámicamente ────────────────
    # Solo si NO quedó ningún marcador explícito sin procesar
    if not PAT_AMBOS.search(resultado):
        for nombre_lower, u in usuarios_db.items():
            if not nombre_lower:
                continue
            if nombre_lower not in resultado.lower():
                continue
            id_firma     = _obtener_id_firma(u)
            patron_nombre = re.compile(re.escape(u['nombre']), re.IGNORECASE)
            nuevo, n      = patron_nombre.subn(id_firma, resultado)
            if n > 0:
                resultado   = nuevo
                reemplazos += n

    return resultado, reemplazos

# ═══════════════════════════════════════════════════════════════════
# VENTANA BUSCADOR
# ═══════════════════════════════════════════════════════════════════

class VentanaBuscador:

    C_BG    = '#F0F4F8'
    C_ACC   = '#1F3864'
    C_BTN   = '#2E6DA4'
    C_OK    = '#1E8449'
    C_ERR   = '#C0392B'
    C_MARCA = '#FFD700'   # amarillo  — marcador explícito
    C_REEM  = '#C8F7C5'   # verde     — reemplazado
    C_DIN   = '#D5F5E3'   # verde claro — detección dinámica

    def __init__(self, parent=None, firma_data: dict = None):
        self.firma_data   = firma_data
        self.archivos: list  = []
        self.contenidos: dict = {}
        self.hojas_excel: dict = {}
        self.var_buscar   = tk.StringVar(value='')
        self._meta: dict  = {}

        self.usuarios_db = _obtener_todos_usuarios()

        self.top = tk.Toplevel(parent) if parent else tk.Tk()
        self.top.title('🔎 Buscador de Marcadores de Firma')
        self.top.configure(bg=self.C_BG)
        self.top.resizable(True, True)
        self._centrar(920, 680)
        self._construir_ui()
        self.top.after(200, self._buscar_automatico)

    def _centrar(self, w, h):
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth()  - w) // 2
        y = (self.top.winfo_screenheight() - h) // 2
        self.top.geometry(f'{w}x{h}+{x}+{y}')

    # ──────────────────────────────────────────────────────────────
    # UI
    # ──────────────────────────────────────────────────────────────

    def _construir_ui(self):
        # Encabezado
        hdr = tk.Frame(self.top, bg=self.C_ACC)
        hdr.pack(fill='x')
        tk.Label(hdr, text='🔎 Buscador de Marcadores de Firma',
                 font=('Segoe UI', 12, 'bold'), fg='white', bg=self.C_ACC
                 ).pack(side='left', padx=16, pady=10)
        if self.firma_data:
            tk.Label(hdr, text=f"Usuario activo: {self.firma_data.get('nombre','')}",
                     font=('Segoe UI', 8), fg='#BDD7EE', bg=self.C_ACC
                     ).pack(side='right', padx=16)

        # Leyenda
        ley = tk.Frame(self.top, bg='#E8EEF6', pady=4)
        ley.pack(fill='x')
        tk.Label(ley,
                 text='  Marcadores: <<Firma_Nombre>>  <<Firma_N1,N2>>  <<Nombre_Nombre>>  '
                      '| Verde claro = nombre detectado dinámicamente',
                 font=('Consolas', 8), bg='#E8EEF6', fg='#1F3864'
                 ).pack(side='left')
        tk.Label(ley, text='Doble clic → reemplazar',
                 font=('Segoe UI', 8, 'italic'), bg='#E8EEF6', fg=self.C_ERR
                 ).pack(side='right', padx=12)

        # Layout principal
        main = tk.Frame(self.top, bg=self.C_BG)
        main.pack(fill='both', expand=True, padx=8, pady=6)

        izq = tk.Frame(main, bg=self.C_BG, width=320)
        izq.pack(side='left', fill='y')
        izq.pack_propagate(False)
        self._panel_izquierdo(izq)

        ttk.Separator(main, orient='vertical').pack(side='left', fill='y', padx=4)

        der = tk.Frame(main, bg=self.C_BG)
        der.pack(side='left', fill='both', expand=True)
        self._panel_derecho(der)

    # ── Panel izquierdo ───────────────────────────────────────────

    def _panel_izquierdo(self, parent):
        # Botones
        fb = tk.Frame(parent, bg=self.C_BG)
        fb.pack(fill='x', pady=(0, 4))
        tk.Button(fb, text='📂 Agregar archivo(s)',
                  font=('Segoe UI', 9), bg=self.C_BTN, fg='white',
                  relief='flat', cursor='hand2', command=self._agregar_archivos
                  ).pack(side='left', ipadx=6, ipady=3)
        tk.Button(fb, text='🗑', font=('Segoe UI', 9),
                  bg='#C0392B', fg='white', relief='flat', cursor='hand2',
                  command=self._limpiar
                  ).pack(side='left', padx=(4, 0), ipady=3, ipadx=6)
        tk.Button(fb, text='🔄 Actualizar', font=('Segoe UI', 9),
                  bg='#555', fg='white', relief='flat', cursor='hand2',
                  command=self._buscar_automatico
                  ).pack(side='left', padx=(4, 0), ipady=3, ipadx=4)

        # Lista de archivos
        tk.Label(parent, text='Archivos cargados:',
                 font=('Segoe UI', 9, 'bold'), bg=self.C_BG).pack(anchor='w')

        f_arch = tk.Frame(parent, bg=self.C_BG)
        f_arch.pack(fill='x')
        self.tree_arch = ttk.Treeview(f_arch, columns=('ic', 'nom', 'tipo'),
                                       show='headings', height=5)
        self.tree_arch.heading('ic',   text='')
        self.tree_arch.heading('nom',  text='Nombre')
        self.tree_arch.heading('tipo', text='Tipo')
        self.tree_arch.column('ic',   width=24,  stretch=False)
        self.tree_arch.column('nom',  width=165)
        self.tree_arch.column('tipo', width=88)
        sb = ttk.Scrollbar(f_arch, orient='vertical', command=self.tree_arch.yview)
        self.tree_arch.configure(yscrollcommand=sb.set)
        self.tree_arch.pack(side='left', fill='x', expand=True)
        sb.pack(side='left', fill='y')
        self.tree_arch.bind('<<TreeviewSelect>>', self._al_seleccionar_archivo)

        ttk.Separator(parent, orient='horizontal').pack(fill='x', pady=6)

        # Usuarios en DB
        tk.Label(parent, text='Usuarios en la base de datos:',
                 font=('Segoe UI', 9, 'bold'), bg=self.C_BG).pack(anchor='w')
        f_usr = tk.Frame(parent, bg=self.C_BG)
        f_usr.pack(fill='x', pady=(2, 6))
        self.tree_usr = ttk.Treeview(f_usr, columns=('usr', 'nom'),
                                      show='headings', height=4)
        self.tree_usr.heading('usr', text='Usuario')
        self.tree_usr.heading('nom', text='Nombre completo')
        self.tree_usr.column('usr', width=90)
        self.tree_usr.column('nom', width=170)
        sb2 = ttk.Scrollbar(f_usr, orient='vertical', command=self.tree_usr.yview)
        self.tree_usr.configure(yscrollcommand=sb2.set)
        self.tree_usr.pack(side='left', fill='x', expand=True)
        sb2.pack(side='left', fill='y')
        self._cargar_usuarios_en_tree()

        ttk.Separator(parent, orient='horizontal').pack(fill='x', pady=4)

        # Resultados
        lf = tk.Frame(parent, bg=self.C_BG)
        lf.pack(fill='x')
        tk.Label(lf, text='Marcadores detectados —',
                 font=('Segoe UI', 9, 'bold'), bg=self.C_BG).pack(side='left')
        tk.Label(lf, text=' doble clic = reemplazar',
                 font=('Segoe UI', 8, 'italic'), fg=self.C_ERR, bg=self.C_BG
                 ).pack(side='left')

        f_res = tk.Frame(parent, bg=self.C_BG)
        f_res.pack(fill='both', expand=True, pady=(2, 0))
        self.tree_res = ttk.Treeview(f_res,
                                      columns=('ub', 'marcador', 'reemplazo'),
                                      show='headings', height=8)
        self.tree_res.heading('ub',        text='Dónde')
        self.tree_res.heading('marcador',  text='Marcador / Detectado')
        self.tree_res.heading('reemplazo', text='→ Reemplazar con')
        self.tree_res.column('ub',        width=65,  anchor='center', stretch=False)
        self.tree_res.column('marcador',  width=125)
        self.tree_res.column('reemplazo', width=115)
        sb3 = ttk.Scrollbar(f_res, orient='vertical', command=self.tree_res.yview)
        self.tree_res.configure(yscrollcommand=sb3.set)
        self.tree_res.pack(side='left', fill='both', expand=True)
        sb3.pack(side='left', fill='y')
        self.tree_res.bind('<<TreeviewSelect>>', self._al_seleccionar_resultado)
        self.tree_res.bind('<Double-1>', self._al_doble_clic)

        self.lbl_conteo = tk.Label(parent, text='',
                                    font=('Segoe UI', 9), bg=self.C_BG,
                                    fg=self.C_OK, wraplength=300)
        self.lbl_conteo.pack(anchor='w', pady=(4, 0))

    def _cargar_usuarios_en_tree(self):
        for item in self.tree_usr.get_children():
            self.tree_usr.delete(item)
        for nombre_lower, u in self.usuarios_db.items():
            self.tree_usr.insert('', 'end', values=(u['usuario'], u['nombre']))

    # ── Panel derecho — vista previa ──────────────────────────────

    def _panel_derecho(self, parent):
        self.frm_info = tk.Frame(parent, bg='#DDE6F0', padx=8, pady=4)
        self.frm_info.pack(fill='x')
        self.lbl_info = tk.Label(self.frm_info, text='Selecciona un archivo',
                                  font=('Segoe UI', 10, 'bold'),
                                  bg='#DDE6F0', fg=self.C_ACC)
        self.lbl_info.pack(side='left')
        self.lbl_ruta = tk.Label(self.frm_info, text='',
                                  font=('Segoe UI', 8), bg='#DDE6F0', fg='#666')
        self.lbl_ruta.pack(side='left', padx=8)

        self.nb_prev = ttk.Notebook(parent)
        self.nb_prev.pack(fill='both', expand=True, pady=(4, 0))

        # Pestaña texto
        tab_txt = tk.Frame(self.nb_prev, bg='#FAFCFF')
        self.nb_prev.add(tab_txt, text=' 📄 Texto ')
        self.txt_prev = tk.Text(tab_txt, font=('Courier New', 9),
                                 bg='#FAFCFF', relief='flat',
                                 wrap='none', state='disabled',
                                 padx=8, pady=6)
        sb_y = ttk.Scrollbar(tab_txt, orient='vertical',   command=self.txt_prev.yview)
        sb_x = ttk.Scrollbar(tab_txt, orient='horizontal', command=self.txt_prev.xview)
        self.txt_prev.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side='right',  fill='y')
        sb_x.pack(side='bottom', fill='x')
        self.txt_prev.pack(fill='both', expand=True)

        # Tags de resaltado
        self.txt_prev.tag_configure('marca_firma',
            background='#FFD700', foreground='#000',
            font=('Courier New', 9, 'bold'))
        self.txt_prev.tag_configure('marca_nombre',
            background='#AED6F1', foreground='#000',
            font=('Courier New', 9, 'bold'))
        self.txt_prev.tag_configure('marca_dinamica',
            background='#D5F5E3', foreground='#000',
            font=('Courier New', 9, 'bold'))
        self.txt_prev.tag_configure('linea_activa', background='#FFFDE7')
        self.txt_prev.tag_configure('invalido',
            background='#FADBD8', foreground='#922B21')

        # Pestaña tabla Excel
        tab_xls = tk.Frame(self.nb_prev, bg='#FAFCFF')
        self.nb_prev.add(tab_xls, text=' 📊 Tabla Excel ')
        self.nb_hojas = ttk.Notebook(tab_xls)
        self.nb_hojas.pack(fill='both', expand=True)
        self._tabs_hojas: list = []

        # Leyenda de colores
        ley_prev = tk.Frame(parent, bg=self.C_BG)
        ley_prev.pack(fill='x', padx=4, pady=(2, 0))
        for color, texto in [('#FFD700', '<<Firma_*>>'),
                              ('#AED6F1', '<<Nombre_*>>'),
                              ('#D5F5E3', 'Nombre dinámico')]:
            tk.Label(ley_prev, text=f'  {texto}  ',
                     bg=color, fg='#000',
                     font=('Segoe UI', 8), relief='flat', padx=4
                     ).pack(side='left', padx=2)

        self.lbl_estado = tk.Label(parent, text='',
                                    font=('Segoe UI', 9, 'bold'),
                                    bg=self.C_BG, fg=self.C_OK, wraplength=560)
        self.lbl_estado.pack(pady=(4, 2))

    # ──────────────────────────────────────────────────────────────
    # ACCIONES — ARCHIVOS
    # ──────────────────────────────────────────────────────────────

    def _agregar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title='Seleccionar documento(s)',
            filetypes=[('Excel y Word', '*.xlsx *.docx'),
                       ('Excel', '*.xlsx'), ('Word', '*.docx'),
                       ('Todos', '*.*')])
        for ruta in rutas:
            if ruta not in self.archivos:
                self.archivos.append(ruta)
                tipo  = _detectar_tipo(ruta)
                icono = ICONOS_TIPO.get(tipo, '📁')
                self.tree_arch.insert('', 'end', iid=ruta,
                                       values=(icono, os.path.basename(ruta), tipo))
        if rutas:
            self._buscar_automatico()

    def _limpiar(self):
        self.archivos.clear()
        self.contenidos.clear()
        self.hojas_excel.clear()
        self._meta.clear()
        for t in self.tree_arch.get_children(): self.tree_arch.delete(t)
        for t in self.tree_res.get_children():  self.tree_res.delete(t)
        self._limpiar_preview()
        self.lbl_conteo.config(text='')
        self.lbl_estado.config(text='')

    def _limpiar_preview(self):
        self.txt_prev.config(state='normal')
        self.txt_prev.delete('1.0', tk.END)
        self.txt_prev.config(state='disabled')

    def _al_seleccionar_archivo(self, event=None):
        sel = self.tree_arch.selection()
        if not sel: return
        ruta  = sel[0]
        tipo  = _detectar_tipo(ruta)
        icono = ICONOS_TIPO.get(tipo, '📁')
        color = COLORES_TIPO.get(tipo, '#333')
        self.lbl_info.config(text=f'{icono} {os.path.basename(ruta)} — {tipo}', fg=color)
        self.lbl_ruta.config(text=ruta)
        self._cargar_contenido(ruta)
        self._cargar_preview_texto(ruta)
        if tipo == TIPO_EXCEL:
            self._cargar_preview_tabla(ruta)
            self.nb_prev.select(1)
        else:
            self.nb_prev.select(0)

    def _cargar_contenido(self, ruta: str):
        if ruta in self.contenidos: return
        tipo = _detectar_tipo(ruta)
        try:
            if tipo == TIPO_EXCEL:
                hojas, _ = _leer_excel(ruta)
                self.hojas_excel[ruta] = hojas
                self.contenidos[ruta]  = _texto_plano_excel(hojas)
            elif tipo == TIPO_WORD:
                parrs = _leer_word(ruta)
                self.contenidos[ruta] = _texto_plano_word(parrs)
            else:
                self.contenidos[ruta] = '[Tipo no soportado]'
        except Exception as e:
            self.contenidos[ruta] = f'[Error al leer: {e}]'

    def _cargar_preview_texto(self, ruta: str):
        contenido = self.contenidos.get(ruta, '')
        self.txt_prev.config(state='normal')
        self.txt_prev.delete('1.0', tk.END)
        self.txt_prev.insert('1.0', contenido)

        # Resaltar marcadores explícitos
        self._resaltar_en_texto(PAT_FIRMA.pattern,  'marca_firma')
        self._resaltar_en_texto(PAT_NOMBRE.pattern, 'marca_nombre')

        # Resaltar nombres detectados dinámicamente (solo en líneas sin marcador)
        for linea_num, linea in enumerate(contenido.splitlines(), start=1):
            if PAT_AMBOS.search(linea):
                continue
            linea_lower = linea.lower()
            for nombre_lower, u in self.usuarios_db.items():
                if not nombre_lower: continue
                if nombre_lower not in linea_lower: continue
                patron = re.compile(re.escape(u['nombre']), re.IGNORECASE)
                for m in patron.finditer(linea):
                    start_idx = f'{linea_num}.{m.start()}'
                    end_idx   = f'{linea_num}.{m.end()}'
                    self.txt_prev.tag_add('marca_dinamica', start_idx, end_idx)

        self.txt_prev.config(state='disabled')

    def _resaltar_en_texto(self, patron: str, tag: str):
        start = '1.0'
        while True:
            pos = self.txt_prev.search(patron, start,
                                        stopindex=tk.END, nocase=True, regexp=True)
            if not pos: break
            line, col = map(int, pos.split('.'))
            linea_txt = self.txt_prev.get(f'{line}.0', f'{line}.end')
            m = re.search(patron, linea_txt[int(col):], re.IGNORECASE)
            if not m: break
            end   = f'{pos}+{len(m.group(0))}c'
            self.txt_prev.tag_add(tag, pos, end)
            start = end

    def _cargar_preview_tabla(self, ruta: str):
        for tab in self.nb_hojas.tabs():
            self.nb_hojas.forget(tab)
        self._tabs_hojas.clear()

        hojas = self.hojas_excel.get(ruta, [])
        if not hojas: return

        for hoja in hojas:
            frm    = tk.Frame(self.nb_hojas, bg='#FAFCFF')
            self.nb_hojas.add(frm, text=f' {hoja["nombre"]} ')
            canvas = tk.Canvas(frm, bg='#FAFCFF', highlightthickness=0)
            sb_vy  = ttk.Scrollbar(frm, orient='vertical',   command=canvas.yview)
            sb_hx  = ttk.Scrollbar(frm, orient='horizontal', command=canvas.xview)
            canvas.configure(yscrollcommand=sb_vy.set, xscrollcommand=sb_hx.set)
            sb_vy.pack(side='right',  fill='y')
            sb_hx.pack(side='bottom', fill='x')
            canvas.pack(fill='both', expand=True)

            interior = tk.Frame(canvas, bg='#FAFCFF')
            canvas.create_window((0, 0), window=interior, anchor='nw')
            interior.bind('<Configure>',
                          lambda e, cv=canvas: cv.configure(scrollregion=cv.bbox('all')))

            col_w  = hoja['col_widths']
            n_cols = max((len(f[1]) for f in hoja['filas']), default=1)
            px_w   = [max(col_w.get(ci, 4) * 7, 48) for ci in range(n_cols)]

            from openpyxl.utils import get_column_letter
            tk.Label(interior, text='', width=4, bg='#C5D5E8',
                     relief='groove', bd=1, font=('Calibri', 8, 'bold')
                     ).grid(row=0, column=0, sticky='nsew', ipadx=2, ipady=2)

            for ci in range(n_cols):
                tk.Label(interior, text=get_column_letter(ci + 1),
                         width=max(px_w[ci] // 7, 5),
                         bg='#C5D5E8', fg='#1F3864',
                         font=('Calibri', 8, 'bold'), relief='groove', bd=1
                         ).grid(row=0, column=ci+1, sticky='nsew', ipadx=2, ipady=2)

            for gr, (n_fila, celdas) in enumerate(hoja['filas'], start=1):
                tk.Label(interior, text=str(n_fila), width=4,
                         bg='#DDE6F0', relief='groove', bd=1,
                         font=('Calibri', 8)
                         ).grid(row=gr, column=0, sticky='nsew', ipadx=2, ipady=1)

                for ci, val in enumerate(celdas):
                    tiene_firma  = bool(PAT_FIRMA.search(val))
                    tiene_nombre = bool(PAT_NOMBRE.search(val))
                    tiene_din    = (not tiene_firma and not tiene_nombre and
                                    any(nl in val.lower()
                                        for nl in self.usuarios_db if nl))

                    if tiene_firma:
                        bg, fg, font = '#FFD700', '#000', ('Calibri', 8, 'bold')
                    elif tiene_nombre:
                        bg, fg, font = '#AED6F1', '#000', ('Calibri', 8, 'bold')
                    elif tiene_din:
                        bg, fg, font = '#D5F5E3', '#000', ('Calibri', 8, 'bold')
                    elif val:
                        bg, fg, font = '#FFFFFF', '#222', ('Calibri', 8)
                    else:
                        bg, fg, font = '#FAFAFA', '#666', ('Calibri', 8)

                    tk.Label(interior, text=val[:24],
                             width=max(px_w[ci] // 7, 5),
                             bg=bg, fg=fg, font=font,
                             anchor='w', relief='groove', bd=1
                             ).grid(row=gr, column=ci+1, sticky='nsew', ipadx=3, ipady=1)

            self._tabs_hojas.append(frm)

    # ──────────────────────────────────────────────────────────────
    # BÚSQUEDA AUTOMÁTICA
    # ──────────────────────────────────────────────────────────────

    def _buscar_automatico(self):
        for item in self.tree_res.get_children():
            self.tree_res.delete(item)
        self._meta.clear()

        if not self.archivos:
            self.lbl_conteo.config(
                text='Agrega un archivo con marcadores <<Firma_...>>', fg='#888')
            return

        total     = 0
        total_din = 0

        for ruta in self.archivos:
            self._cargar_contenido(ruta)
            contenido = self.contenidos.get(ruta, '')
            tipo      = _detectar_tipo(ruta)

            # Marcadores explícitos
            marcadores = _analizar_marcadores(contenido, self.usuarios_db)
            # Detección dinámica
            dinamicos  = _detectar_nombres_dinamicos(contenido, self.usuarios_db)
            total_din += len(dinamicos)
            marcadores += dinamicos

            for mc in marcadores:
                total += 1
                raw = mc['linea_raw']
                m2  = re.match(r'^\[F(\d+)\|(.+?)\]', raw)
                m3  = re.match(r'^\[P(\d+)\]',         raw)
                if m2:
                    ub = f"F{m2.group(1)} / {m2.group(2)[:10]}"
                elif m3:
                    ub = f"Párr.{m3.group(1)}"
                else:
                    ub = f"L{mc['linea_num']}"

                es_dinamico = mc['tipo'] == 'dinamico'
                tag_color   = ('dinamico' if es_dinamico
                               else ('invalido' if mc['usuarios_invalidos'] else 'valido'))

                iid = self.tree_res.insert('', 'end',
                    values=(ub,
                            mc['marcador_original'][:38],
                            mc['reemplazo_preview'][:38]),
                    tags=(tag_color,))
                self._meta[iid] = {'ruta': ruta, 'tipo': tipo, 'marcador': mc}

        self.tree_res.tag_configure('valido',   background='#F0FFF0')
        self.tree_res.tag_configure('invalido',  background='#FFF0F0')
        self.tree_res.tag_configure('dinamico',  background='#EAFAF1')

        if total == 0:
            self.lbl_conteo.config(
                text='No se encontraron marcadores ni nombres conocidos.',
                fg='#856404')
        else:
            n_inv = sum(1 for v in self._meta.values()
                        if v['marcador']['usuarios_invalidos'])
            msg = f'✅ {total} elemento(s) detectado(s)'
            if total_din:
                msg += f'  ({total_din} dinámico(s))'
            if n_inv:
                msg += f'  ⚠ {n_inv} usuario(s) no encontrado(s)'
            self.lbl_conteo.config(
                text=msg, fg=self.C_OK if not n_inv else '#856404')

        if self.archivos and not self.tree_arch.selection():
            self.tree_arch.selection_set(self.archivos[0])
            self._al_seleccionar_archivo()

    # ──────────────────────────────────────────────────────────────
    # NAVEGACIÓN
    # ──────────────────────────────────────────────────────────────

    def _al_seleccionar_resultado(self, event=None):
        sel = self.tree_res.selection()
        if not sel: return
        meta = self._meta.get(sel[0])
        if not meta: return
        ruta = meta['ruta']
        mc   = meta['marcador']

        if not self.tree_arch.selection() or self.tree_arch.selection()[0] != ruta:
            self.tree_arch.selection_set(ruta)
            self._al_seleccionar_archivo()

        n = mc['linea_num']
        self.txt_prev.config(state='normal')
        self.txt_prev.tag_remove('linea_activa', '1.0', tk.END)
        self.txt_prev.tag_add('linea_activa', f'{n}.0', f'{n}.end')
        self.txt_prev.see(f'{n}.0')
        self.txt_prev.config(state='disabled')

    # ──────────────────────────────────────────────────────────────
    # DOBLE CLIC — REEMPLAZAR
    # ──────────────────────────────────────────────────────────────

    def _al_doble_clic(self, event=None):
        sel = self.tree_res.selection()
        if not sel: return
        iid  = sel[0]
        meta = self._meta.get(iid)
        if not meta: return
        ruta = meta['ruta']
        mc   = meta['marcador']
        tipo = meta['tipo']

        if mc['usuarios_invalidos']:
            nombres_inv = ', '.join(mc['usuarios_invalidos'])
            resp = messagebox.askyesno(
                'Usuarios no encontrados',
                f'El marcador contiene nombre(s) que no existen en la DB:\n\n'
                f'  {nombres_inv}\n\n'
                f'¿Continuar de todas formas? (se marcará como NO_ENCONTRADO)',
                parent=self.top)
            if not resp: return

        es_din = mc['tipo'] == 'dinamico'
        origen = 'Nombre detectado dinámicamente' if es_din else 'Marcador explícito'
        msg = (
            f'¿Reemplazar TODOS los marcadores y nombres detectados en este archivo?\n\n'
            f'  Archivo : {os.path.basename(ruta)}\n'
            f'  Tipo    : {origen}\n'
            f'  Elemento: {mc["marcador_original"]}\n'
            f'  → Con   : {mc["reemplazo_preview"]}\n\n'
            f'(Se procesarán todos los elementos del archivo, no solo este)'
        )
        if not messagebox.askyesno('Confirmar reemplazo', msg, parent=self.top):
            return

        usuario_activo = {}
        if self.firma_data:
            usuario_activo = {
                'usuario': self.firma_data.get('usuario', ''),
                'nombre':  self.firma_data.get('nombre',  ''),
            }

        try:
            if tipo == TIPO_EXCEL:
                hojas    = self.hojas_excel.get(ruta, [])
                resultado = _reemplazar_en_excel(
                    ruta, usuario_activo, self.usuarios_db, hojas)
            elif tipo == TIPO_WORD:
                resultado = _reemplazar_en_word(
                    ruta, usuario_activo, self.usuarios_db)
            else:
                resultado = '❌ Tipo de archivo no soportado.'
        except Exception as e:
            resultado = f'❌ Error: {e}'

        exito = resultado.startswith('✅')
        self.lbl_estado.config(text=resultado,
                                fg=self.C_OK if exito else self.C_ERR)
        if exito:
            if ruta in self.contenidos:  del self.contenidos[ruta]
            if ruta in self.hojas_excel: del self.hojas_excel[ruta]
            self._cargar_contenido(ruta)
            self._al_seleccionar_archivo()
            self._buscar_automatico()
            messagebox.showinfo('Reemplazo completado', resultado, parent=self.top)
        else:
            messagebox.showerror('Error', resultado, parent=self.top)

    # ──────────────────────────────────────────────────────────────

    def iniciar(self):
        self.top.mainloop()

# ═══════════════════════════════════════════════════════════════════
# FUNCIÓN DE CONVENIENCIA
# ═══════════════════════════════════════════════════════════════════

def abrir_buscador(parent=None, firma_data: dict = None):
    """Abre la ventana buscadora. firma_data = dict con usuario y nombre."""
    return VentanaBuscador(parent=parent, firma_data=firma_data)


if __name__ == '__main__':
    from modules.db import inicializar_db
    inicializar_db()
    app = VentanaBuscador()
    app.iniciar()