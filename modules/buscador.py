"""
buscador.py — Ventana de búsqueda de firma en documentos
Permite buscar la palabra 'firma', diferenciar tipos de archivo,
mostrar una vista previa con la ubicación de la firma,
e INSERTAR la firma con doble clic sobre el resultado.
"""

import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ─────────────────────────────────────────────
# CONSTANTES DE TIPO DE ARCHIVO
# ─────────────────────────────────────────────

TIPO_EXCEL = "Excel (.xlsx)"
TIPO_WORD  = "Word (.docx)"
TIPO_TEXTO = "Texto (.txt)"
TIPO_OTRO  = "Otro"

ICONOS_TIPO = {
    TIPO_EXCEL: "📊",
    TIPO_WORD:  "📝",
    TIPO_TEXTO: "📄",
    TIPO_OTRO:  "📁",
}

COLORES_TIPO = {
    TIPO_EXCEL: "#1E6B3C",
    TIPO_WORD:  "#1E3A6B",
    TIPO_TEXTO: "#555555",
    TIPO_OTRO:  "#888888",
}


# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────

def _detectar_tipo(ruta: str) -> str:
    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".xlsx": return TIPO_EXCEL
    if ext == ".docx": return TIPO_WORD
    if ext == ".txt":  return TIPO_TEXTO
    return TIPO_OTRO


def _leer_contenido_texto(ruta: str) -> str:
    tipo = _detectar_tipo(ruta)
    try:
        if tipo == TIPO_TEXTO:
            with open(ruta, "r", encoding="utf-8", errors="replace") as f:
                return f.read()

        elif tipo == TIPO_EXCEL:
            try:
                import openpyxl
                wb = openpyxl.load_workbook(ruta, data_only=True)
                lineas = []
                for nombre_hoja in wb.sheetnames:
                    ws = wb[nombre_hoja]
                    lineas.append(f"── Hoja: {nombre_hoja} ──")
                    for idx_fila, fila in enumerate(ws.iter_rows(values_only=True), start=1):
                        celdas = [str(c) if c is not None else "" for c in fila]
                        if any(c.strip() for c in celdas):
                            # Prefijo con hoja + fila exactos para mapeo directo sin re-contar
                            lineas.append(f"[F{idx_fila}|{nombre_hoja}]\t" + "\t".join(celdas))
                return "\n".join(lineas)
            except ImportError:
                return "[openpyxl no instalado — no se puede leer .xlsx]"

        elif tipo == TIPO_WORD:
            try:
                from docx import Document
                doc = Document(ruta)
                parrafos = []
                for idx, p in enumerate(doc.paragraphs, start=1):
                    parrafos.append(f"[P{idx}] {p.text}")
                for tabla in doc.tables:
                    for fila in tabla.rows:
                        parrafos.append("\t".join(c.text for c in fila.cells))
                return "\n".join(parrafos)
            except ImportError:
                return "[python-docx no instalado — no se puede leer .docx]"

        else:
            return "[Vista previa no disponible para este tipo de archivo]"

    except Exception as e:
        return f"[Error al leer el archivo: {e}]"


def _buscar_firma_en_texto(contenido: str, palabra: str = "firma") -> list:
    palabra_lower = palabra.lower()
    resultados = []
    for n_linea, linea in enumerate(contenido.splitlines(), start=1):
        linea_lower = linea.lower()
        col = 0
        while True:
            idx = linea_lower.find(palabra_lower, col)
            if idx == -1:
                break
            resultados.append({
                "linea":    n_linea,
                "columna":  idx + 1,
                "contexto": linea.strip(),
                "linea_raw": linea,   # conserva prefijo [F<n>|<hoja>] o [P<n>]
            })
            col = idx + 1
    return resultados


# ─────────────────────────────────────────────
# INSERCIÓN EN EXCEL
# ─────────────────────────────────────────────

def _liberar_merges_zona(ws, fila_inicio: int, col_inicio: int, filas: int = 3, cols: int = 2):
    """
    Elimina cualquier rango combinado que colisione con la zona donde
    se va a insertar la firma, para evitar el error MergedCell read-only.
    """
    from openpyxl.utils import range_boundaries
    merges_a_eliminar = []
    for rango in list(ws.merged_cells.ranges):
        min_col, min_row, max_col, max_row = range_boundaries(str(rango))
        # Colisión si los rangos se solapan
        if (min_row <= fila_inicio + filas - 1 and max_row >= fila_inicio and
                min_col <= col_inicio + cols - 1 and max_col >= col_inicio):
            merges_a_eliminar.append(str(rango))
    for rango in merges_a_eliminar:
        ws.unmerge_cells(rango)


def _insertar_firma_excel(ruta: str, n_linea_contenido: int, firma_data: dict,
                           linea_raw: str = "") -> str:
    """
    Inserta nombre + ID de firma en el archivo Excel.

    Parsea el prefijo [F<fila>|<hoja>] directamente desde linea_raw para
    un mapeo exacto sin re-contar. Libera merges que colisionen antes de escribir.
    """
    try:
        import openpyxl, re
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        return "openpyxl no instalado."

    # ── 1. Parsear fila y hoja del prefijo [F<n>|<hoja>] ──────────────────
    nombre_hoja = None
    fila_excel  = None

    m = re.match(r'^\[F(\d+)\|(.+?)\]', linea_raw.strip())
    if m:
        fila_excel  = int(m.group(1))
        nombre_hoja = m.group(2)
    else:
        # Fallback: reconstruir mapa contando líneas (método anterior)
        try:
            wb_tmp = openpyxl.load_workbook(ruta, data_only=True)
            n = 0
            for hoja in wb_tmp.sheetnames:
                ws_tmp = wb_tmp[hoja]
                n += 1
                for idx_f, fila in enumerate(ws_tmp.iter_rows(values_only=True), start=1):
                    celdas = [str(c) if c is not None else "" for c in fila]
                    if any(c.strip() for c in celdas):
                        n += 1
                        if n == n_linea_contenido:
                            fila_excel  = idx_f
                            nombre_hoja = hoja
                            break
                if fila_excel:
                    break
        except Exception as e:
            return f"Error al leer el archivo: {e}"

    if not fila_excel or not nombre_hoja:
        return f"No se pudo determinar la celda destino (línea {n_linea_contenido})."

    # ── 2. Abrir libro para escritura ──────────────────────────────────────
    try:
        wb = openpyxl.load_workbook(ruta)
        if nombre_hoja not in wb.sheetnames:
            nombre_hoja = wb.active.title
        ws = wb[nombre_hoja]

        # Primera columna libre en esa fila (ignorar celdas combinadas al buscar)
        col_destino = 1
        max_col = ws.max_column or 1
        while col_destino <= max_col + 2:
            cell = ws.cell(row=fila_excel, column=col_destino)
            # Una MergedCell también tiene .value = None; la saltamos
            val = cell.value if hasattr(cell, "value") else None
            if val is None or str(val).strip() == "":
                break
            col_destino += 1

        # ── 3. Liberar merges que colisionen con las 3 filas × 2 cols ──────
        _liberar_merges_zona(ws, fila_excel, col_destino, filas=3, cols=2)

        # ── 4. Estilos ──────────────────────────────────────────────────────
        fuente_header = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        fuente_label  = Font(name="Calibri", bold=True, size=10, color="1F3864")
        fuente_valor  = Font(name="Calibri", size=10)
        relleno       = PatternFill("solid", fgColor="D6E4F0")
        alineacion    = Alignment(horizontal="left", vertical="center")
        borde = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"),  bottom=Side(style="thin")
        )

        # ── 5. Encabezado con merge ─────────────────────────────────────────
        ws.merge_cells(
            start_row=fila_excel, start_column=col_destino,
            end_row=fila_excel,   end_column=col_destino + 1
        )
        c_enc = ws.cell(row=fila_excel, column=col_destino, value="🔏 FIRMA DIGITAL")
        c_enc.font      = fuente_header
        c_enc.fill      = PatternFill("solid", fgColor="1F3864")
        c_enc.alignment = Alignment(horizontal="center", vertical="center")
        c_enc.border    = borde
        ws.row_dimensions[fila_excel].height = 18

        # ── 6. Filas de datos: Firmado por | valor  /  ID Firma | valor ────
        datos = [
            ("Firmado por", firma_data.get("nombre",   "")),
            ("ID Firma",    firma_data.get("id_firma", "")),
        ]
        for offset, (etiqueta, valor) in enumerate(datos, start=1):
            f = fila_excel + offset
            # Liberar posibles merges en filas de datos también
            _liberar_merges_zona(ws, f, col_destino, filas=1, cols=2)

            c_lbl = ws.cell(row=f, column=col_destino,     value=etiqueta)
            c_val = ws.cell(row=f, column=col_destino + 1, value=valor)
            for c in (c_lbl, c_val):
                c.fill      = relleno
                c.alignment = alineacion
                c.border    = borde
            c_lbl.font = fuente_label
            c_val.font = fuente_valor
            ws.row_dimensions[f].height = 15

        # ── 7. Anchos mínimos ───────────────────────────────────────────────
        col_a = get_column_letter(col_destino)
        col_b = get_column_letter(col_destino + 1)
        if ws.column_dimensions[col_a].width < 14:
            ws.column_dimensions[col_a].width = 14
        if ws.column_dimensions[col_b].width < 26:
            ws.column_dimensions[col_b].width = 26

        wb.save(ruta)
        celda_ref = f"{get_column_letter(col_destino)}{fila_excel}"
        return f"✅ Firma insertada en {nombre_hoja}!{celda_ref}"

    except Exception as e:
        return f"Error al insertar firma: {e}"


# ─────────────────────────────────────────────
# INSERCIÓN EN WORD
# ─────────────────────────────────────────────

def _insertar_firma_word(ruta: str, n_linea_contenido: int, firma_data: dict) -> str:
    """
    Inserta la firma después del párrafo indicado por n_linea_contenido.
    El contenido de Word tiene prefijos [P<número>].
    """
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        import re
    except ImportError:
        return "python-docx no instalado."

    try:
        doc = Document(ruta)
        parrafos = [p for p in doc.paragraphs]

        # Reconstruir mapa: n_linea_contenido → índice de párrafo
        # El contenido tiene líneas "[P1] texto", "[P2] texto", etc.
        # n_linea_contenido es la posición en el texto extraído
        n = 0
        idx_parrafo = None
        for idx, p in enumerate(parrafos):
            n += 1
            if n == n_linea_contenido:
                idx_parrafo = idx
                break

        if idx_parrafo is None:
            # Fallback: insertar al final
            idx_parrafo = len(parrafos) - 1

        # Insertar firma después del párrafo encontrado
        # python-docx no tiene insert_paragraph_after nativo, usamos XML
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        import copy

        parrafo_ref = parrafos[idx_parrafo]._element

        def _nuevo_parrafo_firma(texto: str, negrita: bool = False, color_hex: str = "1F3864"):
            p_new = OxmlElement('w:p')
            r_new = OxmlElement('w:r')
            rPr   = OxmlElement('w:rPr')

            # Fuente
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), 'Calibri')
            rPr.append(rFonts)

            # Tamaño
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '20')
            rPr.append(sz)

            # Negrita
            if negrita:
                b = OxmlElement('w:b')
                rPr.append(b)

            # Color
            color_el = OxmlElement('w:color')
            color_el.set(qn('w:val'), color_hex)
            rPr.append(color_el)

            r_new.append(rPr)
            t_new = OxmlElement('w:t')
            t_new.text = texto
            t_new.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r_new.append(t_new)
            p_new.append(r_new)
            return p_new

        # Líneas a insertar (en orden inverso para que queden en orden correcto)
        lineas_firma = [
            ("🔏 FIRMA DIGITAL",                       True,  "1F3864"),
            (f"Firmado por:  {firma_data.get('nombre','')}", False, "000000"),
            (f"ID Firma:     {firma_data.get('id_firma','')}", False, "2E6DA4"),
        ]

        for texto, negrita, color in reversed(lineas_firma):
            p_nuevo = _nuevo_parrafo_firma(texto, negrita, color)
            parrafo_ref.addnext(p_nuevo)

        doc.save(ruta)
        return f"✅ Firma insertada después del párrafo {idx_parrafo + 1}"

    except Exception as e:
        return f"Error al insertar firma en Word: {e}"


# ─────────────────────────────────────────────
# VENTANA BUSCADOR
# ─────────────────────────────────────────────


# ─────────────────────────────────────────────
# LECTURA ESTRUCTURADA PARA VISTA DE TABLA
# ─────────────────────────────────────────────

def _leer_excel_como_tabla(ruta: str) -> list:
    """
    Retorna lista de hojas, cada una con:
      { "nombre": str, "filas": [[str,...]], "col_widths": [int,...] }
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ruta, data_only=True)
        hojas = []
        for nombre_hoja in wb.sheetnames:
            ws = wb[nombre_hoja]
            filas = []
            col_maxlen = {}
            for idx_fila, fila in enumerate(ws.iter_rows(values_only=True), start=1):
                celdas = [str(c) if c is not None else "" for c in fila]
                filas.append((idx_fila, celdas))
                for ci, v in enumerate(celdas):
                    col_maxlen[ci] = max(col_maxlen.get(ci, 4), min(len(v), 30))
            hojas.append({
                "nombre":     nombre_hoja,
                "filas":      filas,           # [(n_fila, [celdas...]), ...]
                "col_widths": col_maxlen,
            })
        return hojas
    except Exception:
        return []

class VentanaBuscador:
    COLOR_BG       = "#F0F4F8"
    COLOR_ACCENT   = "#1F3864"
    COLOR_BTN      = "#2E6DA4"
    COLOR_OK       = "#1E8449"
    COLOR_ERR      = "#C0392B"
    COLOR_MARCA    = "#FFD700"
    COLOR_MARCA_FG = "#000000"

    def __init__(self, parent=None, firma_data: dict = None):
        self.firma_data = firma_data   # dict con nombre + id_firma

        if parent:
            self.top = tk.Toplevel(parent)
        else:
            self.top = tk.Tk()

        self.top.title("🔎 Buscador de Firma en Documentos")
        self.top.configure(bg=self.COLOR_BG)
        self.top.resizable(True, True)
        self._centrar_ventana(860, 640)

        self.archivos: list  = []
        self.contenidos: dict = {}
        self.var_buscar = tk.StringVar(value="firma")

        # Mapa: iid del resultado → (ruta_archivo, n_linea_en_contenido)
        self._resultado_meta: dict = {}

        self._construir_ui()

    # ──────────────────────────────────────────
    def _centrar_ventana(self, ancho, alto):
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth()  // 2) - (ancho // 2)
        y = (self.top.winfo_screenheight() // 2) - (alto  // 2)
        self.top.geometry(f"{ancho}x{alto}+{x}+{y}")

    # ──────────────────────────────────────────
    # UI
    # ──────────────────────────────────────────

    def _construir_ui(self):
        frame_header = tk.Frame(self.top, bg=self.COLOR_ACCENT, height=50)
        frame_header.pack(fill="x")
        tk.Label(
            frame_header,
            text="🔎  Buscador de Firma en Documentos",
            font=("Segoe UI", 12, "bold"),
            fg="white", bg=self.COLOR_ACCENT
        ).pack(side="left", padx=20, pady=12)

        # Indicador de firma cargada
        if self.firma_data:
            nombre = self.firma_data.get("nombre", "")
            id_f   = self.firma_data.get("id_firma", "")
            tk.Label(
                frame_header,
                text=f"Firma activa: {nombre}  |  {id_f}",
                font=("Segoe UI", 8),
                fg="#A8D4FF", bg=self.COLOR_ACCENT
            ).pack(side="right", padx=16, pady=12)

        panel_izq = tk.Frame(self.top, bg=self.COLOR_BG, width=310)
        panel_izq.pack(side="left", fill="y", padx=(10, 5), pady=10)
        panel_izq.pack_propagate(False)
        self._construir_panel_archivos(panel_izq)

        ttk.Separator(self.top, orient="vertical").pack(side="left", fill="y", pady=10)

        panel_der = tk.Frame(self.top, bg=self.COLOR_BG)
        panel_der.pack(side="left", fill="both", expand=True, padx=(5, 10), pady=10)
        self._construir_panel_preview(panel_der)

    # ──────────────────────────────────────────
    # PANEL IZQUIERDO
    # ──────────────────────────────────────────

    def _construir_panel_archivos(self, parent):
        frame_btns = tk.Frame(parent, bg=self.COLOR_BG)
        frame_btns.pack(fill="x", pady=(0, 6))

        tk.Button(
            frame_btns, text="📂 Agregar archivo(s)",
            font=("Segoe UI", 9), bg=self.COLOR_BTN, fg="white",
            relief="flat", cursor="hand2",
            command=self._agregar_archivos
        ).pack(side="left", ipadx=6, ipady=4)

        tk.Button(
            frame_btns, text="🗑 Limpiar",
            font=("Segoe UI", 9), bg="#C0392B", fg="white",
            relief="flat", cursor="hand2",
            command=self._limpiar_archivos
        ).pack(side="left", padx=(6, 0), ipadx=6, ipady=4)

        tk.Label(parent, text="Archivos cargados:",
                 font=("Segoe UI", 9, "bold"), bg=self.COLOR_BG).pack(anchor="w")

        frame_lista = tk.Frame(parent, bg=self.COLOR_BG)
        frame_lista.pack(fill="both", expand=False)

        cols = ("icono", "nombre", "tipo")
        self.tree_archivos = ttk.Treeview(
            frame_lista, columns=cols, show="headings", height=6, selectmode="browse"
        )
        self.tree_archivos.heading("icono",  text="")
        self.tree_archivos.heading("nombre", text="Nombre")
        self.tree_archivos.heading("tipo",   text="Tipo")
        self.tree_archivos.column("icono",  width=28, anchor="center", stretch=False)
        self.tree_archivos.column("nombre", width=155)
        self.tree_archivos.column("tipo",   width=90)

        sb_arch = ttk.Scrollbar(frame_lista, orient="vertical", command=self.tree_archivos.yview)
        self.tree_archivos.configure(yscrollcommand=sb_arch.set)
        self.tree_archivos.pack(side="left", fill="both", expand=True)
        sb_arch.pack(side="left", fill="y")
        self.tree_archivos.bind("<<TreeviewSelect>>", self._al_seleccionar_archivo)

        # Búsqueda
        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=8)
        tk.Label(parent, text="Palabra a buscar:",
                 font=("Segoe UI", 9, "bold"), bg=self.COLOR_BG).pack(anchor="w")

        frame_busq = tk.Frame(parent, bg=self.COLOR_BG)
        frame_busq.pack(fill="x", pady=(4, 0))

        self.entry_buscar = ttk.Entry(frame_busq, textvariable=self.var_buscar,
                                      font=("Segoe UI", 10))
        self.entry_buscar.pack(side="left", fill="x", expand=True, ipady=4)
        self.entry_buscar.bind("<Return>", lambda e: self._buscar())

        tk.Button(frame_busq, text="🔍",
                  font=("Segoe UI", 10), bg=self.COLOR_BTN, fg="white",
                  relief="flat", cursor="hand2",
                  command=self._buscar).pack(side="left", padx=(4, 0), ipady=4, ipadx=6)

        # Resultados con instrucción de doble clic
        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=8)

        frame_lbl_res = tk.Frame(parent, bg=self.COLOR_BG)
        frame_lbl_res.pack(fill="x")
        tk.Label(frame_lbl_res, text="Coincidencias  —  ",
                 font=("Segoe UI", 9, "bold"), bg=self.COLOR_BG).pack(side="left")
        tk.Label(frame_lbl_res,
                 text="doble clic para insertar firma",
                 font=("Segoe UI", 8, "italic"), bg=self.COLOR_BG,
                 fg="#2E6DA4").pack(side="left")

        frame_res = tk.Frame(parent, bg=self.COLOR_BG)
        frame_res.pack(fill="both", expand=True, pady=(4, 0))

        cols2 = ("fila_ref", "palabra")
        self.tree_resultados = ttk.Treeview(
            frame_res, columns=cols2, show="headings", height=10, selectmode="browse"
        )
        self.tree_resultados.heading("fila_ref", text="Ubicación")
        self.tree_resultados.heading("palabra",  text="Palabra encontrada")
        self.tree_resultados.column("fila_ref", width=80,  anchor="center", stretch=False)
        self.tree_resultados.column("palabra",  width=190)

        sb_res = ttk.Scrollbar(frame_res, orient="vertical", command=self.tree_resultados.yview)
        self.tree_resultados.configure(yscrollcommand=sb_res.set)
        self.tree_resultados.pack(side="left", fill="both", expand=True)
        sb_res.pack(side="left", fill="y")

        self.tree_resultados.bind("<<TreeviewSelect>>", self._al_seleccionar_resultado)
        # ── DOBLE CLIC → insertar firma ──────────────────────────────────────
        self.tree_resultados.bind("<Double-1>", self._al_doble_clic_resultado)

        self.label_conteo = tk.Label(parent, text="",
                                     font=("Segoe UI", 9), bg=self.COLOR_BG, fg=self.COLOR_OK)
        self.label_conteo.pack(anchor="w", pady=(4, 0))

    # ──────────────────────────────────────────
    # PANEL DERECHO — vista previa
    # ──────────────────────────────────────────

    def _construir_panel_preview(self, parent):
        # ── Encabezado del archivo ──────────────────────────────────────
        self.frame_info_archivo = tk.Frame(parent, bg="#DDE6F0", padx=10, pady=6)
        self.frame_info_archivo.pack(fill="x")

        self.label_info_tipo = tk.Label(
            self.frame_info_archivo,
            text="Selecciona un archivo de la lista",
            font=("Segoe UI", 10, "bold"),
            bg="#DDE6F0", fg=self.COLOR_ACCENT
        )
        self.label_info_tipo.pack(side="left")

        self.label_info_ruta = tk.Label(
            self.frame_info_archivo, text="",
            font=("Segoe UI", 8), bg="#DDE6F0", fg="#666"
        )
        self.label_info_ruta.pack(side="left", padx=(10, 0))

        # ── Contenedor principal del preview (apila Text y Tabla) ───────
        self.frame_preview_contenedor = tk.Frame(parent, bg=self.COLOR_BG)
        self.frame_preview_contenedor.pack(fill="both", expand=True, pady=(6, 0))

        # ── Vista TEXTO (Word / TXT) ─────────────────────────────────────
        self.frame_text_wrap = tk.Frame(self.frame_preview_contenedor, bg=self.COLOR_BG)
        self.text_preview = tk.Text(
            self.frame_text_wrap, font=("Courier New", 9), bg="#FAFCFF",
            relief="flat", wrap="none", state="disabled", padx=8, pady=6
        )
        sb_y = ttk.Scrollbar(self.frame_text_wrap, orient="vertical",   command=self.text_preview.yview)
        sb_x = ttk.Scrollbar(self.frame_text_wrap, orient="horizontal", command=self.text_preview.xview)
        self.text_preview.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right",  fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.text_preview.pack(fill="both", expand=True)

        self.text_preview.tag_configure(
            "firma_mark",
            background=self.COLOR_MARCA, foreground=self.COLOR_MARCA_FG,
            font=("Courier New", 9, "bold")
        )
        self.text_preview.tag_configure("linea_activa", background="#FFFDE7")

        # ── Vista TABLA (Excel) ──────────────────────────────────────────
        self.frame_tabla_wrap = tk.Frame(self.frame_preview_contenedor, bg=self.COLOR_BG)

        # Notebook de pestañas por hoja
        self.nb_hojas = ttk.Notebook(self.frame_tabla_wrap)
        self.nb_hojas.pack(fill="both", expand=True)

        # Variable para rastrear las pestañas activas
        self._tabs_excel: list = []

        # Por defecto mostrar vista texto
        self.frame_text_wrap.pack(fill="both", expand=True)

        # ── Leyenda + estado ─────────────────────────────────────────────
        frame_leyenda = tk.Frame(parent, bg=self.COLOR_BG)
        frame_leyenda.pack(fill="x", pady=(4, 0))

        tk.Label(frame_leyenda, text="  Resaltado: ",
                 font=("Segoe UI", 8), bg=self.COLOR_BG, fg="#555").pack(side="left")
        tk.Label(frame_leyenda, text=" firma ",
                 font=("Courier New", 8, "bold"),
                 bg=self.COLOR_MARCA, fg=self.COLOR_MARCA_FG).pack(side="left")
        tk.Label(frame_leyenda, text="  |  Fondo amarillo = fila encontrada  |  ",
                 font=("Segoe UI", 8), bg=self.COLOR_BG, fg="#555").pack(side="left")
        tk.Label(frame_leyenda, text="doble clic en resultado → insertar firma",
                 font=("Segoe UI", 8, "italic"),
                 bg=self.COLOR_BG, fg="#2E6DA4").pack(side="left")

        self.label_estado_insercion = tk.Label(
            parent, text="",
            font=("Segoe UI", 9, "bold"),
            bg=self.COLOR_BG, fg=self.COLOR_OK,
            wraplength=500
        )
        self.label_estado_insercion.pack(pady=(4, 0))

    # ──────────────────────────────────────────
    # ACCIONES — archivos
    # ──────────────────────────────────────────

    def _agregar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccionar documento(s)",
            filetypes=[
                ("Documentos Office y texto", "*.xlsx *.docx *.txt"),
                ("Excel",  "*.xlsx"), ("Word", "*.docx"),
                ("Texto",  "*.txt"),  ("Todos", "*.*"),
            ]
        )
        for ruta in rutas:
            if ruta not in self.archivos:
                self.archivos.append(ruta)
                tipo  = _detectar_tipo(ruta)
                icono = ICONOS_TIPO[tipo]
                self.tree_archivos.insert(
                    "", "end", iid=ruta,
                    values=(icono, os.path.basename(ruta), tipo)
                )

    def _limpiar_archivos(self):
        self.archivos.clear()
        self.contenidos.clear()
        self._resultado_meta.clear()
        for item in self.tree_archivos.get_children():
            self.tree_archivos.delete(item)
        for item in self.tree_resultados.get_children():
            self.tree_resultados.delete(item)
        self._limpiar_preview()
        self.label_conteo.config(text="")
        self.label_estado_insercion.config(text="")

    def _al_seleccionar_archivo(self, event=None):
        sel = self.tree_archivos.selection()
        if not sel:
            return
        ruta  = sel[0]
        tipo  = _detectar_tipo(ruta)
        icono = ICONOS_TIPO[tipo]
        color = COLORES_TIPO[tipo]

        self.label_info_tipo.config(
            text=f"{icono}  {os.path.basename(ruta)}  —  {tipo}", fg=color
        )
        self.label_info_ruta.config(text=ruta)

        if ruta not in self.contenidos:
            self.contenidos[ruta] = _leer_contenido_texto(ruta)
        self._cargar_preview(ruta)

        palabra = self.var_buscar.get().strip()
        if palabra:
            self._resaltar_en_preview(self.contenidos[ruta], palabra)

    def _limpiar_preview(self):
        self.text_preview.config(state="normal")
        self.text_preview.delete("1.0", tk.END)
        self.text_preview.config(state="disabled")
        self._ocultar_tabla_excel()
        self.label_info_tipo.config(
            text="Selecciona un archivo de la lista", fg=self.COLOR_ACCENT
        )
        self.label_info_ruta.config(text="")

    def _ocultar_tabla_excel(self):
        self.frame_tabla_wrap.pack_forget()
        self.frame_text_wrap.pack(fill="both", expand=True)

    def _mostrar_tabla_excel(self, ruta: str, palabra_busqueda: str = ""):
        """Construye la vista de tabla con celdas reales para archivos Excel."""
        # Limpiar pestañas anteriores
        for tab in self.nb_hojas.tabs():
            self.nb_hojas.forget(tab)
        self._tabs_excel.clear()

        hojas = _leer_excel_como_tabla(ruta)
        if not hojas:
            self._ocultar_tabla_excel()
            return

        self.frame_text_wrap.pack_forget()
        self.frame_tabla_wrap.pack(fill="both", expand=True)

        palabra_lower = palabra_busqueda.lower()

        for hoja in hojas:
            # Frame con canvas + scrollbar por hoja
            frame_hoja = tk.Frame(self.nb_hojas, bg="#FAFCFF")
            self.nb_hojas.add(frame_hoja, text=f"  {hoja['nombre']}  ")

            canvas = tk.Canvas(frame_hoja, bg="#FAFCFF", highlightthickness=0)
            sb_vy = ttk.Scrollbar(frame_hoja, orient="vertical",   command=canvas.yview)
            sb_hx = ttk.Scrollbar(frame_hoja, orient="horizontal", command=canvas.xview)
            canvas.configure(yscrollcommand=sb_vy.set, xscrollcommand=sb_hx.set)

            sb_vy.pack(side="right",  fill="y")
            sb_hx.pack(side="bottom", fill="x")
            canvas.pack(fill="both", expand=True)

            # Frame interior para las celdas
            interior = tk.Frame(canvas, bg="#FAFCFF")
            canvas_window = canvas.create_window((0, 0), window=interior, anchor="nw")

            def _on_configure(event, cv=canvas):
                cv.configure(scrollregion=cv.bbox("all"))
            interior.bind("<Configure>", _on_configure)

            # Calcular anchos de columna en píxeles
            col_widths = hoja["col_widths"]
            n_cols = max((len(f[1]) for f in hoja["filas"]), default=1)
            px_widths = [max(col_widths.get(ci, 4) * 7, 50) for ci in range(n_cols)]

            # Encabezado de columnas (A, B, C...)
            from openpyxl.utils import get_column_letter
            # Número de fila (esquina)
            tk.Label(interior, text="", width=5,
                     bg="#C5D5E8", fg="#1F3864",
                     font=("Calibri", 8, "bold"),
                     relief="groove", borderwidth=1
                     ).grid(row=0, column=0, sticky="nsew", ipadx=2, ipady=2)

            for ci in range(n_cols):
                tk.Label(interior,
                         text=get_column_letter(ci + 1),
                         width=max(px_widths[ci] // 7, 6),
                         bg="#C5D5E8", fg="#1F3864",
                         font=("Calibri", 8, "bold"),
                         relief="groove", borderwidth=1
                         ).grid(row=0, column=ci + 1, sticky="nsew", ipadx=2, ipady=2)

            # Filas de datos
            for grid_row, (n_fila, celdas) in enumerate(hoja["filas"], start=1):
                # Número de fila
                tk.Label(interior, text=str(n_fila), width=5,
                         bg="#DDE6F0", fg="#1F3864",
                         font=("Calibri", 8),
                         relief="groove", borderwidth=1
                         ).grid(row=grid_row, column=0, sticky="nsew", ipadx=2, ipady=1)

                # Celdas
                for ci, valor in enumerate(celdas):
                    # Resaltar si contiene la palabra buscada
                    contiene = palabra_lower and palabra_lower in valor.lower()
                    bg_color = self.COLOR_MARCA if contiene else "#FFFFFF"
                    fg_color = self.COLOR_MARCA_FG if contiene else "#222222"
                    fuente   = ("Calibri", 8, "bold") if contiene else ("Calibri", 8)

                    tk.Label(interior,
                             text=valor,
                             width=max(px_widths[ci] // 7, 6),
                             bg=bg_color, fg=fg_color,
                             font=fuente,
                             anchor="w",
                             relief="groove", borderwidth=1
                             ).grid(row=grid_row, column=ci + 1,
                                    sticky="nsew", ipadx=3, ipady=1)

            self._tabs_excel.append(frame_hoja)

    def _resaltar_tabla_excel(self, palabra: str, n_fila_excel: int = None, nombre_hoja: str = None):
        """Navega a la pestaña correcta y hace scroll a la fila con la firma."""
        if not self._tabs_excel:
            return
        # Seleccionar la pestaña de la hoja correspondiente
        if nombre_hoja:
            for idx, tab_id in enumerate(self.nb_hojas.tabs()):
                tab_text = self.nb_hojas.tab(tab_id, "text").strip()
                if tab_text == nombre_hoja:
                    self.nb_hojas.select(idx)
                    break

    def _cargar_preview(self, ruta: str):
        tipo = _detectar_tipo(ruta)
        palabra = self.var_buscar.get().strip()

        if tipo == TIPO_EXCEL:
            # Vista tabla con celdas reales
            self._mostrar_tabla_excel(ruta, palabra_busqueda=palabra)
        else:
            # Vista texto para Word / TXT
            self._ocultar_tabla_excel()
            contenido = self.contenidos.get(ruta, "")
            self.text_preview.config(state="normal")
            self.text_preview.delete("1.0", tk.END)
            self.text_preview.insert("1.0", contenido)
            self.text_preview.config(state="disabled")

    # ──────────────────────────────────────────
    # BÚSQUEDA
    # ──────────────────────────────────────────

    def _buscar(self):
        palabra = self.var_buscar.get().strip()
        if not palabra:
            return
        if not self.archivos:
            messagebox.showinfo("Sin archivos", "Agrega al menos un archivo antes de buscar.")
            return

        for item in self.tree_resultados.get_children():
            self.tree_resultados.delete(item)
        self._resultado_meta.clear()

        total = 0
        for ruta in self.archivos:
            if ruta not in self.contenidos:
                self.contenidos[ruta] = _leer_contenido_texto(ruta)

            resultados = _buscar_firma_en_texto(self.contenidos[ruta], palabra)

            for r in resultados:
                total += 1
                # Extraer la palabra exacta del contexto
                import re as _re
                m = _re.search(re.escape(palabra), r["contexto"], _re.IGNORECASE)
                palabra_hallada = m.group(0) if m else palabra

                # Referencia de ubicación: Fila Excel o párrafo Word
                raw = r.get("linea_raw", "")
                m2 = _re.match(r"^\[F(\d+)\|", raw)
                m3 = _re.match(r"^\[P(\d+)\]", raw)
                if m2:
                    fila_ref = f"Fila {m2.group(1)}"
                elif m3:
                    fila_ref = f"Párr. {m3.group(1)}"
                else:
                    fila_ref = f"Línea {r['linea']}"

                iid = self.tree_resultados.insert(
                    "", "end",
                    values=(fila_ref, palabra_hallada),
                    tags=(ruta,)
                )
                # Guardar metadatos para la inserción posterior
                self._resultado_meta[iid] = {
                    "ruta":      ruta,
                    "linea":     r["linea"],
                    "tipo":      _detectar_tipo(ruta),
                    "linea_raw": r.get("linea_raw", ""),
                }

            sel = self.tree_archivos.selection()
            if sel and sel[0] == ruta:
                self._resaltar_en_preview(self.contenidos[ruta], palabra)

        if total == 0:
            self.label_conteo.config(
                text=f"No se encontró '{palabra}' en los archivos cargados.",
                fg=self.COLOR_ERR
            )
        else:
            self.label_conteo.config(
                text=f"✅  {total} coincidencia(s).  Doble clic para insertar firma.",
                fg=self.COLOR_OK
            )

        if not self.tree_archivos.selection() and self.archivos:
            self.tree_archivos.selection_set(self.archivos[0])
            self._al_seleccionar_archivo()

    def _resaltar_en_preview(self, contenido: str, palabra: str):
        self.text_preview.config(state="normal")
        self.text_preview.tag_remove("firma_mark",   "1.0", tk.END)
        self.text_preview.tag_remove("linea_activa", "1.0", tk.END)

        if not palabra:
            self.text_preview.config(state="disabled")
            return

        start = "1.0"
        while True:
            pos = self.text_preview.search(palabra, start, stopindex=tk.END, nocase=True)
            if not pos:
                break
            end = f"{pos}+{len(palabra)}c"
            self.text_preview.tag_add("firma_mark", pos, end)
            start = end

        self.text_preview.config(state="disabled")

    def _al_seleccionar_resultado(self, event=None):
        sel = self.tree_resultados.selection()
        if not sel:
            return
        item = sel[0]
        vals = self.tree_resultados.item(item, "values")
        tags = self.tree_resultados.item(item, "tags")
        if not vals or not tags:
            return

        ruta = tags[0]
        meta = self._resultado_meta.get(item, {})

        # Asegurar que el archivo correcto esté seleccionado en el panel izquierdo
        if not self.tree_archivos.selection() or self.tree_archivos.selection()[0] != ruta:
            self.tree_archivos.selection_set(ruta)
            self._al_seleccionar_archivo()

        tipo = _detectar_tipo(ruta)

        if tipo == TIPO_EXCEL:
            # Navegar a la hoja correcta en la vista tabla
            linea_raw = meta.get("linea_raw", "")
            m = re.match(r"^\[F\d+\|(.+?)\]", linea_raw)
            nombre_hoja = m.group(1) if m else None
            self._resaltar_tabla_excel(
                self.var_buscar.get().strip(),
                nombre_hoja=nombre_hoja
            )
        else:
            # Vista texto: resaltar línea activa
            n_linea = meta.get("linea", 1)
            linea_idx = f"{n_linea}.0"
            self.text_preview.config(state="normal")
            self.text_preview.tag_remove("linea_activa", "1.0", tk.END)
            self.text_preview.tag_add("linea_activa", linea_idx, f"{n_linea}.end")
            self.text_preview.see(linea_idx)
            self.text_preview.config(state="disabled")

    # ──────────────────────────────────────────
    # DOBLE CLIC → INSERTAR FIRMA
    # ──────────────────────────────────────────

    def _al_doble_clic_resultado(self, event=None):
        """Inserta la firma en el documento en la posición del resultado seleccionado."""
        sel = self.tree_resultados.selection()
        if not sel:
            return
        iid  = sel[0]
        meta = self._resultado_meta.get(iid)
        if not meta:
            return

        # Verificar que hay una firma activa
        if not self.firma_data:
            messagebox.showwarning(
                "Sin firma activa",
                "No hay una firma generada.\n\n"
                "Ve a la pantalla principal, genera una firma y luego abre el buscador."
            )
            return

        ruta   = meta["ruta"]
        linea  = meta["linea"]
        tipo   = meta["tipo"]
        nombre = os.path.basename(ruta)
        vals   = self.tree_resultados.item(iid, "values")
        contexto = vals[2] if vals else ""

        # ── Confirmación ─────────────────────────────────────────────────────
        msg = (
            f"¿Insertar la firma en este documento?\n\n"
            f"  Archivo:  {nombre}\n"
            f"  Línea:    {linea}\n"
            f"  Contexto: {contexto[:60]}{'…' if len(contexto) > 60 else ''}\n\n"
            f"  Firmado por:  {self.firma_data.get('nombre', '')}\n"
            f"  ID Firma:     {self.firma_data.get('id_firma', '')}"
        )
        if not messagebox.askyesno("Confirmar inserción de firma", msg):
            return

        # ── Insertar según tipo ───────────────────────────────────────────────
        if tipo == TIPO_EXCEL:
            resultado_msg = _insertar_firma_excel(ruta, linea, self.firma_data, linea_raw=meta.get("linea_raw", ""))
        elif tipo == TIPO_WORD:
            resultado_msg = _insertar_firma_word(ruta, linea, self.firma_data)
        else:
            messagebox.showinfo(
                "Tipo no soportado",
                f"La inserción de firma solo está disponible\npara archivos Excel (.xlsx) y Word (.docx).\n\nTipo detectado: {tipo}"
            )
            return

        # ── Mostrar resultado ─────────────────────────────────────────────────
        exito = resultado_msg.startswith("✅")
        self.label_estado_insercion.config(
            text=resultado_msg,
            fg=self.COLOR_OK if exito else self.COLOR_ERR
        )

        if exito:
            # Recargar preview para reflejar los cambios
            if ruta in self.contenidos:
                del self.contenidos[ruta]
            self._al_seleccionar_archivo()
            palabra = self.var_buscar.get().strip()
            if palabra:
                self._resaltar_en_preview(self.contenidos.get(ruta, ""), palabra)

            messagebox.showinfo("Firma insertada", resultado_msg)
        else:
            messagebox.showerror("Error al insertar", resultado_msg)

    # ──────────────────────────────────────────
    def iniciar(self):
        self.top.mainloop()


# ─────────────────────────────────────────────
# FUNCIÓN DE CONVENIENCIA
# ─────────────────────────────────────────────

def abrir_buscador(parent=None, firma_data: dict = None):
    buscador = VentanaBuscador(parent=parent, firma_data=firma_data)
    if firma_data:
        id_firma = firma_data.get("id_firma", "")
        if id_firma:
            buscador.var_buscar.set(id_firma[:8])
    return buscador


# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = VentanaBuscador()
    app.iniciar()