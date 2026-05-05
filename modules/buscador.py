"""
buscador.py — Ventana de búsqueda de firma en documentos
Permite buscar la palabra 'firma', diferenciar tipos de archivo,
y mostrar una vista previa con la ubicación de la firma.
"""

import os
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
    TIPO_EXCEL: "#1E6B3C",   # verde Excel
    TIPO_WORD:  "#1E3A6B",   # azul Word
    TIPO_TEXTO: "#555555",
    TIPO_OTRO:  "#888888",
}


# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────

def _detectar_tipo(ruta: str) -> str:
    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".xlsx":
        return TIPO_EXCEL
    if ext == ".docx":
        return TIPO_WORD
    if ext == ".txt":
        return TIPO_TEXTO
    return TIPO_OTRO


def _leer_contenido_texto(ruta: str) -> str:
    """
    Extrae el contenido de texto del archivo según su tipo.
    Devuelve el texto plano o un mensaje de error.
    """
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
                    for fila in ws.iter_rows(values_only=True):
                        celdas = [str(c) if c is not None else "" for c in fila]
                        if any(c.strip() for c in celdas):
                            lineas.append("\t".join(celdas))
                return "\n".join(lineas)
            except ImportError:
                return "[openpyxl no instalado — no se puede leer .xlsx]"

        elif tipo == TIPO_WORD:
            try:
                from docx import Document
                doc = Document(ruta)
                parrafos = [p.text for p in doc.paragraphs]
                # Incluir tablas
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


def _buscar_firma_en_texto(contenido: str, palabra: str = "firma") -> list[dict]:
    """
    Busca todas las ocurrencias de `palabra` (sin distinción de mayúsculas)
    en `contenido` y devuelve una lista de coincidencias con contexto.

    Cada elemento: {"linea": int, "columna": int, "contexto": str}
    """
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
            })
            col = idx + 1
    return resultados


# ─────────────────────────────────────────────
# VENTANA BUSCADOR
# ─────────────────────────────────────────────

class VentanaBuscador:
    """
    Ventana independiente para:
    • Seleccionar uno o varios archivos.
    • Buscar la palabra 'firma' (o cualquier otra).
    • Diferenciar visualmente el tipo de archivo.
    • Mostrar vista previa resaltada del contenido con la firma.
    """

    COLOR_BG      = "#F0F4F8"
    COLOR_ACCENT  = "#1F3864"
    COLOR_BTN     = "#2E6DA4"
    COLOR_OK      = "#1E8449"
    COLOR_ERR     = "#C0392B"
    COLOR_MARCA   = "#FFD700"   # fondo del resaltado de firma
    COLOR_MARCA_FG = "#000000"

    def __init__(self, parent: tk.Tk | None = None, firma_data: dict | None = None):
        """
        parent    : ventana padre (opcional, para mantener jerarquía).
        firma_data: dict con la firma actual (de AppFirmaDigital.ultima_firma).
                    Si se pasa, se usa su texto como palabra a buscar adicionalmente.
        """
        self.firma_data = firma_data

        if parent:
            self.top = tk.Toplevel(parent)
        else:
            self.top = tk.Tk()

        self.top.title("🔎 Buscador de Firma en Documentos")
        self.top.configure(bg=self.COLOR_BG)
        self.top.resizable(True, True)
        self._centrar_ventana(820, 620)

        # Estado interno
        self.archivos: list[str] = []          # rutas seleccionadas
        self.contenidos: dict[str, str] = {}   # ruta → texto extraído
        self.var_buscar = tk.StringVar(value="firma")

        self._construir_ui()

    # ──────────────────────────────────────────
    # GEOMETRÍA
    # ──────────────────────────────────────────

    def _centrar_ventana(self, ancho, alto):
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth()  // 2) - (ancho // 2)
        y = (self.top.winfo_screenheight() // 2) - (alto  // 2)
        self.top.geometry(f"{ancho}x{alto}+{x}+{y}")

    # ──────────────────────────────────────────
    # UI PRINCIPAL
    # ──────────────────────────────────────────

    def _construir_ui(self):
        # ── Encabezado ──
        frame_header = tk.Frame(self.top, bg=self.COLOR_ACCENT, height=50)
        frame_header.pack(fill="x")
        tk.Label(
            frame_header,
            text="🔎  Buscador de Firma en Documentos",
            font=("Segoe UI", 12, "bold"),
            fg="white", bg=self.COLOR_ACCENT
        ).pack(side="left", padx=20, pady=12)

        # ── Panel izquierdo (archivos + resultados) ──
        panel_izq = tk.Frame(self.top, bg=self.COLOR_BG, width=300)
        panel_izq.pack(side="left", fill="y", padx=(10, 5), pady=10)
        panel_izq.pack_propagate(False)

        self._construir_panel_archivos(panel_izq)

        # ── Separador vertical ──
        ttk.Separator(self.top, orient="vertical").pack(side="left", fill="y", pady=10)

        # ── Panel derecho (vista previa) ──
        panel_der = tk.Frame(self.top, bg=self.COLOR_BG)
        panel_der.pack(side="left", fill="both", expand=True, padx=(5, 10), pady=10)

        self._construir_panel_preview(panel_der)

    # ──────────────────────────────────────────
    # PANEL IZQUIERDO — archivos + búsqueda
    # ──────────────────────────────────────────

    def _construir_panel_archivos(self, parent):
        # Botones de selección
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

        # Lista de archivos cargados
        tk.Label(
            parent, text="Archivos cargados:",
            font=("Segoe UI", 9, "bold"), bg=self.COLOR_BG
        ).pack(anchor="w")

        frame_lista = tk.Frame(parent, bg=self.COLOR_BG)
        frame_lista.pack(fill="both", expand=True)

        cols = ("icono", "nombre", "tipo")
        self.tree_archivos = ttk.Treeview(
            frame_lista, columns=cols, show="headings", height=8,
            selectmode="browse"
        )
        self.tree_archivos.heading("icono",  text="")
        self.tree_archivos.heading("nombre", text="Nombre")
        self.tree_archivos.heading("tipo",   text="Tipo")

        self.tree_archivos.column("icono",  width=28, anchor="center", stretch=False)
        self.tree_archivos.column("nombre", width=160)
        self.tree_archivos.column("tipo",   width=90)

        sb_arch = ttk.Scrollbar(frame_lista, orient="vertical",
                                command=self.tree_archivos.yview)
        self.tree_archivos.configure(yscrollcommand=sb_arch.set)
        self.tree_archivos.pack(side="left", fill="both", expand=True)
        sb_arch.pack(side="left", fill="y")

        self.tree_archivos.bind("<<TreeviewSelect>>", self._al_seleccionar_archivo)

        # ── Búsqueda ──
        sep = ttk.Separator(parent, orient="horizontal")
        sep.pack(fill="x", pady=8)

        tk.Label(
            parent, text="Palabra a buscar:",
            font=("Segoe UI", 9, "bold"), bg=self.COLOR_BG
        ).pack(anchor="w")

        frame_busq = tk.Frame(parent, bg=self.COLOR_BG)
        frame_busq.pack(fill="x", pady=(4, 0))

        self.entry_buscar = ttk.Entry(
            frame_busq, textvariable=self.var_buscar,
            font=("Segoe UI", 10)
        )
        self.entry_buscar.pack(side="left", fill="x", expand=True, ipady=4)
        self.entry_buscar.bind("<Return>", lambda e: self._buscar())

        tk.Button(
            frame_busq, text="🔍",
            font=("Segoe UI", 10), bg=self.COLOR_BTN, fg="white",
            relief="flat", cursor="hand2",
            command=self._buscar
        ).pack(side="left", padx=(4, 0), ipady=4, ipadx=6)

        # ── Resultados de búsqueda ──
        sep2 = ttk.Separator(parent, orient="horizontal")
        sep2.pack(fill="x", pady=8)

        tk.Label(
            parent, text="Coincidencias encontradas:",
            font=("Segoe UI", 9, "bold"), bg=self.COLOR_BG
        ).pack(anchor="w")

        frame_res = tk.Frame(parent, bg=self.COLOR_BG)
        frame_res.pack(fill="both", expand=True, pady=(4, 0))

        cols2 = ("linea", "col", "contexto")
        self.tree_resultados = ttk.Treeview(
            frame_res, columns=cols2, show="headings", height=8,
            selectmode="browse"
        )
        self.tree_resultados.heading("linea",    text="Línea")
        self.tree_resultados.heading("col",      text="Col")
        self.tree_resultados.heading("contexto", text="Contexto")

        self.tree_resultados.column("linea",    width=40,  anchor="center", stretch=False)
        self.tree_resultados.column("col",      width=36,  anchor="center", stretch=False)
        self.tree_resultados.column("contexto", width=180)

        sb_res = ttk.Scrollbar(frame_res, orient="vertical",
                               command=self.tree_resultados.yview)
        self.tree_resultados.configure(yscrollcommand=sb_res.set)
        self.tree_resultados.pack(side="left", fill="both", expand=True)
        sb_res.pack(side="left", fill="y")

        self.tree_resultados.bind("<<TreeviewSelect>>", self._al_seleccionar_resultado)

        self.label_conteo = tk.Label(
            parent, text="",
            font=("Segoe UI", 9), bg=self.COLOR_BG, fg=self.COLOR_OK
        )
        self.label_conteo.pack(anchor="w", pady=(4, 0))

    # ──────────────────────────────────────────
    # PANEL DERECHO — vista previa
    # ──────────────────────────────────────────

    def _construir_panel_preview(self, parent):
        # Cabecera del archivo seleccionado
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
            self.frame_info_archivo,
            text="",
            font=("Segoe UI", 8),
            bg="#DDE6F0", fg="#666"
        )
        self.label_info_ruta.pack(side="left", padx=(10, 0))

        # Área de texto con scroll
        frame_text = tk.Frame(parent, bg=self.COLOR_BG)
        frame_text.pack(fill="both", expand=True, pady=(6, 0))

        self.text_preview = tk.Text(
            frame_text,
            font=("Courier New", 9),
            bg="#FAFCFF",
            relief="flat",
            wrap="none",
            state="disabled",
            padx=8, pady=6
        )

        sb_y = ttk.Scrollbar(frame_text, orient="vertical",   command=self.text_preview.yview)
        sb_x = ttk.Scrollbar(frame_text, orient="horizontal", command=self.text_preview.xview)
        self.text_preview.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

        sb_y.pack(side="right",  fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.text_preview.pack(fill="both", expand=True)

        # Configurar tags de resaltado
        self.text_preview.tag_configure(
            "firma_mark",
            background=self.COLOR_MARCA,
            foreground=self.COLOR_MARCA_FG,
            font=("Courier New", 9, "bold")
        )
        self.text_preview.tag_configure(
            "linea_activa",
            background="#FFFDE7"
        )

        # Leyenda
        frame_leyenda = tk.Frame(parent, bg=self.COLOR_BG)
        frame_leyenda.pack(fill="x", pady=(4, 0))

        tk.Label(
            frame_leyenda,
            text="  ⬛ Resaltado: ",
            font=("Segoe UI", 8), bg=self.COLOR_BG, fg="#555"
        ).pack(side="left")

        lbl_marca = tk.Label(
            frame_leyenda,
            text=" firma ",
            font=("Courier New", 8, "bold"),
            bg=self.COLOR_MARCA, fg=self.COLOR_MARCA_FG
        )
        lbl_marca.pack(side="left")

        tk.Label(
            frame_leyenda,
            text="  |  Fondo amarillo claro = línea seleccionada",
            font=("Segoe UI", 8), bg=self.COLOR_BG, fg="#555"
        ).pack(side="left")

    # ──────────────────────────────────────────
    # ACCIONES
    # ──────────────────────────────────────────

    def _agregar_archivos(self):
        """Abre diálogo multi-selección y carga los archivos."""
        rutas = filedialog.askopenfilenames(
            title="Seleccionar documento(s)",
            filetypes=[
                ("Documentos Office y texto", "*.xlsx *.docx *.txt"),
                ("Excel",  "*.xlsx"),
                ("Word",   "*.docx"),
                ("Texto",  "*.txt"),
                ("Todos",  "*.*"),
            ]
        )
        for ruta in rutas:
            if ruta not in self.archivos:
                self.archivos.append(ruta)
                tipo  = _detectar_tipo(ruta)
                icono = ICONOS_TIPO[tipo]
                nombre = os.path.basename(ruta)
                self.tree_archivos.insert(
                    "", "end",
                    iid=ruta,
                    values=(icono, nombre, tipo)
                )

    def _limpiar_archivos(self):
        """Elimina todos los archivos cargados."""
        self.archivos.clear()
        self.contenidos.clear()
        for item in self.tree_archivos.get_children():
            self.tree_archivos.delete(item)
        for item in self.tree_resultados.get_children():
            self.tree_resultados.delete(item)
        self._limpiar_preview()
        self.label_conteo.config(text="")

    def _al_seleccionar_archivo(self, event=None):
        """Carga y muestra la vista previa del archivo seleccionado."""
        sel = self.tree_archivos.selection()
        if not sel:
            return
        ruta = sel[0]   # iid == ruta
        tipo  = _detectar_tipo(ruta)
        icono = ICONOS_TIPO[tipo]
        color = COLORES_TIPO[tipo]

        self.label_info_tipo.config(
            text=f"{icono}  {os.path.basename(ruta)}  —  {tipo}",
            fg=color
        )
        self.label_info_ruta.config(text=ruta)

        # Extraer contenido si no está en caché
        if ruta not in self.contenidos:
            self.contenidos[ruta] = _leer_contenido_texto(ruta)

        self._cargar_preview(ruta)

        # Si ya había una búsqueda activa, re-resaltar
        palabra = self.var_buscar.get().strip()
        if palabra:
            self._resaltar_en_preview(self.contenidos[ruta], palabra)

    def _limpiar_preview(self):
        self.text_preview.config(state="normal")
        self.text_preview.delete("1.0", tk.END)
        self.text_preview.config(state="disabled")
        self.label_info_tipo.config(
            text="Selecciona un archivo de la lista",
            fg=self.COLOR_ACCENT
        )
        self.label_info_ruta.config(text="")

    def _cargar_preview(self, ruta: str):
        """Carga el contenido en el área de texto."""
        contenido = self.contenidos.get(ruta, "")
        self.text_preview.config(state="normal")
        self.text_preview.delete("1.0", tk.END)
        self.text_preview.insert("1.0", contenido)
        self.text_preview.config(state="disabled")

    def _buscar(self):
        """Busca la palabra en todos los archivos cargados."""
        palabra = self.var_buscar.get().strip()
        if not palabra:
            return
        if not self.archivos:
            messagebox.showinfo("Sin archivos", "Agrega al menos un archivo antes de buscar.")
            return

        # Limpiar resultados anteriores
        for item in self.tree_resultados.get_children():
            self.tree_resultados.delete(item)

        total = 0
        for ruta in self.archivos:
            if ruta not in self.contenidos:
                self.contenidos[ruta] = _leer_contenido_texto(ruta)

            resultados = _buscar_firma_en_texto(self.contenidos[ruta], palabra)
            nombre = os.path.basename(ruta)
            tipo   = _detectar_tipo(ruta)
            icono  = ICONOS_TIPO[tipo]

            for r in resultados:
                total += 1
                self.tree_resultados.insert(
                    "", "end",
                    values=(r["linea"], r["columna"], r["contexto"]),
                    tags=(ruta,)  # guardamos la ruta como tag para recuperarla
                )
            # Resaltar en preview si este archivo está seleccionado
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
                text=f"✅  {total} coincidencia(s) encontrada(s).",
                fg=self.COLOR_OK
            )

        # Si no hay archivo seleccionado, seleccionar el primero automáticamente
        if not self.tree_archivos.selection() and self.archivos:
            self.tree_archivos.selection_set(self.archivos[0])
            self._al_seleccionar_archivo()

    def _resaltar_en_preview(self, contenido: str, palabra: str):
        """Resalta todas las ocurrencias de `palabra` en el área de preview."""
        self.text_preview.config(state="normal")

        # Quitar resaltados anteriores
        self.text_preview.tag_remove("firma_mark",   "1.0", tk.END)
        self.text_preview.tag_remove("linea_activa", "1.0", tk.END)

        if not palabra:
            self.text_preview.config(state="disabled")
            return

        palabra_lower = palabra.lower()
        start = "1.0"
        while True:
            pos = self.text_preview.search(
                palabra_lower, start, stopindex=tk.END, nocase=True
            )
            if not pos:
                break
            end = f"{pos}+{len(palabra)}c"
            self.text_preview.tag_add("firma_mark", pos, end)
            start = end

        self.text_preview.config(state="disabled")

    def _al_seleccionar_resultado(self, event=None):
        """
        Al hacer click en un resultado, navega a esa línea en la vista previa
        del archivo correspondiente.
        """
        sel = self.tree_resultados.selection()
        if not sel:
            return

        item = sel[0]
        vals = self.tree_resultados.item(item, "values")
        tags = self.tree_resultados.item(item, "tags")
        if not vals or not tags:
            return

        n_linea = vals[0]
        ruta    = tags[0]

        # Asegurar que el archivo correcto esté seleccionado y en preview
        if not self.tree_archivos.selection() or self.tree_archivos.selection()[0] != ruta:
            self.tree_archivos.selection_set(ruta)
            self._al_seleccionar_archivo()

        # Navegar a la línea
        linea_idx = f"{n_linea}.0"
        self.text_preview.config(state="normal")
        self.text_preview.tag_remove("linea_activa", "1.0", tk.END)
        self.text_preview.tag_add("linea_activa", linea_idx, f"{n_linea}.end")
        self.text_preview.see(linea_idx)
        self.text_preview.config(state="disabled")

    # ──────────────────────────────────────────
    # ARRANQUE (si se usa como app independiente)
    # ──────────────────────────────────────────

    def iniciar(self):
        """Inicia el loop principal (solo cuando se usa como app independiente)."""
        self.top.mainloop()


# ─────────────────────────────────────────────
# FUNCIÓN DE CONVENIENCIA — llamar desde app.py
# ─────────────────────────────────────────────

def abrir_buscador(parent: tk.Tk | None = None, firma_data: dict | None = None):
    """
    Abre la ventana buscador como Toplevel hijo de `parent`.
    Se puede pasar `firma_data` (ultima_firma de AppFirmaDigital) para
    pre-cargar la firma actual como palabra de búsqueda.
    """
    buscador = VentanaBuscador(parent=parent, firma_data=firma_data)
    if firma_data:
        # Pre-cargar el ID de firma como término de búsqueda adicional
        id_firma = firma_data.get("id_firma", "")
        if id_firma:
            buscador.var_buscar.set(id_firma[:8])   # primeros 8 chars del hash
    return buscador


# ─────────────────────────────────────────────
# EJECUCIÓN DIRECTA
# ─────────────────────────────────────────────

if __name__ == "__main__":
    app = VentanaBuscador()
    app.iniciar()