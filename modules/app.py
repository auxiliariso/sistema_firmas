"""
app.py — Ventana principal de la aplicación
Se muestra tras el login exitoso. Permite firmar documentos y verificar firmas.
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from modules.firma  import crear_firma, formatear_firma_texto, validar_firma
from modules.utils  import insertar_firma_excel, insertar_firma_word
from modules.db     import obtener_firmas_usuario


# ─────────────────────────────────────────────
# VENTANA PRINCIPAL
# ─────────────────────────────────────────────

class AppFirmaDigital:
    """Ventana principal con pestañas: Firmar / Verificar / Historial."""

    COLOR_BG     = "#F0F4F8"
    COLOR_ACCENT = "#1F3864"
    COLOR_BTN    = "#2E6DA4"
    COLOR_OK     = "#1E8449"
    COLOR_ERR    = "#C0392B"

    def __init__(self, usuario_data: dict):
        self.usuario = usuario_data
        self.ultima_firma = None

        self.root = tk.Tk()
        self.root.title("🔏 Sistema de Firma Digital")
        self.root.configure(bg=self.COLOR_BG)
        self.root.resizable(True, True)
        self._centrar_ventana(620, 540)

        self._construir_ui()

    def _centrar_ventana(self, ancho, alto):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  // 2) - (ancho // 2)
        y = (self.root.winfo_screenheight() // 2) - (alto  // 2)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")

    # ──────────────────────────────────────────
    # CONSTRUCCIÓN DE UI
    # ──────────────────────────────────────────

    def _construir_ui(self):
        # Encabezado
        frame_header = tk.Frame(self.root, bg=self.COLOR_ACCENT, height=60)
        frame_header.pack(fill="x")

        tk.Label(
            frame_header,
            text=f"🔏  Firma Digital   |   👤 {self.usuario['nombre']}",
            font=("Segoe UI", 12, "bold"),
            fg="white", bg=self.COLOR_ACCENT
        ).pack(side="left", padx=20, pady=15)

        # Notebook con pestañas
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=("Segoe UI", 10), padding=[12, 5])

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Pestaña 1: Firmar documento
        self.tab_firmar = tk.Frame(self.notebook, bg=self.COLOR_BG)
        self.notebook.add(self.tab_firmar, text="  ✍ Firmar Documento  ")
        self._construir_tab_firmar()

        # Pestaña 2: Verificar firma
        self.tab_verificar = tk.Frame(self.notebook, bg=self.COLOR_BG)
        self.notebook.add(self.tab_verificar, text="  🔍 Verificar Firma  ")
        self._construir_tab_verificar()

        # Pestaña 3: Historial
        self.tab_historial = tk.Frame(self.notebook, bg=self.COLOR_BG)
        self.notebook.add(self.tab_historial, text="  📋 Historial  ")
        self._construir_tab_historial()

    # ──────────────────────────────────────────
    # TAB 1: FIRMAR
    # ──────────────────────────────────────────

    def _construir_tab_firmar(self):
        frame = tk.Frame(self.tab_firmar, bg=self.COLOR_BG, padx=30, pady=20)
        frame.pack(fill="both", expand=True)

        # Selección de documento
        tk.Label(
            frame, text="Documento a firmar:",
            font=("Segoe UI", 10, "bold"), bg=self.COLOR_BG
        ).pack(anchor="w")

        frame_doc = tk.Frame(frame, bg=self.COLOR_BG)
        frame_doc.pack(fill="x", pady=(4, 12))

        self.var_documento = tk.StringVar()
        self.entry_doc = ttk.Entry(
            frame_doc, textvariable=self.var_documento,
            font=("Segoe UI", 10), state="readonly"
        )
        self.entry_doc.pack(side="left", fill="x", expand=True, ipady=4)

        tk.Button(
            frame_doc, text="📂 Buscar",
            font=("Segoe UI", 9), bg="#E8EEF4",
            relief="flat", cursor="hand2",
            command=self._seleccionar_documento
        ).pack(side="left", padx=(6, 0), ipady=4, ipadx=8)

        # También permitir escribir nombre manual
        tk.Label(
            frame, text="— o escribe el nombre del documento —",
            font=("Segoe UI", 9), bg=self.COLOR_BG, fg="#888"
        ).pack()

        self.entry_doc_manual = ttk.Entry(frame, font=("Segoe UI", 10))
        self.entry_doc_manual.pack(fill="x", ipady=4, pady=(4, 16))

        # Botón firmar
        tk.Button(
            frame,
            text="  🔏  Generar Firma Digital  ",
            font=("Segoe UI", 11, "bold"),
            bg=self.COLOR_BTN, fg="white",
            activebackground="#1A4F7A", activeforeground="white",
            relief="flat", cursor="hand2",
            command=self._generar_firma
        ).pack(fill="x", ipady=8)

        # Área de resultado
        tk.Label(
            frame, text="Firma generada:",
            font=("Segoe UI", 10, "bold"), bg=self.COLOR_BG
        ).pack(anchor="w", pady=(16, 4))

        self.text_firma = tk.Text(
            frame, height=7,
            font=("Courier New", 10),
            state="disabled",
            bg="#EAF2FF", relief="flat",
            wrap="none"
        )
        self.text_firma.pack(fill="x")

        # Botones de acción post-firma
        frame_botones = tk.Frame(frame, bg=self.COLOR_BG)
        frame_botones.pack(fill="x", pady=(10, 0))

        self.btn_insertar_excel = tk.Button(
            frame_botones, text="📊 Insertar en Excel",
            font=("Segoe UI", 9), bg="#217346", fg="white",
            relief="flat", cursor="hand2",
            command=self._insertar_en_excel, state="disabled"
        )
        self.btn_insertar_excel.pack(side="left", ipadx=8, ipady=4, padx=(0, 6))

        self.btn_insertar_word = tk.Button(
            frame_botones, text="📝 Insertar en Word",
            font=("Segoe UI", 9), bg="#2B579A", fg="white",
            relief="flat", cursor="hand2",
            command=self._insertar_en_word, state="disabled"
        )
        self.btn_insertar_word.pack(side="left", ipadx=8, ipady=4)

        self.label_accion = tk.Label(
            frame, text="",
            font=("Segoe UI", 9), bg=self.COLOR_BG, fg=self.COLOR_OK,
            wraplength=500
        )
        self.label_accion.pack(pady=(8, 0))

    def _seleccionar_documento(self):
        """Abre diálogo para seleccionar un archivo existente."""
        ruta = filedialog.askopenfilename(
            title="Seleccionar documento",
            filetypes=[
                ("Documentos Office", "*.xlsx *.docx"),
                ("Excel", "*.xlsx"),
                ("Word",  "*.docx"),
                ("Todos", "*.*"),
            ]
        )
        if ruta:
            self.var_documento.set(ruta)
            self.entry_doc_manual.delete(0, tk.END)

    def _obtener_nombre_documento(self) -> str:
        """Retorna el nombre del documento seleccionado o escrito."""
        ruta = self.var_documento.get().strip()
        if ruta:
            return os.path.basename(ruta)
        manual = self.entry_doc_manual.get().strip()
        return manual if manual else "sin_nombre.docx"

    def _generar_firma(self):
        """Genera la firma digital y la muestra en pantalla."""
        nombre_doc = self._obtener_nombre_documento()

        try:
            self.ultima_firma = crear_firma(self.usuario, nombre_doc)
            texto = formatear_firma_texto(self.ultima_firma)

            # Mostrar en el Text widget
            self.text_firma.config(state="normal")
            self.text_firma.delete("1.0", tk.END)
            self.text_firma.insert("1.0", texto)
            self.text_firma.config(state="disabled")

            # Habilitar botones de inserción
            self.btn_insertar_excel.config(state="normal")
            self.btn_insertar_word.config(state="normal")
            self.label_accion.config(
                text=f"✅ Firma generada y guardada en el registro.",
                fg=self.COLOR_OK
            )

            # Refrescar historial
            self._refrescar_historial()

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar la firma:\n{e}")

    def _insertar_en_excel(self):
        """Inserta la firma en un archivo Excel."""
        if not self.ultima_firma:
            return
        ruta = self.var_documento.get().strip()
        if not ruta or not ruta.endswith('.xlsx'):
            # Preguntar dónde guardar
            ruta = filedialog.asksaveasfilename(
                title="Guardar / seleccionar Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=self._obtener_nombre_documento().replace('.docx', '.xlsx')
            )
        if ruta:
            try:
                insertar_firma_excel(ruta, self.ultima_firma)
                self.label_accion.config(
                    text=f"✅ Firma insertada en: {os.path.basename(ruta)}",
                    fg=self.COLOR_OK
                )
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def _insertar_en_word(self):
        """Inserta la firma en un archivo Word."""
        if not self.ultima_firma:
            return
        ruta = self.var_documento.get().strip()
        if not ruta or not ruta.endswith('.docx'):
            ruta = filedialog.asksaveasfilename(
                title="Guardar / seleccionar Word",
                defaultextension=".docx",
                filetypes=[("Word", "*.docx")],
                initialfile=self._obtener_nombre_documento().replace('.xlsx', '.docx')
            )
        if ruta:
            try:
                insertar_firma_word(ruta, self.ultima_firma)
                self.label_accion.config(
                    text=f"✅ Firma insertada en: {os.path.basename(ruta)}",
                    fg=self.COLOR_OK
                )
            except Exception as e:
                messagebox.showerror("Error", str(e))

    # ──────────────────────────────────────────
    # TAB 2: VERIFICAR
    # ──────────────────────────────────────────

    def _construir_tab_verificar(self):
        frame = tk.Frame(self.tab_verificar, bg=self.COLOR_BG, padx=30, pady=30)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text="Introduce el ID de firma para verificar su autenticidad:",
            font=("Segoe UI", 10), bg=self.COLOR_BG
        ).pack(anchor="w")

        frame_input = tk.Frame(frame, bg=self.COLOR_BG)
        frame_input.pack(fill="x", pady=(8, 16))

        self.entry_id_verificar = ttk.Entry(
            frame_input,
            font=("Courier New", 12),
        )
        self.entry_id_verificar.pack(side="left", fill="x", expand=True, ipady=5)
        self.entry_id_verificar.bind("<Return>", lambda e: self._verificar_firma())

        tk.Button(
            frame_input, text="🔍 Verificar",
            font=("Segoe UI", 10, "bold"),
            bg=self.COLOR_BTN, fg="white",
            relief="flat", cursor="hand2",
            command=self._verificar_firma
        ).pack(side="left", padx=(8, 0), ipady=5, ipadx=12)

        # Resultado de verificación
        self.frame_resultado_ver = tk.Frame(
            frame, bg="#EAF2FF",
            relief="flat", padx=16, pady=16
        )
        self.frame_resultado_ver.pack(fill="x", pady=(0, 0))

        self.label_resultado_ver = tk.Label(
            self.frame_resultado_ver,
            text="Ingresa un ID de firma arriba para verificar.",
            font=("Segoe UI", 10),
            bg="#EAF2FF", fg="#555",
            justify="left", anchor="w", wraplength=500
        )
        self.label_resultado_ver.pack(fill="x")

    def _verificar_firma(self):
        id_firma = self.entry_id_verificar.get().strip().upper()
        if not id_firma:
            return

        resultado = validar_firma(id_firma)

        if resultado['valida']:
            datos = resultado['datos']
            texto = (
                f"✅  FIRMA VÁLIDA\n\n"
                f"  Firmado por:  {datos['usuario']}\n"
                f"  Fecha:        {datos['fecha_hora']}\n"
                f"  Documento:    {datos['documento']}\n"
                f"  ID Firma:     {datos['id_firma']}"
            )
            self.label_resultado_ver.config(
                text=texto, fg=self.COLOR_OK,
                font=("Courier New", 10)
            )
            self.frame_resultado_ver.config(bg="#E8F5E9")
            self.label_resultado_ver.config(bg="#E8F5E9")
        else:
            self.label_resultado_ver.config(
                text="❌  FIRMA NO ENCONTRADA\n\nEste ID no existe en la base de datos.",
                fg=self.COLOR_ERR,
                font=("Segoe UI", 10)
            )
            self.frame_resultado_ver.config(bg="#FDECEA")
            self.label_resultado_ver.config(bg="#FDECEA")

    # ──────────────────────────────────────────
    # TAB 3: HISTORIAL
    # ──────────────────────────────────────────

    def _construir_tab_historial(self):
        frame = tk.Frame(self.tab_historial, bg=self.COLOR_BG, padx=10, pady=10)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame, text=f"Firmas registradas por: {self.usuario['nombre']}",
            font=("Segoe UI", 10, "bold"), bg=self.COLOR_BG
        ).pack(anchor="w", pady=(0, 8))

        # Treeview tabla
        cols = ("id_firma", "fecha_hora", "documento")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=14)

        self.tree.heading("id_firma",   text="ID Firma")
        self.tree.heading("fecha_hora", text="Fecha y Hora")
        self.tree.heading("documento",  text="Documento")

        self.tree.column("id_firma",   width=100, anchor="center")
        self.tree.column("fecha_hora", width=130, anchor="center")
        self.tree.column("documento",  width=280)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="left", fill="y")

        tk.Button(
            frame, text="🔄 Actualizar",
            font=("Segoe UI", 9),
            bg="#E8EEF4", relief="flat", cursor="hand2",
            command=self._refrescar_historial
        ).pack(pady=(8, 0), anchor="e")

        self._refrescar_historial()

    def _refrescar_historial(self):
        """Recarga el historial desde la base de datos."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        firmas = obtener_firmas_usuario(self.usuario['usuario'])
        for f in firmas:
            self.tree.insert("", "end", values=(f['id_firma'], f['fecha_hora'], f['documento']))

    # ──────────────────────────────────────────
    # ARRANQUE
    # ──────────────────────────────────────────

    def iniciar(self):
        self.root.mainloop()
