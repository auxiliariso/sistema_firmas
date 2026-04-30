"""
login.py — Módulo de login
Ventana de autenticación con tkinter.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

# Fix de path: asegura que la raíz del proyecto esté en sys.path
_ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT_DIR not in sys.path:
    sys.path.insert(0, _ROOT_DIR)
from modules.db import verificar_credenciales, obtener_intentos_restantes


# ─────────────────────────────────────────────
# VENTANA DE LOGIN
# ─────────────────────────────────────────────

class VentanaLogin:
    """
    Ventana modal de autenticación.
    Retorna los datos del usuario autenticado o None si se cancela.
    """

    def __init__(self, master=None):
        self.usuario_autenticado = None

        # Crear ventana
        self.root = tk.Tk() if master is None else tk.Toplevel(master)
        self.root.title("🔐 Firma Digital — Acceso")
        self.root.resizable(False, False)
        self._centrar_ventana(380, 320)
        self.root.configure(bg="#F0F4F8")

        self._construir_ui()

    def _centrar_ventana(self, ancho: int, alto: int):
        """Centra la ventana en la pantalla."""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  // 2) - (ancho // 2)
        y = (self.root.winfo_screenheight() // 2) - (alto  // 2)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")

    def _construir_ui(self):
        """Construye todos los elementos de la UI."""
        COLOR_BG     = "#F0F4F8"
        COLOR_ACCENT = "#1F3864"
        COLOR_BTN    = "#2E6DA4"
        COLOR_BTN_HV = "#1A4F7A"

        # ── Encabezado ─────────────────────────────
        frame_header = tk.Frame(self.root, bg=COLOR_ACCENT, height=70)
        frame_header.pack(fill="x")

        tk.Label(
            frame_header,
            text="🔏  Sistema de Firma Digital",
            font=("Segoe UI", 13, "bold"),
            fg="white", bg=COLOR_ACCENT
        ).pack(pady=18)

        # ── Formulario ─────────────────────────────
        frame_form = tk.Frame(self.root, bg=COLOR_BG, padx=40, pady=20)
        frame_form.pack(fill="both", expand=True)

        # Usuario
        tk.Label(
            frame_form, text="Usuario:",
            font=("Segoe UI", 10), bg=COLOR_BG, fg="#333333",
            anchor="w"
        ).pack(fill="x", pady=(0, 2))

        self.entry_usuario = ttk.Entry(frame_form, font=("Segoe UI", 11))
        self.entry_usuario.pack(fill="x", ipady=4)

        # Contraseña (oculta con show="•")
        tk.Label(
            frame_form, text="Contraseña:",
            font=("Segoe UI", 10), bg=COLOR_BG, fg="#333333",
            anchor="w"
        ).pack(fill="x", pady=(12, 2))

        self.entry_password = ttk.Entry(
            frame_form,
            font=("Segoe UI", 11),
            show="•"          # ← Campo tipo password, nunca muestra el texto
        )
        self.entry_password.pack(fill="x", ipady=4)

        # Label de estado (errores / advertencias)
        self.label_estado = tk.Label(
            frame_form, text="",
            font=("Segoe UI", 9), bg=COLOR_BG, fg="#C0392B",
            wraplength=280
        )
        self.label_estado.pack(pady=(8, 0))

        # ── Botón de login ──────────────────────────
        btn_frame = tk.Frame(self.root, bg=COLOR_BG, padx=40)
        btn_frame.pack(fill="x", pady=(0, 20))

        self.btn_login = tk.Button(
            btn_frame,
            text="  Iniciar sesión  ",
            font=("Segoe UI", 11, "bold"),
            bg=COLOR_BTN, fg="white",
            activebackground=COLOR_BTN_HV, activeforeground="white",
            relief="flat", cursor="hand2",
            command=self._intentar_login
        )
        self.btn_login.pack(fill="x", ipady=6)

        # Enter activa el login
        self.root.bind("<Return>", lambda e: self._intentar_login())

        # Foco inicial en usuario
        self.entry_usuario.focus_set()

    def _intentar_login(self):
        """Valida credenciales y actualiza la UI según el resultado."""
        usuario  = self.entry_usuario.get().strip()
        password = self.entry_password.get()   # NO hacer .strip() en passwords

        if not usuario or not password:
            self.label_estado.config(text="⚠ Ingresa usuario y contraseña.")
            return

        resultado = verificar_credenciales(usuario, password)

        if resultado is None:
            # Credenciales incorrectas
            restantes = obtener_intentos_restantes(usuario)
            if restantes > 0:
                self.label_estado.config(
                    text=f"❌ Credenciales incorrectas. Intentos restantes: {restantes}"
                )
            else:
                self.label_estado.config(
                    text="🔒 Usuario bloqueado. Contacta al administrador."
                )
            self.entry_password.delete(0, tk.END)  # Limpiar campo password

        elif isinstance(resultado, dict) and resultado.get('bloqueado'):
            self.label_estado.config(
                text="🔒 Usuario bloqueado por múltiples intentos fallidos.\nContacta al administrador."
            )

        else:
            # ✅ Login exitoso
            self.usuario_autenticado = resultado
            self.root.destroy()

    def mostrar(self) -> dict | None:
        """
        Muestra la ventana y espera.
        Retorna datos del usuario autenticado o None.
        """
        self.root.mainloop()
        return self.usuario_autenticado