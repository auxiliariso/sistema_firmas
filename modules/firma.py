"""
firma.py — Módulo de firma digital
Genera, formatea y valida firmas digitales.
"""

import hashlib
import string
import random
from datetime import datetime
from modules.db import guardar_firma, verificar_firma


# ─────────────────────────────────────────────
# GENERACIÓN DE FIRMA
# ─────────────────────────────────────────────

def generar_id_firma(usuario: str, fecha_hora: str, documento: str) -> str:
    """
    Genera un ID de firma único basado en:
      - nombre de usuario
      - fecha y hora exacta
      - nombre del documento
    
    Proceso:
      1. Concatenar los datos en una cadena semilla
      2. Calcular SHA-256
      3. Tomar los primeros 8 caracteres en MAYÚSCULAS alfanumérico
    
    Retorna un ID corto y único, ej: "A82K91LX"
    """
    semilla = f"{usuario}|{fecha_hora}|{documento}"
    hash_completo = hashlib.sha256(semilla.encode('utf-8')).hexdigest()

    # Convertir hex a alfanumérico legible (letras y números)
    # Tomamos los primeros 8 chars del hash y los convertimos a base 36
    segmento = int(hash_completo[:16], 16)
    chars = string.digits + string.ascii_uppercase
    id_corto = ""
    for _ in range(8):
        id_corto = chars[segmento % 36] + id_corto
        segmento //= 36

    return id_corto


def crear_firma(usuario_data: dict, documento: str) -> dict:
    """
    Crea una firma digital completa para un documento.

    Parámetros:
        usuario_data: dict con 'usuario' y 'nombre' (desde la DB)
        documento: nombre del archivo que se va a firmar

    Retorna un dict con todos los datos de la firma.
    """
    fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M")
    id_firma = generar_id_firma(usuario_data['usuario'], fecha_hora, documento)

    firma = {
        'nombre':     usuario_data['nombre'],
        'usuario':    usuario_data['usuario'],
        'fecha_hora': fecha_hora,
        'id_firma':   id_firma,
        'documento':  documento,
    }

    # Persistir en la base de datos
    guardar_firma(
        usuario=firma['usuario'],
        fecha_hora=firma['fecha_hora'],
        id_firma=firma['id_firma'],
        documento=firma['documento']
    )

    return firma


# ─────────────────────────────────────────────
# FORMATEO
# ─────────────────────────────────────────────

def formatear_firma_texto(firma: dict) -> str:
    """
    Retorna la firma en formato texto listo para insertar en documentos.
    
    Ejemplo de salida:
        ─────────────────────────────
        Firmado por: Juan Pérez
        Fecha:       2026-04-30 16:30
        ID:          A82K91LX
        ─────────────────────────────
    """
    linea = "─" * 34
    return (
        f"{linea}\n"
        f"Firmado por: {firma['nombre']}\n"
        f"Fecha:       {firma['fecha_hora']}\n"
        f"ID:          {firma['id_firma']}\n"
        f"{linea}"
    )


def formatear_firma_dict(firma: dict) -> dict:
    """
    Retorna los campos de la firma como diccionario limpio
    (útil para insertar celda por celda en Excel).
    """
    return {
        'Firmado por': firma['nombre'],
        'Fecha':       firma['fecha_hora'],
        'ID Firma':    firma['id_firma'],
        'Documento':   firma['documento'],
    }


# ─────────────────────────────────────────────
# VALIDACIÓN
# ─────────────────────────────────────────────

def validar_firma(id_firma: str) -> dict:
    """
    Verifica si un ID de firma existe en la base de datos.

    Retorna:
        {'valida': True/False, 'datos': {...} | None, 'mensaje': str}
    """
    registro = verificar_firma(id_firma)

    if registro:
        return {
            'valida':   True,
            'datos':    registro,
            'mensaje':  f"✅ Firma válida — Firmado por {registro['usuario']} el {registro['fecha_hora']}"
        }
    else:
        return {
            'valida':   False,
            'datos':    None,
            'mensaje':  "❌ Firma no encontrada en la base de datos."
        }
