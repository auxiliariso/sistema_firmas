# 🔏 Sistema de Firma Digital Interna

Prototipo funcional de firma digital para documentos Word y Excel.
Funciona completamente **offline**, con base de datos local SQLite.

---

## 📁 Estructura del proyecto

```
firma_digital/
│
├── main.py                  ← Punto de entrada principal
├── demo_documentos.py       ← Demo de integración con archivos
├── firma_digital.db         ← Base de datos SQLite (se crea automáticamente)
│
├── modules/
│   ├── db.py                ← Operaciones de base de datos
│   ├── firma.py             ← Generación y validación de firmas
│   ├── login.py             ← Ventana de autenticación (tkinter)
│   ├── app.py               ← Ventana principal (tkinter)
│   └── utils.py             ← Integración con Excel y Word
│
└── output_examples/         ← Documentos de ejemplo generados
    ├── reporte_ventas_firmado.xlsx
    └── contrato_servicio_firmado.docx
```

---

## ⚙️ Instalación

### 1. Requisitos
- Python 3.10 o superior
- tkinter (incluido en Python estándar)

### 2. Instalar dependencias

```bash
pip install openpyxl python-docx
```

### 3. Ejecutar

```bash
python main.py
```

---

## 👤 Usuarios de demostración

| Usuario   | Contraseña    | Nombre            |
|-----------|---------------|-------------------|
| jperez    | Firma2026!    | Juan Pérez        |
| mgarcia   | Seguro#123    | María García      |
| admin     | Admin@999     | Administrador     |

> ⚠️ Las contraseñas se almacenan con **PBKDF2-SHA256 + salt aleatorio**.
> Nunca se guardan ni muestran en texto plano.

---

## 🔐 Seguridad implementada

| Característica                     | Detalle                                        |
|------------------------------------|------------------------------------------------|
| Hash de contraseñas                | PBKDF2-SHA256, 260,000 iteraciones             |
| Salt aleatorio por usuario         | 16 bytes (os.urandom), único por usuario       |
| Campo password oculto              | `show="•"` en el widget Entry de tkinter       |
| Límite de intentos                 | Bloqueo automático tras 3 intentos fallidos    |
| Contraseña nunca en memoria extra  | No se almacena en variables globales/logs      |
| ID de firma único                  | SHA-256 de usuario+fecha+documento (8 chars)   |

---

## ✍️ Cómo funciona la firma

1. El usuario inicia sesión con sus credenciales
2. Selecciona o escribe el nombre del documento a firmar
3. El sistema genera:
   - Un **ID único** basado en SHA-256 (usuario + fecha + documento)
   - Un registro en la tabla `log_firmas` de la base de datos
4. La firma se puede **insertar en Excel o Word** con un clic
5. Cualquier firma puede **verificarse** introduciendo su ID

### Formato de firma

```
──────────────────────────────────
Firmado por: Juan Pérez
Fecha:       2026-04-30 16:30
ID:          A82K91LX
──────────────────────────────────
```

---

## 📊 Base de datos

### Tabla `usuarios`

| Columna          | Tipo    | Descripción                          |
|------------------|---------|--------------------------------------|
| usuario          | TEXT PK | Nombre de usuario único              |
| password_hash    | TEXT    | Hash PBKDF2-SHA256                   |
| salt             | TEXT    | Salt aleatorio hex (32 chars)        |
| nombre           | TEXT    | Nombre completo                      |
| bloqueado        | BOOLEAN | 1 = bloqueado por intentos fallidos  |
| intentos_fallidos| INTEGER | Contador (reset al login exitoso)    |

### Tabla `log_firmas`

| Columna   | Tipo    | Descripción                    |
|-----------|---------|--------------------------------|
| id        | INTEGER | Auto-incremental               |
| usuario   | TEXT    | Quién firmó                    |
| fecha_hora| TEXT    | Cuándo (YYYY-MM-DD HH:MM)      |
| id_firma  | TEXT    | ID único de la firma (8 chars) |
| documento | TEXT    | Nombre del documento firmado   |

---

## 🚀 Convertir a ejecutable .exe (Windows)

### Paso 1: Instalar PyInstaller
```bash
pip install pyinstaller
```

### Paso 2: Generar el .exe
```bash
# Opción A: Un solo archivo ejecutable (recomendado para distribución)
pyinstaller --onefile --windowed --name FirmaDigital main.py

# Opción B: Carpeta con dependencias (arranque más rápido)
pyinstaller --onedir --windowed --name FirmaDigital main.py
```

### Paso 3: Encontrar el ejecutable
El archivo generado estará en:
```
dist/FirmaDigital.exe        ← Opción A
dist/FirmaDigital/           ← Opción B
```

### Notas para el .exe
- La base de datos `firma_digital.db` se crea automáticamente junto al .exe
- Si usas `--onefile`, PyInstaller descomprime en `%TEMP%` al ejecutar
- Para que la DB persista junto al .exe, el código ya usa `os.path.dirname(os.path.abspath(__file__))` correctamente

---

## 🧪 Demo de integración con documentos

```bash
python demo_documentos.py
```

Genera en `output_examples/`:
- `reporte_ventas_firmado.xlsx` — Excel con tabla de ventas + firma al final
- `contrato_servicio_firmado.docx` — Contrato Word + firma al final

---

## ➕ Crear nuevos usuarios desde código

```python
from modules.db import inicializar_db, crear_usuario

inicializar_db()
crear_usuario("nuevo_user", "MiPasswordSegura123!", "Nombre Completo")
```

## 🔍 Verificar una firma desde código

```python
from modules.firma import validar_firma

resultado = validar_firma("A82K91LX")
if resultado['valida']:
    print("Firma válida:", resultado['datos'])
else:
    print("Firma no encontrada")
```
