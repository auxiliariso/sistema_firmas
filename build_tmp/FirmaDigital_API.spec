# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\HP\\Downloads\\firma_documentos\\api.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\HP\\Downloads\\firma_documentos\\modules', 'modules')],
    hiddenimports=['modules.db', 'modules.firma', 'modules.login', 'modules.app', 'modules.utils', 'openpyxl', 'openpyxl.styles', 'openpyxl.utils', 'openpyxl.utils.cell', 'openpyxl.writer.excel', 'openpyxl.reader.excel', 'docx', 'docx.shared', 'docx.enum.text', 'docx.oxml', 'docx.oxml.ns', 'docx.oxml.shared', 'flask', 'flask.json', 'werkzeug', 'werkzeug.serving', 'jinja2', 'click', 'itsdangerous', 'sqlite3', 'hashlib', 'secrets', 'getpass', 'tkinter', 'tkinter.ttk', 'tkinter.messagebox', 'tkinter.filedialog', '_tkinter', 'email.mime.text', 'email.mime.multipart'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FirmaDigital_API',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
