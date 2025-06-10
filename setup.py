"""
Setup para crear ejecutable de FideRAPPI
Utiliza cx_Freeze para generar un ejecutable portable
"""

import sys
import os
from cx_Freeze import setup, Executable

# Determinar el directorio base
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

# Lista de archivos adicionales a incluir
additional_files = [
    # Archivos de configuración
    (os.path.join(base_dir, 'config'), 'config'),
    
    # Directorio de assets (imágenes, iconos)
    (os.path.join(base_dir, 'assets'), 'assets'),
    
    # Templates de Excel (si existen)
    (os.path.join(base_dir, 'templates'), 'templates'),
    
    # Icono de la aplicación
    'logo-banco-nacion.ico',
    
    # Driver de Edge para Selenium
    'msedgedriver.exe',
    
    # Archivos de documentación
    'README.md',
    'CHANGELOG.md',
]

# Filtrar archivos que realmente existen
existing_files = []
for item in additional_files:
    if isinstance(item, tuple):
        source, target = item
        if os.path.exists(source):
            existing_files.append((source, target))
    else:
        if os.path.exists(item):
            existing_files.append(item)

# Paquetes que deben incluirse explícitamente
packages = [
    "customtkinter",
    "pandas", 
    "xlwings",
    "pyautogui",
    "pyperclip",
    "keyboard",
    "PIL",
    "selenium",
    "PyPDF2",
    "pathlib",
    "datetime",
    "threading",
    "tkinter",
    "json",
    "os",
    "sys",
    "re",
    "time",
    "traceback"
]

# Módulos que deben excluirse para reducir tamaño
excludes = [
    "tkinter.test",
    "unittest",
    "test",
    "tests",
    "pytest",
    "numpy.tests",
    "pandas.tests",
    "matplotlib",
    "scipy"
]

# Configuración base para Windows
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Evita mostrar consola en Windows

# Opciones de build
build_exe_options = {
    "packages": packages,
    "include_files": existing_files,
    "excludes": excludes,
    "zip_include_packages": ["encodings", "importlib"],
    "optimize": 2,  # Optimización máxima
    "include_msvcrt": True,  # Incluir runtime de Visual C++
}

# Información del ejecutable
executable = Executable(
    script="main.py",
    base=base,
    icon="logo-banco-nacion.ico" if os.path.exists("logo-banco-nacion.ico") else None,
    target_name="FideRAPPI.exe",
    shortcut_name="FideRAPPI",
    shortcut_dir="DesktopFolder",
)

# Configuración del setup
setup(
    name="FideRAPPI",
    version="2.0.0",
    description="Sistema de Carga de Datos Bancarios - Automatización para operaciones de fideicomiso",
    author="FideRAPPI Team",
    author_email="",
    url="",
    long_description="""
    FideRAPPI es una aplicación de escritorio para automatizar operaciones bancarias
    relacionadas con fideicomisos. Incluye funcionalidades para:
    
    - Operaciones CCE (Cámara de Compensación Electrónica)
    - Cuentas de Ahorros
    - Cuentas Corrientes  
    - Transferencias LBTR (Sistema de Liquidación Bruta en Tiempo Real)
    - Operaciones de Cargo
    - Utilidades adicionales (unión de PDFs)
    
    La aplicación está diseñada para ser portable y fácil de distribuir.
    """,
    executables=[executable],
    options={"build_exe": build_exe_options}
)

# Instrucciones post-instalación
print("""
=================================================================
                    FIDERAPI - SETUP COMPLETADO
=================================================================

Para crear el ejecutable, ejecute:
    python setup.py build

El ejecutable se creará en el directorio 'build/'

Archivos necesarios adicionales:
- msedgedriver.exe (para operaciones LBTR)
- Plantillas de Excel en 'templates/'
- Archivos de configuración en 'config/'

=================================================================
""")

# Script alternativo usando auto-py-to-exe (comentado)
"""
Alternativa usando auto-py-to-exe:

1. Instalar: pip install auto-py-to-exe
2. Ejecutar: auto-py-to-exe
3. Configurar en la interfaz gráfica:
   - Script Location: main.py
   - Onefile: No (para mejor rendimiento)
   - Console Window: No
   - Icon: logo-banco-nacion.ico
   - Additional Files: agregar carpetas config/, assets/, templates/
   - Hidden Imports: customtkinter, xlwings, selenium

Comando directo con pyinstaller:
pyinstaller --name="FideRAPPI" --windowed --icon=logo-banco-nacion.ico 
--add-data="config;config" --add-data="assets;assets" 
--add-data="templates;templates" --add-binary="msedgedriver.exe;." 
--hidden-import=customtkinter --hidden-import=xlwings 
--hidden-import=selenium main.py
"""