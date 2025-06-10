from cx_Freeze import setup, Executable
import sys

# Lista de archivos adicionales a incluir
additional_files = [
    ('assets', 'assets'),  # Directorio de recursos (imágenes, íconos, etc.)
    'assets/icons/logo-banco-nacion.ico',
    'msedgedriver.exe'
]

# Configuración base para Windows
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# Opciones de build
build_exe_options = {
    "packages": ["customtkinter", "pandas", "selenium", "PIL", "openpyxl"],
    "include_files": additional_files,
    "excludes": ["tkinter.test"],
    "zip_include_packages": ["encodings", "PySide6"],
    "silent": False,   # Para ver errores durante el build
}

setup(
    name="FideRAPPI",
    version="2.0",
    description="Cargos y abonos automatizados para el área Soporte Mesa de Dinero",
    author="Tu Nombre",
    executables=[
        Executable(
            "main.py",  # Ruta al archivo principal
            base=base,
            icon="assets/icons/logo-banco-nacion.ico",
            target_name="FideRAPPI.exe"
        )
    ],
    options={"build_exe": build_exe_options}
)