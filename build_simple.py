#!/usr/bin/env python3
"""
Script simplificado para construir FideRAPPI.exe
Específicamente diseñado para instalaciones de PyInstaller en directorio de usuario
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

# Ruta específica para tu instalación de PyInstaller
PYINSTALLER_PATH = r"C:\Users\jsantillana\AppData\Local\Programs\Python\Python312\Scripts\pyinstaller.exe"

def check_pyinstaller():
    """Verifica que PyInstaller esté disponible"""
    if os.path.exists(PYINSTALLER_PATH):
        print(f"✅ PyInstaller encontrado en: {PYINSTALLER_PATH}")
        return True
    else:
        print(f"❌ PyInstaller no encontrado en: {PYINSTALLER_PATH}")
        print("💡 Ajusta la variable PYINSTALLER_PATH en el script")
        return False

def check_required_files():
    """Verifica archivos necesarios"""
    files_to_check = [
        'main.py',
        'FideRAPPI.spec', 
        'msedgedriver.exe',
        'src',
        'assets'
    ]
    
    missing = []
    for item in files_to_check:
        if not os.path.exists(item):
            missing.append(item)
    
    if missing:
        print(f"❌ Archivos faltantes: {', '.join(missing)}")
        return False
    
    print("✅ Todos los archivos necesarios están presentes")
    return True

def prepare_directories():
    """Crea directorios necesarios"""
    dirs = ['logs', 'output', 'templates', 'config']
    
    for dir_name in dirs:
        os.makedirs(dir_name, exist_ok=True)
        print(f"📁 {dir_name}")
    
    # Crear .gitkeep en logs
    gitkeep = Path('logs/.gitkeep')
    if not gitkeep.exists():
        gitkeep.touch()

def clean_build():
    """Limpia construcciones anteriores"""
    dirs_to_clean = ['build', 'dist']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"🗑️ Eliminado {dir_name}")

def build_exe():
    """Construye el ejecutable"""
    print("🚀 Construyendo ejecutable...")
    
    cmd = [
        PYINSTALLER_PATH,
        'FideRAPPI.spec',
        '--clean',
        '--noconfirm'
    ]
    
    print(f"🔧 Comando: {' '.join(cmd)}")
    
    try:
        # Ejecutar comando
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=os.getcwd())
        
        if result.returncode == 0:
            print("✅ Construcción exitosa")
            return True
        else:
            print("❌ Error en la construcción:")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            return False
            
    except Exception as e:
        print(f"❌ Error ejecutando PyInstaller: {e}")
        return False

def verify_result():
    """Verifica que el ejecutable se creó correctamente"""
    exe_path = Path('dist/FideRAPPI.exe')
    
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"✅ Ejecutable creado: {exe_path}")
        print(f"📊 Tamaño: {size_mb:.1f} MB")
        return True
    else:
        print(f"❌ Ejecutable no encontrado en {exe_path}")
        return False

def post_build():
    """Configuración post-construcción"""
    dist_dir = Path('dist')
    
    # Crear directorios en dist
    for dir_name in ['templates', 'output', 'logs', 'config']:
        (dist_dir / dir_name).mkdir(exist_ok=True)
    
    # Verificar msedgedriver.exe
    dist_driver = dist_dir / 'msedgedriver.exe'
    if not dist_driver.exists():
        source_driver = Path('msedgedriver.exe')
        if source_driver.exists():
            shutil.copy2(source_driver, dist_driver)
            print(f"📋 Copiado: msedgedriver.exe → dist/")
        else:
            print("⚠️ msedgedriver.exe no encontrado")
    
    print("✅ Configuración completada")

def main():
    print("🎯 FideRAPPI - Constructor Simplificado")
    print("=" * 45)
    
    # Verificaciones
    if not check_pyinstaller():
        print("\n💡 SOLUCIÓN MANUAL:")
        print("1. Verifica la ruta de PyInstaller")
        print("2. O ejecuta directamente:")
        print(f'   {PYINSTALLER_PATH} FideRAPPI.spec --clean --noconfirm')
        return False
    
    if not check_required_files():
        return False
    
    # Proceso de construcción
    print("\n🔧 Preparando construcción...")
    prepare_directories()
    clean_build()
    
    print("\n🏗️ Construyendo...")
    if not build_exe():
        return False
    
    print("\n🔍 Verificando resultado...")
    if not verify_result():
        return False
    
    print("\n⚙️ Configuración final...")
    post_build()
    
    print("\n🎉 ¡CONSTRUCCIÓN COMPLETADA!")
    print("📦 Ejecutable en: dist/FideRAPPI.exe")
    print("\n💡 Para probar:")
    print("   cd dist")
    print("   .\\FideRAPPI.exe")
    
    return True

if __name__ == "__main__":
    success = main()
    if not success:
        print("\n❌ La construcción falló")
        input("Presiona Enter para continuar...")
    else:
        print("\n✅ Todo listo!")
        input("Presiona Enter para continuar...")
    sys.exit(0 if success else 1)