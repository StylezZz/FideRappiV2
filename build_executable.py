#!/usr/bin/env python3
"""
Script para construir el ejecutable de FideRAPPI usando PyInstaller
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def check_dependencies():
    """Verifica que las dependencias necesarias estén instaladas"""
    required_packages = [
        'pyinstaller',
        'customtkinter',
        'xlwings', 
        'pandas',
        'selenium',
        'pyautogui',
        'keyboard',
        'PIL',
        'openpyxl'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"❌ Paquetes faltantes: {', '.join(missing_packages)}")
        print("Instálalos con: pip install " + " ".join(missing_packages))
        return False
    
    print("✅ Todas las dependencias están instaladas")
    return True

def check_required_files():
    """Verifica que los archivos necesarios existan"""
    required_files = [
        'main.py',
        'FideRAPPI.spec',
        'msedgedriver.exe'
    ]
    
    required_dirs = [
        'src',
        'assets'
    ]
    
    missing_files = []
    
    # Verificar archivos
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    # Verificar directorios
    for dir in required_dirs:
        if not os.path.exists(dir):
            missing_files.append(f"{dir}/ (directorio)")
    
    if missing_files:
        print(f"❌ Archivos/directorios faltantes: {', '.join(missing_files)}")
        return False
    
    print("✅ Todos los archivos necesarios están presentes")
    return True

def prepare_build_environment():
    """Prepara el entorno para la construcción"""
    print("🔧 Preparando entorno de construcción...")
    
    # Crear directorios necesarios
    directories = ['logs', 'output', 'templates', 'config']
    
    for dir_name in directories:
        os.makedirs(dir_name, exist_ok=True)
        print(f"   📁 Directorio {dir_name} listo")
    
    # Crear archivo .gitkeep en logs si no existe
    gitkeep_path = Path('logs/.gitkeep')
    if not gitkeep_path.exists():
        gitkeep_path.touch()
    
    # Verificar que assets/icons tiene el icono
    icon_path = Path('assets/icons/logo-banco-nacion.ico')
    if not icon_path.exists():
        print(f"⚠️  Icono no encontrado en {icon_path}")
        # Crear directorio de iconos si no existe
        icon_path.parent.mkdir(parents=True, exist_ok=True)
    
    print("✅ Entorno preparado")

def clean_previous_builds():
    """Limpia construcciones anteriores"""
    print("🧹 Limpiando construcciones anteriores...")
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"   🗑️  Eliminado {dir_name}")
    
    # Limpiar archivos .pyc recursivamente
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.endswith('.pyc'):
                os.remove(os.path.join(root, file))
    
    print("✅ Limpieza completada")

def build_executable():
    """Construye el ejecutable usando PyInstaller"""
    print("🚀 Iniciando construcción del ejecutable...")
    
    # Comando de PyInstaller
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        'FideRAPPI.spec',
        '--clean',
        '--noconfirm'
    ]
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ Construcción completada exitosamente")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Error durante la construcción:")
        print(f"Código de salida: {e.returncode}")
        print(f"Stdout: {e.stdout}")
        print(f"Stderr: {e.stderr}")
        return False

def verify_executable():
    """Verifica que el ejecutable se haya creado correctamente"""
    exe_path = Path('dist/FideRAPPI.exe')
    
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"✅ Ejecutable creado: {exe_path}")
        print(f"   📊 Tamaño: {size_mb:.1f} MB")
        return True
    else:
        print(f"❌ Ejecutable no encontrado en {exe_path}")
        return False

def post_build_setup():
    """Configuración post-construcción"""
    print("🔧 Configuración post-construcción...")
    
    dist_dir = Path('dist')
    
    # Crear directorios necesarios en dist
    necessary_dirs = ['templates', 'output', 'logs', 'config']
    
    for dir_name in necessary_dirs:
        target_dir = dist_dir / dir_name
        target_dir.mkdir(exist_ok=True)
        print(f"   📁 Creado {target_dir}")
    
    # Copiar archivos de configuración si existen
    config_source = Path('config/info.json')
    config_target = dist_dir / 'config/info.json'
    
    if config_source.exists():
        shutil.copy2(config_source, config_target)
        print(f"   📄 Copiado {config_source} → {config_target}")
    
    # Verificar que msedgedriver.exe esté en dist
    driver_path = dist_dir / 'msedgedriver.exe'
    if not driver_path.exists():
        source_driver = Path('msedgedriver.exe')
        if source_driver.exists():
            shutil.copy2(source_driver, driver_path)
            print(f"   🚗 Copiado driver: {source_driver} → {driver_path}")
        else:
            print(f"   ⚠️  Driver no encontrado: {source_driver}")
    
    print("✅ Configuración post-construcción completada")

def main():
    """Función principal del script de construcción"""
    print("🎯 FideRAPPI - Constructor de Ejecutable")
    print("=" * 50)
    
    # Verificaciones previas
    if not check_dependencies():
        return False
    
    if not check_required_files():
        return False
    
    # Preparar entorno
    prepare_build_environment()
    
    # Limpiar construcciones anteriores
    clean_previous_builds()
    
    # Construir ejecutable
    if not build_executable():
        return False
    
    # Verificar resultado
    if not verify_executable():
        return False
    
    # Configuración final
    post_build_setup()
    
    print("\n🎉 ¡Construcción completada exitosamente!")
    print(f"📦 Ejecutable disponible en: dist/FideRAPPI.exe")
    print("\n💡 Consejos:")
    print("   • Asegúrate de que msedgedriver.exe esté en el mismo directorio que el ejecutable")
    print("   • Los templates deben estar en la carpeta templates/")
    print("   • Los archivos procesados se guardarán en output/")
    
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)