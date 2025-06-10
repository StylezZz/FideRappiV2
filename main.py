#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FideRAPPI - Sistema de Carga de Datos Bancarios
Punto de entrada principal de la aplicación
"""

import sys
import os
import traceback
from pathlib import Path

# Configurar el path para hacer la aplicación portable
if getattr(sys, 'frozen', False):
    # Si está ejecutándose como ejecutable compilado
    BASE_DIR = Path(sys.executable).parent
else:
    # Si está ejecutándose como script Python
    BASE_DIR = Path(__file__).parent

# Añadir el directorio base al path
sys.path.insert(0, str(BASE_DIR))

try:
    from src.interface.main_window import FideRappiApp
    from src.utils.logger import setup_logger
    
    def main():
        """Función principal de la aplicación"""
        try:
            # Configurar logging
            logger = setup_logger()
            logger.info("Iniciando FideRAPPI...")
            
            # Crear y ejecutar la aplicación
            app = FideRappiApp()
            app.title("FideRAPPI - Sistema de Carga de Datos v2.0")
            app.mainloop()
            
        except Exception as e:
            error_msg = f"Error al iniciar la aplicación: {str(e)}\n{traceback.format_exc()}"
            print(error_msg)
            
            # Intentar mostrar error en ventana si es posible
            try:
                import tkinter as tk
                from tkinter import messagebox
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("Error de Aplicación", error_msg)
            except:
                pass
            
            sys.exit(1)
    
    if __name__ == "__main__":
        main()
        
except ImportError as e:
    error_msg = f"Error de importación: {str(e)}\nAsegúrese de que todos los módulos estén instalados correctamente."
    print(error_msg)
    
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error de Importación", error_msg)
    except:
        pass
    
    sys.exit(1)