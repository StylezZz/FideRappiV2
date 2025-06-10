"""
Gestor de configuración para FideRAPPI
Maneja la lectura y escritura del archivo de configuración JSON
"""

import json
import os
from pathlib import Path
from typing import Tuple, Optional, Dict, Any
from tkinter import messagebox

class ConfigManager:
    """Clase para manejar la configuración de la aplicación"""
    
    def __init__(self):
        self.base_dir = self._get_base_dir()
        self.config_dir = self.base_dir / "config"
        self.config_file = self.config_dir / "info.json"
        self._ensure_config_exists()
    
    def _get_base_dir(self) -> Path:
        """Obtiene el directorio base de la aplicación de forma portable"""
        import sys
        if getattr(sys, 'frozen', False):
            # Aplicación compilada
            return Path(sys.executable).parent
        else:
            # Script Python
            return Path(__file__).parent.parent.parent
    
    def _ensure_config_exists(self):
        """Asegura que el directorio y archivo de configuración existan"""
        try:
            self.config_dir.mkdir(exist_ok=True)
            
            if not self.config_file.exists():
                self._create_default_config()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear la configuración: {e}")
    
    def _create_default_config(self):
        """Crea un archivo de configuración por defecto"""
        default_config = {
            "files": {
                "CCE": {
                    "ruta_origen": str(self.base_dir / "templates" / "CCE-Formato.xlsm"),
                    "ruta_destino": str(self.base_dir / "output")
                },
                "CTA_CTES": {
                    "ruta_origen": str(self.base_dir / "templates" / "Corriente-Formato.xlsm"),
                    "ruta_destino": str(self.base_dir / "output")
                },
                "AHORROS": {
                    "ruta_origen": str(self.base_dir / "templates" / "Ahorros-Formato.xlsm"),
                    "ruta_destino": str(self.base_dir / "output")
                },
                "Cargo": {
                    "ruta_origen": str(self.base_dir / "templates" / "Cargo-Formato.xlsx"),
                    "ruta_destino": str(self.base_dir / "output")
                },
                "LBTR": {
                    "ruta_origen": str(self.base_dir / "templates" / "LBTR-Formato.xlsm"),
                    "ruta_destino": str(self.base_dir / "output")
                }
            },
            "lbtr_details": {
                "link": "http://10.7.25.159:9080/LBTR-web/#/login"
            }
        }
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=4, ensure_ascii=False)
    
    def leer_json(self, tipo_operacion: str) -> Tuple[str, str]:
        """
        Lee las rutas de origen y destino para un tipo de operación
        
        Args:
            tipo_operacion: Tipo de operación (CCE, AHORROS, etc.)
        
        Returns:
            Tupla con (ruta_origen, ruta_destino)
        """
        try:
            with open(self.config_file, "r", encoding='utf-8') as archivo_json:
                datos = json.load(archivo_json)
                ruta_origen = datos["files"][tipo_operacion]["ruta_origen"]
                ruta_destino = datos["files"][tipo_operacion]["ruta_destino"]
            return ruta_origen, ruta_destino
        except (FileNotFoundError, KeyError) as fn:
            messagebox.showerror(title="Error", message=f'No se pudo leer el archivo json: {fn}')
            return "", ""
        except json.JSONDecodeError as jde:
            messagebox.showerror(title="Error", message=f'Error al decodificar el archivo JSON: {jde}')
            return "", ""
    
    def lbtr_credenciales(self) -> str:
        """
        Obtiene el enlace de LBTR
        
        Returns:
            URL del sistema LBTR
        """
        try:
            with open(self.config_file, 'r', encoding='utf-8') as file:
                data = json.load(file)
            return data['lbtr_details']['link']
        except (FileNotFoundError, KeyError) as fn:
            messagebox.showerror(title="Error", message=f'No se pudo leer el archivo json: {fn}')
            return ""
        except json.JSONDecodeError as jde:
            messagebox.showerror(title="Error", message=f'Error al decodificar el archivo JSON: {jde}')
            return ""
    
    def modificar_json(self, tipo_operacion: str, nueva_ruta: str) -> bool:
        """
        Modifica la ruta de origen para un tipo de operación
        
        Args:
            tipo_operacion: Tipo de operación
            nueva_ruta: Nueva ruta del archivo
        
        Returns:
            True si se guardó correctamente
        """
        try:
            with open(self.config_file, 'r', encoding='utf-8') as archivo:
                datos = json.load(archivo)
            
            datos["files"][tipo_operacion]["ruta_origen"] = nueva_ruta
            
            with open(self.config_file, 'w', encoding='utf-8') as archivo:
                json.dump(datos, archivo, indent=4, ensure_ascii=False)
                
            messagebox.showinfo(
                message="Se cambió la ruta del archivo origen.",
                title="Guardado exitoso!")
            return True
        except FileNotFoundError:
            messagebox.showerror(title="Error", message="El archivo de configuración no se encontró.")
            return False
        except json.JSONDecodeError:
            messagebox.showerror(title="Error", message="Error al leer el archivo JSON.")
            return False
        except Exception as e:
            messagebox.showerror(title="Error", message=f"Error al guardar: {e}")
            return False
    
    def modificar_ruta_final_json(self, tipo_operacion: str, ruta_destino: str) -> bool:
        """
        Modifica la ruta de destino para un tipo de operación
        
        Args:
            tipo_operacion: Tipo de operación
            ruta_destino: Nueva ruta de destino
        
        Returns:
            True si se guardó correctamente
        """
        try:
            with open(self.config_file, 'r', encoding='utf-8') as archivo:
                datos = json.load(archivo)
            
            datos["files"][tipo_operacion]["ruta_destino"] = ruta_destino
            
            with open(self.config_file, 'w', encoding='utf-8') as archivo:
                json.dump(datos, archivo, indent=4, ensure_ascii=False)
                
            messagebox.showinfo(
                message="Se cambió la ruta donde se guardará el archivo modificado.",
                title="Guardado exitoso!")
            return True
        except Exception as e:
            messagebox.showerror(title="Error", message=f"Error al guardar: {e}")
            return False
    
    def save_link_lbtr(self, link_oficial: str) -> bool:
        """
        Guarda el enlace de LBTR
        
        Args:
            link_oficial: Nueva URL del sistema LBTR
        
        Returns:
            True si se guardó correctamente
        """
        try:
            if "http" not in link_oficial:
                messagebox.showinfo(
                    message="URL no válida. Debe comenzar con http o https.",
                    title="Error de guardado")
                return False
            
            with open(self.config_file, 'r', encoding='utf-8') as archivo:
                datos = json.load(archivo)
                
            datos['lbtr_details']['link'] = link_oficial
                
            with open(self.config_file, 'w', encoding='utf-8') as archivo:
                json.dump(datos, archivo, indent=4, ensure_ascii=False)
                    
            messagebox.showinfo(
                message="Se cambió el enlace de la página LBTR.",
                title="Guardado exitoso!")
            return True
        except Exception as e:
            messagebox.showerror(title="Error", message=f"Error al guardar: {e}")
            return False
    
    def get_config(self) -> Dict[str, Any]:
        """
        Obtiene toda la configuración
        
        Returns:
            Diccionario con toda la configuración
        """
        try:
            with open(self.config_file, 'r', encoding='utf-8') as archivo:
                return json.load(archivo)
        except Exception:
            return {}
    
    def get_base_directory(self) -> Path:
        """Retorna el directorio base de la aplicación"""
        return self.base_dir