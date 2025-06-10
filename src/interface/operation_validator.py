"""
Validador de operaciones para FideRAPPI
Valida ventanas del host y ejecuta operaciones correspondientes
"""

import threading
import pyautogui
import pyperclip
from tkinter import messagebox, filedialog
import os
from typing import Optional, List

from src.utils.logger import LoggerMixin
from src.utils.config_manager import ConfigManager


class OperationValidator(LoggerMixin):
    """Clase para validar y ejecutar operaciones en el host bancario"""
    
    def __init__(self, tipo_operacion: str):
        """
        Inicializa el validador de operaciones
        
        Args:
            tipo_operacion: Tipo de operación (CCE, AHORROS, etc.)
        """
        self.tipo_operacion = tipo_operacion
        self.config_manager = ConfigManager()
        self.ventanas = []
        self.cant_ventanas = 0
        self.ventana_produccion = None
        self.cargo_activo = False
        
        # Mapeo de operaciones a códigos y clases
        self.operaciones_dict = {
            "CCE": {
                "codigo_abono": "SICA6033",
                "codigo_cargo": "SITB6062",
                "clase": "cce_operations",
                "metodo_abono": "execute_cce",
                "metodo_cargo": "execute_cargo_cce"
            },
            "CTA_CTES": {
                "codigo_abono": "SITB6061", 
                "codigo_cargo": "SITB6062",
                "clase": "cte_operations",
                "metodo_abono": "execute_ctas_ctes",
                "metodo_cargo": "execute_ctas_ctes"
            },
            "AHORROS": {
                "codigo_abono": "SITB6010",
                "codigo_cargo": "SITB6062", 
                "clase": "ahorros_operations",
                "metodo_abono": "execute_ahorros",
                "metodo_cargo": "execute_cargo_ahorros"
            },
            "LBTR": {
                "codigo_cargo": "SITB6062",
                "clase": "lbtr_operations", 
                "metodo_cargo": "execute_cargo_lbtr"
            },
            "Cargo": {
                "codigo_abono": "SITB6062",
                "clase": "cargo_operations",
                "metodo_abono": "execute_carga"
            }
        }
    
    def validar_y_ejecutar_operacion(self, es_cargo: bool = False) -> bool:
        """
        Valida ventanas de producción y ejecuta la operación correspondiente
        
        Args:
            es_cargo: True si es operación de cargo, False si es abono
        
        Returns:
            True si se ejecutó correctamente
        """
        try:
            self.cargo_activo = es_cargo
            
            # Buscar ventanas de producción
            self.ventanas = pyautogui.getWindowsWithTitle("prod")
            self.cant_ventanas = len(self.ventanas)
            
            self.logger.info(f"Encontradas {self.cant_ventanas} ventanas de producción")
            
            if self.cant_ventanas == 1:
                # Solo una ventana, usarla directamente
                return self._validar_ventana(self.ventanas[0])
            elif self.cant_ventanas > 1:
                # Múltiples ventanas, permitir selección
                return self._seleccionar_ventana()
            else:
                messagebox.showinfo(
                    title="Notificación",
                    message="No se encontraron ventanas de producción abiertas."
                )
                return False
                
        except Exception as e:
            self.logger.error(f"Error en validación de operación: {e}")
            messagebox.showerror(
                title="Problema detectado",
                message=f'Error en la detección de ventanas de host: {e}'
            )
            return False
    
    def _seleccionar_ventana(self) -> bool:
        """
        Permite al usuario seleccionar una ventana de producción
        
        Returns:
            True si se seleccionó y validó correctamente
        """
        try:
            from tkinter import Toplevel, ttk, Button
            import customtkinter
            
            # Crear ventana de selección
            ventana_seleccion = customtkinter.CTkToplevel()
            ventana_seleccion.title("Seleccionar ventana")
            ventana_seleccion.geometry("500x120")
            ventana_seleccion.transient()
            ventana_seleccion.grab_set()
            
            resultado = [False]  # Lista para poder modificar desde funciones internas
            
            # Lista de títulos de ventanas
            titulos_ventanas = [ventana.title for ventana in self.ventanas if ventana.title]
            
            # ComboBox para selección
            cmb_ventanas = customtkinter.CTkComboBox(
                ventana_seleccion,
                state="readonly", 
                values=titulos_ventanas,
                width=300
            )
            cmb_ventanas.set("--Seleccione--")
            cmb_ventanas.pack(padx=20, pady=20)
            
            def elegir_ventana():
                seleccion = cmb_ventanas.get()
                if seleccion == "--Seleccione--":
                    messagebox.showinfo("Selección", "Por favor seleccione una opción.")
                    return
                
                # Encontrar la ventana seleccionada
                self.ventana_produccion = next(
                    (v for v in self.ventanas if v.title == seleccion), 
                    None
                )
                
                if self.ventana_produccion:
                    ventana_seleccion.destroy()
                    resultado[0] = self._validar_ventana(self.ventana_produccion)
                else:
                    messagebox.showinfo("Selección", "Opción inválida.")
            
            def cancelar():
                ventana_seleccion.destroy()
                resultado[0] = False
            
            # Botones
            frame_botones = customtkinter.CTkFrame(ventana_seleccion)
            frame_botones.pack(pady=10)
            
            btn_elegir = customtkinter.CTkButton(
                frame_botones, 
                text="Elegir",
                command=elegir_ventana
            )
            btn_elegir.pack(side="left", padx=10)
            
            btn_cancelar = customtkinter.CTkButton(
                frame_botones,
                text="Cancelar", 
                command=cancelar
            )
            btn_cancelar.pack(side="left", padx=10)
            
            # Esperar a que se cierre la ventana
            ventana_seleccion.wait_window()
            
            return resultado[0]
            
        except Exception as e:
            self.logger.error(f"Error en selección de ventana: {e}")
            messagebox.showerror(
                "Problema detectado",
                f'Error en la selección de ventanas: {e}'
            )
            return False
    
    def _validar_ventana(self, ventana) -> bool:
        """
        Valida una ventana específica y ejecuta la operación
        
        Args:
            ventana: Ventana a validar
        
        Returns:
            True si se validó y ejecutó correctamente
        """
        try:
            # Activar y maximizar ventana
            ventana.maximize()
            ventana.activate()
            pyautogui.sleep(0.5)
            
            # Copiar contenido del menú
            pyautogui.hotkey('ctrl', 'c')
            pyautogui.sleep(0.3)
            contenido_menu = pyperclip.paste()
            lineas_menu = contenido_menu.splitlines()
            
            self.logger.info(f"Validando ventana para {self.tipo_operacion}")
            
            # Verificar operación válida
            if self.tipo_operacion not in self.operaciones_dict:
                raise ValueError(f"Operación no reconocida: {self.tipo_operacion}")
            
            config_operacion = self.operaciones_dict[self.tipo_operacion]
            
            # Determinar código a buscar según tipo de operación
            if self.cargo_activo:
                codigo_buscar = config_operacion.get("codigo_cargo")
                metodo_ejecutar = config_operacion.get("metodo_cargo")
            else:
                codigo_buscar = config_operacion.get("codigo_abono")
                metodo_ejecutar = config_operacion.get("metodo_abono")
            
            if not codigo_buscar or not metodo_ejecutar:
                raise ValueError(f"Configuración incompleta para {self.tipo_operacion}")
            
            # Verificar que el código existe en el menú
            codigo_encontrado = any(codigo_buscar in linea for linea in lineas_menu)
            if not codigo_encontrado:
                raise ValueError(f"No se encuentra el código {codigo_buscar} en la ventana")
            
            # Importar y ejecutar operación
            return self._ejecutar_operacion(config_operacion, metodo_ejecutar, ventana)
            
        except Exception as e:
            self.logger.error(f"Error validando ventana: {e}")
            messagebox.showerror("ERROR", f'Error en la verificación de ventanas: {e}')
            return False
    
    def _ejecutar_operacion(self, config_operacion: dict, metodo_ejecutar: str, ventana) -> bool:
        """
        Ejecuta la operación correspondiente
        
        Args:
            config_operacion: Configuración de la operación
            metodo_ejecutar: Nombre del método a ejecutar
            ventana: Ventana del emulador
        
        Returns:
            True si se ejecutó correctamente
        """
        try:
            # Importar la clase correspondiente
            clase_nombre = config_operacion["clase"]
            
            if clase_nombre == "cce_operations":
                from src.operations.cce_operations import CCEOperations
                operacion = CCEOperations()
            elif clase_nombre == "ahorros_operations":
                from src.operations.ahorros_operations import AhorrosOperations
                operacion = AhorrosOperations()
            elif clase_nombre == "cte_operations":
                from src.operations.cte_operations import CTEOperations
                operacion = CTEOperations()
            elif clase_nombre == "lbtr_operations":
                from src.operations.lbtr_operations import LBTROperations
                operacion = LBTROperations()
            elif clase_nombre == "cargo_operations":
                from src.operations.cargo_operations import CargoOperations
                operacion = CargoOperations()
            else:
                raise ValueError(f"Clase no reconocida: {clase_nombre}")
            
            # Obtener método a ejecutar
            metodo = getattr(operacion, metodo_ejecutar)
            
            # Configurar argumentos según el tipo de operación
            if self.cargo_activo and "cargo" in metodo_ejecutar:
                # Para operaciones de cargo, solicitar archivo
                ruta_xlc = self._solicitar_archivo_historial()
                if not ruta_xlc:
                    return False
                
                # Ejecutar en hilo separado
                self._ejecutar_con_hilos(operacion, metodo, ventana, ruta_xlc)
            else:
                # Para operaciones normales
                self._ejecutar_con_hilos(operacion, metodo, ventana)
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error ejecutando operación: {e}")
            messagebox.showerror("ERROR", f'Error ejecutando operación: {e}')
            return False
    
    def _solicitar_archivo_historial(self) -> Optional[str]:
        """
        Solicita al usuario seleccionar un archivo de historial
        
        Returns:
            Ruta del archivo seleccionado o None si se canceló
        """
        try:
            # Obtener directorio de procesados
            ruta_origen, _ = self.config_manager.leer_json(self.tipo_operacion)
            if ruta_origen:
                directorio_base = os.path.dirname(ruta_origen)
                directorio_inicial = os.path.join(directorio_base, "Procesados")
            else:
                directorio_inicial = None
            
            # Solicitar archivo
            ruta_archivo = filedialog.askopenfilename(
                title="Seleccionar archivo del historial",
                initialdir=directorio_inicial,
                filetypes=[
                    ("Archivos Excel", "*.xlsx"),
                    ("Archivos Excel con macros", "*.xlsm"),
                    ("Todos los archivos", "*.*")
                ]
            )
            
            if ruta_archivo:
                self.logger.info(f"Archivo de historial seleccionado: {ruta_archivo}")
            else:
                self.logger.info("Selección de archivo cancelada")
            
            return ruta_archivo if ruta_archivo else None
            
        except Exception as e:
            self.logger.error(f"Error solicitando archivo: {e}")
            return None
    
    def _ejecutar_con_hilos(self, operacion, metodo, ventana, *args):
        """
        Ejecuta la operación en hilos separados
        
        Args:
            operacion: Instancia de la operación
            metodo: Método a ejecutar
            ventana: Ventana del emulador
            *args: Argumentos adicionales
        """
        try:
            # Crear hilos
            hilos = []
            
            # Hilo para detección de botones (si existe el método)
            if hasattr(operacion, 'detectar_botones'):
                hilo_botones = threading.Thread(target=operacion.detectar_botones)
                hilos.append(hilo_botones)
            
            # Hilo para la operación principal
            if args:
                hilo_operacion = threading.Thread(target=metodo, args=(ventana, *args))
            else:
                hilo_operacion = threading.Thread(target=metodo, args=(ventana,))
            hilos.append(hilo_operacion)
            
            # Iniciar hilos
            for hilo in hilos:
                hilo.start()
            
            self.logger.info(f"Hilos iniciados para {self.tipo_operacion}")
            
        except Exception as e:
            self.logger.error(f"Error ejecutando con hilos: {e}")
            raise