"""
Operaciones de Cargo
Maneja cargos individuales a cuentas
"""

import time
import os
import pandas as pd
import pyautogui
import pyperclip
import xlwings as xw
from tkinter import messagebox
import keyboard
import datetime
from typing import Optional

from src.core.base_logic import BaseLogic
from src.utils.config_manager import ConfigManager
from src.utils.file_manager import FileManager


class CargoOperations(BaseLogic):
    """Clase para manejar operaciones de Cargo individual"""
    
    def __init__(self):
        super().__init__("Cargo")
        self.config_manager = ConfigManager()
        self.file_manager = FileManager()
        
        # Definir tipos de datos para las columnas
        self.dicc_tabla = {
            'Id': str,
            'COD': str,
            'Cuenta': str,
            'Importe': float,
            'Memorandum': str,
            'Motivo': str,
            'Glosa1': str,
            'Glosa2': str,
            'Glosa3': str,
            'Mensaje_emulacion': str,
            'Observacion': str
        }
    
    def detectar_botones(self):
        """Detecta la pulsación de teclas para detener el proceso"""
        def on_key_event(event):
            if event.event_type == keyboard.KEY_DOWN and (
                keyboard.is_pressed('alt gr') or event.name == 'esc'
            ):
                self.detener_proceso = True
        
        try:
            self.detener_proceso = False
            self.deteccion_activa = True
            self.logger.info("Iniciando detección de botones ESC para Cargo")
            
            keyboard.hook(on_key_event)
            while self.deteccion_activa and not self.detener_proceso:
                time.sleep(0.1)
                
            self.logger.info("Detección de botones finalizada")
        except Exception as e:
            self.logger.error(f"Error en detección de botones: {e}")
            messagebox.showerror("ERROR", f"Problemas con la detección de botones: {e}")
        finally:
            keyboard.unhook(on_key_event)
    
    def execute_carga(self, ventana) -> bool:
        """
        Ejecuta el proceso de carga individual
        
        Args:
            ventana: Ventana del emulador bancario
        
        Returns:
            True si se completó correctamente
        """
        wb = None
        book = None
        ruta_procesado = ''
        finalizado = False
        
        try:
            self.iniciar_operacion()
            
            # Obtener configuración
            ruta_origen, ruta_destino = self.config_manager.leer_json("Cargo")
            if not ruta_origen:
                messagebox.showerror("Error", "No se pudo obtener la configuración de Cargo")
                return False
            
            directorio = os.path.dirname(ruta_origen)
            
            # Abrir Excel
            book = xw.App(visible=False)
            wb = book.books.open(ruta_origen)
            hoja = wb.sheets['Cargo']
            
            # Leer datos del Excel
            tabla_cargo = pd.read_excel(ruta_origen, sheet_name='Cargo', header=0, dtype=self.dicc_tabla)
            
            fecha_actual = self.get_fecha_actual()
            lista_memo_cce = set()
            cont_cargados = 0
            cont_no_cargados = 0
            
            self.logger.info(f"Procesando {len(tabla_cargo)} registros de Cargo")
            
            # Procesar cada fila
            for indice, fila in tabla_cargo.iterrows():
                if self.detener_proceso:
                    break
                
                fila_df = indice + 2
                
                # Extraer datos de la fila
                cod = str(fila['COD']).strip()
                cuenta = self.limpiar_numero_cuenta(str(fila['Cuenta']))
                importe = float(fila['Importe'])
                memo = str(fila['Memorandum']).strip()
                motivo = str(fila['Motivo']).strip()
                glosa1 = str(fila['Glosa1']).strip()
                glosa2 = str(fila['Glosa2']).strip() if not pd.isna(fila['Glosa2']) else ""
                glosa3 = str(fila['Glosa3']).strip() if not pd.isna(fila['Glosa3']) else ""
                obs = fila['Observacion']
                
                self.logger.info(f"Procesando fila {fila_df} - Memo: {memo}, Importe: {importe}")
                
                # Agregar memo a la lista
                if indice == 0:
                    lista_memo_cce.add(memo)
                if memo not in lista_memo_cce:
                    lista_memo_cce.add(memo)
                
                # Verificar si ya está procesado
                if not pd.isna(obs):
                    self.logger.info(f"Fila {fila_df} ya tiene observación: {obs}")
                    continue
                
                # Procesar cargo
                resultado = self._procesar_cargo_individual(
                    ventana, hoja, fila_df, cuenta, importe, memo, 
                    motivo, glosa1, glosa2, glosa3
                )
                
                if resultado:
                    cont_cargados += 1
                else:
                    cont_no_cargados += 1
                
                wb.save()
            
            # Finalizar proceso
            self.finalizar_operacion()
            
            if self.detener_proceso:
                messagebox.showwarning(
                    "Proceso detenido",
                    "Se ha procedido a detener todos los procesos."
                )
            else:
                # Guardar archivo procesado
                ruta_procesado = self._guardar_archivo_procesado(
                    wb, lista_memo_cce, directorio, fecha_actual
                )
                finalizado = True
                
                messagebox.showinfo(
                    "Proceso finalizado",
                    f"Cargos realizados = {cont_cargados}\n"
                    f"Cargos no realizados = {cont_no_cargados}"
                )
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error en ejecución Cargo: {e}")
            messagebox.showerror("Error de ejecución", 
                               f"No se ha podido completar la ejecución: {e}")
            return False
        finally:
            self.finalizar_operacion()
            if wb and book:
                try:
                    wb.save()
                    wb.close()
                    book.quit()
                except Exception as e:
                    self.logger.warning(f"Error cerrando Excel: {e}")
            
            # Abrir archivo procesado
            if ruta_procesado and finalizado:
                try:
                    os.startfile(ruta_procesado)
                except Exception as e:
                    self.logger.warning(f"No se pudo abrir archivo procesado: {e}")
    
    def _procesar_cargo_individual(self, ventana, hoja, fila: int, cuenta: str,
                                 importe: float, memo: str, motivo: str,
                                 glosa1: str, glosa2: str, glosa3: str) -> bool:
        """Procesa un cargo individual"""
        try:
            self.logger.info(f"Iniciando carga en el sistema para memo: {memo}")
            
            # Activar ventana del emulador
            if not ventana.isActive:
                ventana.maximize()
                ventana.activate()
            
            time.sleep(0.5)  # Esperar a que se active
            
            # Ejecutar secuencia de cargo
            pyautogui.press('f5')
            time.sleep(0.3)
            pyautogui.write('042')  # Código de cargo
            time.sleep(0.2)
            pyautogui.write(cuenta)
            time.sleep(0.2)
            pyautogui.write(self.formatear_monto(importe))
            time.sleep(0.2)
            pyautogui.press('tab')
            time.sleep(0.2)
            pyautogui.write(memo)
            time.sleep(0.2)
            pyautogui.press('tab')
            time.sleep(0.2)
            pyautogui.write(motivo)
            time.sleep(0.2)
            pyautogui.write(glosa1)
            time.sleep(0.2)
            pyautogui.press('tab')
            time.sleep(0.2)
            pyautogui.write(glosa2)
            time.sleep(0.2)
            pyautogui.press('tab')
            time.sleep(0.2)
            pyautogui.write(glosa3)
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(2.5)  # Esperar respuesta del emulador
            
            # Capturar respuesta de validación
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.5)
            foto_panel = pyperclip.paste()
            time.sleep(0.3)
            panel = foto_panel.splitlines()
            
            self.logger.info("=== PANEL DE VALIDACIÓN ===")
            for i, linea in enumerate(panel):
                self.logger.info(f"Línea {i}: '{linea}'")
            self.logger.info("=== FIN PANEL ===")
            
            # Extraer mensaje inteligentemente
            msj_emulacion = self._extraer_mensaje_validacion(panel)
            self.logger.info(f"Mensaje extraído: '{msj_emulacion}'")
            
            if "DATOS CORRECTOS PUEDE GRABAR" in msj_emulacion or "PUEDE GRABAR" in msj_emulacion:
                self.logger.info(f"Validación exitosa: {msj_emulacion}")
                return self._grabar_cargo(hoja, fila, memo)
            else:
                hoja.range(f'J{fila}').value = msj_emulacion
                hoja.range(f'K{fila}').value = "DATOS ERRONEOS"
                pyautogui.press('f5')
                self.logger.error(f"Error en validación para la fila {fila}: {msj_emulacion}")
                return False
                
        except Exception as e:
            self.logger.error(f"Error procesando cargo individual: {e}")
            return False
    
    def _grabar_cargo(self, hoja, fila: int, memo: str) -> bool:
        """Graba el cargo en el emulador"""
        try:
            pyautogui.press('f4')  # Grabar
            time.sleep(2.5)  # Esperar respuesta del emulador
            
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.5)
            foto_grabacion = pyperclip.paste()
            time.sleep(0.3)
            
            foto_grabacion_lineas = foto_grabacion.splitlines()
            
            self.logger.info("=== PANEL DE GRABACIÓN ===")
            for i, linea in enumerate(foto_grabacion_lineas):
                self.logger.info(f"Línea {i}: {linea}")
            self.logger.info("=== FIN GRABACIÓN ===")
            
            # Extraer mensaje de grabación
            msj_grabacion = self._extraer_mensaje_grabacion(foto_grabacion_lineas)
            self.logger.info(f"Mensaje de grabación: '{msj_grabacion}'")
            
            if "GRABACION CORRECTA" in msj_grabacion or ("GRABACION" in msj_grabacion and "CORRECTA" in msj_grabacion):
                hoja.range(f'J{fila}').value = msj_grabacion
                hoja.range(f'K{fila}').value = "OK"
                pyautogui.press('f5')
                self.logger.info(f"Grabación exitosa para la fila {fila}")
                return True
            else:
                hoja.range(f'J{fila}').value = msj_grabacion
                hoja.range(f'K{fila}').value = "REVISAR"
                pyautogui.press('f5')
                self.logger.error(f"Error en grabación para la fila {fila}: {msj_grabacion}")
                return False
                
        except Exception as e:
            self.logger.error(f"Error grabando cargo: {e}")
            return False
    
    def _extraer_mensaje_validacion(self, panel: list) -> str:
        """Extrae el mensaje de validación de forma inteligente"""
        # Buscar línea con 'PUEDE GRABAR' o 'DATOS CORRECTOS'
        for linea in panel:
            if "PUEDE GRABAR" in linea or "DATOS CORRECTOS" in linea:
                return linea.strip()
        
        # Si no encuentra, buscar líneas significativas (sin estar vacías)
        lineas_significativas = [l.strip() for l in panel if len(l.strip()) > 15]
        
        if lineas_significativas:
            # Retornar la última línea significativa
            return lineas_significativas[-1]
        
        return "NO SE DETECTÓ MENSAJE DE VALIDACIÓN"
    
    def _extraer_mensaje_grabacion(self, lineas: list) -> str:
        """Extrae el mensaje de grabación de forma inteligente"""
        # Buscar línea con 'GRABACION CORRECTA'
        for linea in lineas:
            if "GRABACION CORRECTA" in linea:
                return linea.strip()
        
        # Buscar líneas con palabras clave
        for linea in lineas:
            if "GRABACION" in linea and ("CORRECTA" in linea or "EXITOSA" in linea or "OK" in linea):
                return linea.strip()
        
        # Si no encuentra, buscar líneas significativas con 'GRABACION'
        lineas_significativas = [l.strip() for l in lineas if len(l.strip()) > 15 and "GRABACION" in l]
        
        if lineas_significativas:
            return lineas_significativas[-1]
        
        # Última opción: buscar cualquier línea relevante
        lineas_significativas = [l.strip() for l in lineas if len(l.strip()) > 15]
        if lineas_significativas:
            return lineas_significativas[-1]
        
        return "NO SE DETECTÓ MENSAJE DE GRABACIÓN"