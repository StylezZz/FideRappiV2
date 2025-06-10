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
            
            # Ejecutar secuencia de cargo
            pyautogui.press('f5')
            pyautogui.write('042')  # Código de cargo
            pyautogui.write(cuenta)
            pyautogui.write(self.formatear_monto(importe))
            pyautogui.press('tab')
            pyautogui.write(memo)
            pyautogui.press('tab')
            pyautogui.write('84')  # Motivo fijo
            pyautogui.write(glosa1)
            pyautogui.press('tab')
            pyautogui.write(glosa2)
            pyautogui.press('tab')
            pyautogui.write(glosa3)
            pyautogui.press('enter')
            time.sleep(2)
            
            # Capturar respuesta de validación
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            foto_panel = pyperclip.paste()
            time.sleep(0.3)
            panel = foto_panel.splitlines()
            
            self.logger.info("Contenido del panel (análisis detallado):")
            for i, linea in enumerate(panel):
                self.logger.info(f"Línea {i}: '{linea}'")
            
            # Buscar mensaje de validación
            msj_emulacion = ""
            for linea in panel:
                if "DATOS CORRECTOS PUEDE GRABAR" in linea:
                    msj_emulacion = linea.strip()
                    self.logger.info(f"Mensaje de validación encontrado: '{msj_emulacion}'")
                    break
            
            # Si no se encuentra, usar línea predeterminada
            if not msj_emulacion and len(panel) > 23:
                msj_emulacion = panel[23].strip()
                self.logger.info(f"Usando línea predeterminada 23: '{msj_emulacion}'")
            
            self.logger.info(f"Mensaje final de validación: '{msj_emulacion}'")
            
            if "DATOS CORRECTOS PUEDE GRABAR" in msj_emulacion:
                self.logger.info(f"Validación exitosa: {msj_emulacion}")
                return self._grabar_cargo(hoja, fila, memo)
            else:
                # Error en validación
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
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            foto_grabacion = pyperclip.paste()
            time.sleep(0.3)
            
            foto_grabacion_lineas = foto_grabacion.splitlines()
            self.logger.info("Contenido de foto_grabacion:")
            for i, linea in enumerate(foto_grabacion_lineas):
                self.logger.info(f"Línea {i}: {linea}")
            
            # Buscar mensaje de grabación
            msj_grabacion = ""
            for linea in foto_grabacion_lineas:
                if "GRABACION CORRECTA" in linea:
                    msj_grabacion = linea.strip()
                    self.logger.info(f"¡Mensaje de grabación correcta encontrado: '{msj_grabacion}'!")
                    break
            
            # Si no se encontró, buscar en línea específica
            if not msj_grabacion and len(foto_grabacion_lineas) > 29:
                msj_grabacion = foto_grabacion_lineas[29].strip()
                self.logger.info(f"Usando línea 29: '{msj_grabacion}'")
            
            # Buscar cualquier mensaje relevante si aún está vacío
            if not msj_grabacion:
                for linea in foto_grabacion_lineas:
                    if any(msg in linea for msg in ["GRABACION", "ERROR", "RECHAZADO", "CORRECTO"]):
                        msj_grabacion = linea.strip()
                        self.logger.info(f"Encontrado mensaje alternativo: '{msj_grabacion}'")
                        break
            
            self.logger.info(f"Mensaje final de grabación: '{msj_grabacion}'")
            
            # Si aún está vacío, establecer valor predeterminado
            if not msj_grabacion:
                msj_grabacion = "NO SE DETECTÓ MENSAJE DE GRABACIÓN"
            
            if "GRABACION CORRECTA" in msj_grabacion:
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
    
    def _guardar_archivo_procesado(self, wb, lista_memos: set, directorio: str, 
                                 fecha_actual: datetime) -> str:
        """Guarda el archivo procesado de cargo"""
        try:
            nombre_archivo = self.file_manager.generar_nombre_archivo_procesado(
                "Cargo", list(lista_memos)
            )
            
            ruta_directorio = self.file_manager.crear_directorio_procesados(
                directorio, "Cargo", fecha_actual
            )
            
            ruta_procesado = os.path.join(ruta_directorio, nombre_archivo)
            wb.save(ruta_procesado)
            
            self.logger.info(f"Archivo Cargo procesado guardado: {ruta_procesado}")
            return ruta_procesado
            
        except Exception as e:
            self.logger.error(f"Error guardando archivo procesado: {e}")
            return ""