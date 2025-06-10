"""
Operaciones para cuentas de Ahorros
Maneja abonos y cargos de cuentas de ahorro
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
from typing import Optional, List

from src.core.base_logic import BaseLogic
from src.utils.config_manager import ConfigManager
from src.utils.file_manager import FileManager


class AhorrosOperations(BaseLogic):
    """Clase para manejar operaciones de Ahorros"""
    
    def __init__(self):
        super().__init__("AHORROS")
        self.config_manager = ConfigManager()
        self.file_manager = FileManager()
        
        # Definir tipos de datos para las columnas
        self.dicc_tabla = {
            'ID': str,
            'Memo': str,
            'Cuenta_cargo': str,
            'Beneficiario': str,
            'Cuenta_abono': str,
            'Monto': float,
            'ITF': float,
            'Msj_abono': str,
            'Beneficiario_final': str,
            'Secuencia': str,
            'Estado': str
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
            self.logger.info("Iniciando detección de botones ESC para Ahorros")
            
            keyboard.hook(on_key_event)
            while self.deteccion_activa and not self.detener_proceso:
                time.sleep(0.1)
                
            self.logger.info("Detección de botones finalizada")
        except Exception as e:
            self.logger.error(f"Error en detección de botones: {e}")
            messagebox.showerror("ERROR", f"Problemas con la detección de botones: {e}")
        finally:
            keyboard.unhook(on_key_event)
    
    def execute_ahorros(self, ventana) -> bool:
        """
        Ejecuta el proceso de abono a cuentas de ahorro
        
        Args:
            ventana: Ventana del emulador bancario
        
        Returns:
            True si se completó correctamente
        """
        wb_ahorros = None
        book_ahorros = None
        ruta_procesado = ''
        finalizado = False
        
        try:
            self.iniciar_operacion()
            
            # Obtener configuración
            ruta_origen, ruta_destino = self.config_manager.leer_json("AHORROS")
            if not ruta_origen:
                messagebox.showerror("Error", "No se pudo obtener la configuración de Ahorros")
                return False
            
            directorio = os.path.dirname(ruta_origen)
            
            # Abrir Excel
            book_ahorros = xw.App(visible=False)
            wb_ahorros = book_ahorros.books.open(ruta_origen)
            hoja_ahorros = wb_ahorros.sheets['Ahorros']
            
            # Leer datos del Excel
            tabla_ahorros = pd.read_excel(ruta_origen, sheet_name='Ahorros', header=0, dtype=self.dicc_tabla)
            
            fecha_actual = self.get_fecha_actual()
            lista_memo_ahorros = set()
            cont_abonados = 0
            cont_abonos_incorrectos = 0
            
            self.logger.info(f"Procesando {len(tabla_ahorros)} registros de Ahorros")
            
            # Procesar cada fila
            for indice, fila in tabla_ahorros.iterrows():
                if self.detener_proceso:
                    break
                
                fila_ahorros = indice + 2
                
                # Extraer datos de la fila
                memorandum = str(fila['Memo']).strip()
                beneficiario = self.limpiar_texto_beneficiario(str(fila['Beneficiario']))
                cuenta_abono = self.limpiar_numero_cuenta(str(fila['Cuenta_abono']))
                monto = float(fila['Monto'])
                itf = fila['ITF']
                msj_abono = fila['Msj_abono']
                beneficiario_host = fila['Beneficiario_final']
                secuencia = fila['Secuencia']
                estado = fila['Estado']
                
                # Agregar memo a la lista
                if indice == 0:
                    lista_memo_ahorros.add(memorandum)
                if memorandum not in lista_memo_ahorros:
                    lista_memo_ahorros.add(memorandum)
                
                # Validar si debe procesarse
                if not self._debe_procesar_registro(estado, cuenta_abono):
                    continue
                
                # Procesar abono
                resultado = self._procesar_abono_ahorros(
                    ventana, hoja_ahorros, fila_ahorros, cuenta_abono, 
                    memorandum, monto, beneficiario, directorio, fecha_actual
                )
                
                if resultado['exito']:
                    cont_abonados += 1
                    if not resultado['beneficiario_correcto']:
                        cont_abonos_incorrectos += 1
                
                wb_ahorros.save()
            
            # Finalizar proceso
            self.finalizar_operacion()
            wb_ahorros.save()
            
            if self.detener_proceso:
                messagebox.showwarning(
                    "Proceso detenido",
                    "Se ha procedido a detener todos los procesos."
                )
            elif cont_abonados == 0 and cont_abonos_incorrectos == 0:
                messagebox.showinfo(
                    "Proceso no iniciado",
                    "No se ha realizado ningún abono, el excel ya está procesado o está vacío."
                )
            else:
                # Guardar archivo procesado
                ruta_procesado = self._guardar_archivo_procesado(
                    wb_ahorros, lista_memo_ahorros, directorio, fecha_actual
                )
                finalizado = True
                
                messagebox.showinfo(
                    "Proceso terminado",
                    f"Abonos realizados = {cont_abonados}\n"
                    f"Abonos rectificados = {cont_abonos_incorrectos}/{cont_abonados}"
                )
            
            return True
            
        except FileNotFoundError as fn:
            self.logger.error(f"Archivo no encontrado: {fn}")
            messagebox.showerror("Archivo no encontrado", 
                               f"Archivo excel no encontrado: {fn}")
            return False
        except Exception as e:
            self.logger.error(f"Error en ejecución Ahorros: {e}")
            messagebox.showerror("Error de ejecución", 
                               f"No se ha podido completar la ejecución Ahorros: {e}")
            return False
        finally:
            self.finalizar_operacion()
            if wb_ahorros and book_ahorros:
                try:
                    wb_ahorros.save()
                    wb_ahorros.close()
                    book_ahorros.quit()
                except Exception as e:
                    self.logger.warning(f"Error cerrando Excel: {e}")
            
            # Abrir archivo procesado
            if ruta_procesado and finalizado:
                try:
                    os.startfile(ruta_procesado)
                except Exception as e:
                    self.logger.warning(f"No se pudo abrir archivo procesado: {e}")
    
    def _debe_procesar_registro(self, estado: str, cuenta_abono: str) -> bool:
        """Determina si un registro debe ser procesado"""
        return pd.isna(estado) and len(cuenta_abono) == 11
    
    def _procesar_abono_ahorros(self, ventana, hoja, fila: int, cuenta_abono: str,
                              memorandum: str, monto: float, beneficiario_original: str,
                              directorio: str, fecha_actual: datetime) -> dict:
        """Procesa un abono individual a cuenta de ahorros"""
        try:
            # Activar ventana del emulador
            if not ventana.isActive:
                ventana.maximize()
                ventana.activate()
            
            # Limpiar clipboard y ejecutar comandos
            pyperclip.copy('')
            pyautogui.press('f5')
            pyautogui.write('441')  # Código para ahorros
            time.sleep(self.intervalo)
            pyautogui.write(cuenta_abono)
            time.sleep(self.intervalo)
            pyautogui.press('Tab')
            pyautogui.write(memorandum)
            
            if self.detener_proceso:
                return {'exito': False, 'beneficiario_correcto': True}
            
            time.sleep(self.intervalo)
            pyautogui.press('Tab')
            pyautogui.press('Tab')
            pyautogui.press('Tab')
            pyautogui.write(self.formatear_monto(monto))
            time.sleep(self.intervalo)
            
            if self.detener_proceso:
                return {'exito': False, 'beneficiario_correcto': True}
            
            pyautogui.press('f1')  # Grabar
            time.sleep(self.intervalo)
            pyautogui.hotkey('ctrl', 'c')
            pyautogui.sleep(0.3)
            
            # Procesar respuesta
            panel_host_ahorros = pyperclip.paste()
            lineas_ahorros = panel_host_ahorros.splitlines()
            
            if len(lineas_ahorros) > 23:
                msj_abono = lineas_ahorros[23].strip()
            else:
                msj_abono = "Error: Respuesta incompleta del sistema"
            
            if "OK" in msj_abono:
                return self._procesar_respuesta_exitosa(
                    hoja, fila, lineas_ahorros, beneficiario_original, 
                    directorio, fecha_actual, memorandum
                )
            else:
                # Error en el abono
                hoja.range(f'K{fila}').value = "Error con los datos"
                hoja.range(f'H{fila}').value = msj_abono
                return {'exito': False, 'beneficiario_correcto': True}
                
        except Exception as e:
            self.logger.error(f"Error procesando abono ahorros: {e}")
            return {'exito': False, 'beneficiario_correcto': True}
    
    def _procesar_respuesta_exitosa(self, hoja, fila: int, lineas: List[str],
                                  beneficiario_original: str, directorio: str,
                                  fecha_actual: datetime, memorandum: str) -> dict:
        """Procesa una respuesta exitosa del emulador"""
        try:
            # Extraer información de la respuesta
            secuencia = ""
            beneficiario_host = ""
            impuesto_itf = 0
            
            if len(lineas) > 13:
                # Obtener secuencia
                buscar_secuencia = lineas[13][:41] if len(lineas[13]) > 41 else lineas[13]
                secuencia = buscar_secuencia.replace('SECUENCIA', '').strip()
                hoja.range(f'J{fila}').value = secuencia
                
                # Obtener beneficiario del host
                if len(lineas[13]) > 41:
                    beneficiario_host = lineas[13][41:].strip()
                    hoja.range(f'I{fila}').value = beneficiario_host
            
            if len(lineas) > 15:
                # Obtener ITF
                buscar_itf = lineas[15]
                impuesto_itf = buscar_itf.replace('IMPUESTO ITF', '').strip()
                hoja.range(f'G{fila}').value = impuesto_itf
            
            # Guardar mensaje final
            if len(lineas) > 23:
                msj_abono = lineas[23].strip()
                hoja.range(f'H{fila}').value = msj_abono
            
            # Verificar si el beneficiario coincide
            beneficiario_correcto = True
            if beneficiario_original != beneficiario_host:
                beneficiario_correcto = self._manejar_beneficiario_incorrecto(
                    hoja, fila, beneficiario_original, secuencia, 
                    memorandum, directorio, fecha_actual
                )
            else:
                hoja.range(f'K{fila}').value = "GRABADO"
            
            return {
                'exito': True, 
                'beneficiario_correcto': beneficiario_correcto
            }
            
        except Exception as e:
            self.logger.error(f"Error procesando respuesta exitosa: {e}")
            return {'exito': False, 'beneficiario_correcto': True}
    
    def _manejar_beneficiario_incorrecto(self, hoja, fila: int, beneficiario: str,
                                       secuencia: str, memorandum: str,
                                       directorio: str, fecha_actual: datetime) -> bool:
        """Maneja el caso cuando el beneficiario no coincide"""
        try:
            # Capturar pantalla como evidencia
            filename = f"Memo {memorandum}_{beneficiario}-{secuencia}.png"
            ruta_screenshot = os.path.join(
                directorio, "Procesados", "Ahorros", "Alerta", 
                str(fecha_actual.year), filename
            )
            
            os.makedirs(os.path.dirname(ruta_screenshot), exist_ok=True)
            screenshot = pyautogui.screenshot()
            screenshot.save(ruta_screenshot)
            
            # Preguntar al usuario qué hacer
            respuesta = messagebox.askyesno(
                "Problemas",
                "Beneficiarios no coinciden ¿Deseas continuar?\n"
                "Si en caso va a extornar, no cierre esta ventana y haga el extorno.\n"
                "Marcar SÍ, grabará en el excel como GRABADO el abono, "
                "en caso marque NO se grabará como EXTORNADO"
            )
            
            if respuesta:
                hoja.range(f'K{fila}').value = "GRABADO"
                time.sleep(3)
                return True
            else:
                hoja.range(f'K{fila}').value = "EXTORNADO"
                return False
                
        except Exception as e:
            self.logger.error(f"Error manejando beneficiario incorrecto: {e}")
            hoja.range(f'K{fila}').value = "ERROR"
            return False
    
    def _guardar_archivo_procesado(self, wb, lista_memos: set, directorio: str, 
                                 fecha_actual: datetime) -> str:
        """Guarda el archivo procesado de ahorros"""
        try:
            nombre_archivo = self.file_manager.generar_nombre_archivo_procesado(
                "Ahorros", list(lista_memos)
            )
            
            ruta_directorio = self.file_manager.crear_directorio_procesados(
                directorio, "Ahorros", fecha_actual
            )
            
            ruta_procesado = os.path.join(ruta_directorio, nombre_archivo)
            wb.save(ruta_procesado)
            
            self.logger.info(f"Archivo Ahorros procesado guardado: {ruta_procesado}")
            return ruta_procesado
            
        except Exception as e:
            self.logger.error(f"Error guardando archivo procesado: {e}")
            return ""
    
    def execute_cargo_ahorros(self, ventana, archivo_xlc: str) -> bool:
        """
        Ejecuta el proceso de cargo para ahorros
        
        Args:
            ventana: Ventana del emulador
            archivo_xlc: Ruta del archivo Excel con historial
        
        Returns:
            True si se completó correctamente
        """
        try:
            self.logger.info(f"Iniciando cargo Ahorros desde: {archivo_xlc}")
            
            # Leer datos del archivo
            tabla_cargo = pd.read_excel(archivo_xlc, sheet_name='Ahorros', header=0, dtype=self.dicc_tabla)
            
            # Validar consistencia de datos
            memo_unicos = tabla_cargo['Memo'].nunique()
            cuenta_cargo_unicos = tabla_cargo['Cuenta_cargo'].nunique()
            
            if memo_unicos != 1 or cuenta_cargo_unicos != 1:
                messagebox.showinfo("Problema", 
                                  "Cuentas y/o números de memos diferentes en el excel")
                return False
            
            # Obtener datos únicos
            memo = tabla_cargo['Memo'].iloc[0]
            cuenta = tabla_cargo['Cuenta_cargo'].iloc[0]
            fecha_actual = self.get_fecha_actual()
            
            # Calcular totales de registros grabados
            filtro_grabados = tabla_cargo['Estado'] == 'GRABADO'
            suma_montos = tabla_cargo.loc[filtro_grabados, 'Monto'].sum()
            suma_itf = tabla_cargo.loc[filtro_grabados, 'ITF'].sum()
            
            suma_montos_str = self.formatear_monto(suma_montos)
            suma_itf_str = self.formatear_monto(suma_itf)
            
            # Activar ventana del emulador
            if not ventana.isActive:
                ventana.maximize()
                ventana.activate()
            
            # Ejecutar secuencia de cargo
            time.sleep(2)
            pyautogui.press('f5')
            pyautogui.write('042')  # Código de cargo
            pyautogui.write(cuenta)
            pyautogui.write(suma_montos_str)
            pyautogui.press('tab')
            pyautogui.write(memo)
            pyautogui.press('tab')
            pyautogui.write('84')  # Motivo
            pyautogui.write(f"MEMO {memo}-{fecha_actual.year}-BN-7101")
            pyautogui.press('tab')
            pyautogui.write(f"AHORROS S/.{suma_montos_str} ITF {suma_itf_str}")
            
            self.logger.info(f"Cargo Ahorros ejecutado - Monto: {suma_montos_str}, ITF: {suma_itf_str}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error en cargo Ahorros: {e}")
            messagebox.showerror("ERROR", 
                               f"El archivo seleccionado no es válido o no cumple el formato: {e}")
            return False
    
    @staticmethod
    def leer_xlc(ruta_xlc: str, memo: str, nro_cuenta: str, limpiar: bool = False) -> bool:
        """
        Lee y procesa un archivo Excel de Ahorros
        
        Args:
            ruta_xlc: Ruta del archivo a procesar
            memo: Número de memorándum
            nro_cuenta: Número de cuenta
            limpiar: Si debe limpiar el excel antes de procesar
        
        Returns:
            True si se procesó correctamente
        """
        config_manager = ConfigManager()
        file_manager = FileManager()
        logger = BaseLogic("AHORROS_Reader")
        
        ruta_origen = ''
        libro_activo = False
        
        try:
            # Obtener configuración
            ruta_origen, _ = config_manager.leer_json('AHORROS')
            if not ruta_origen:
                messagebox.showerror("Error", "No se pudo obtener la configuración")
                return False
            
            # Tipos de datos esperados
            dtype = {
                'N°': str,
                'Beneficiario': str,
                'Nº  Cuenta': str,
                'Tipo de Cuenta': str,
                'Entidad Financiera': str,
                'Monto (S/)': float
            }
            
            # Abrir archivos Excel
            with xw.App(visible=False) as book_ahorros:
                wb = book_ahorros.books.open(ruta_origen)
                hoja = wb.sheets['Ahorros']
                
                # Leer todas las hojas del archivo fuente
                hojas = pd.read_excel(ruta_xlc, dtype=dtype, sheet_name=None)
                
                hojas_validas = []
                hoja_no_util = False
                
                # Validar hojas
                for nombre_hoja, df in hojas.items():
                    if "DETRACCION" in nombre_hoja.upper():
                        continue
                    if set(df.columns) == set(dtype.keys()):
                        hojas_validas.append(df)
                    else:
                        hoja_no_util = True
                
                if not hojas_validas:
                    messagebox.showerror("Error", "No se encontraron hojas válidas en el archivo")
                    return False
                
                # Limpiar excel si se solicita
                if limpiar:
                    ultima_fila2 = hoja.range('E' + str(hoja.cells.last_cell.row)).end('up').row
                    valor_d2 = hoja.range('D2').value
                    if ultima_fila2 != 2 or valor_d2:
                        hoja.range('2:90').delete()
                        hoja.range('G2:K2').value = None
                
                # Combinar datos válidos
                df = pd.concat(hojas_validas, ignore_index=True)
                
                # Procesar datos - filtrar solo cuentas de ahorros
                df['Nº  Cuenta'] = df['Nº  Cuenta'].str.replace('-', '', regex=False)
                df_filtro = df.dropna(subset=['N°', 'Nº  Cuenta', 'Monto (S/)'])
                df_ahorros = df_filtro[df_filtro['Nº  Cuenta'].str.len() == 11]
                df_final = df_ahorros[df_ahorros['Tipo de Cuenta'] == 'AHORROS']
                df_final.reset_index(drop=True, inplace=True)
                
                # Encontrar última fila
                ultima_fila = hoja.range('E' + str(hoja.cells.last_cell.row)).end('up').row
                valor_f2 = hoja.range('F2').value
                if valor_f2 is None:
                    ultima_fila = 1
                
                # Procesar cada registro
                for _, fila in df_final.iterrows():
                    ultima_fila += 1
                    
                    beneficiario = BaseLogic("").limpiar_texto_beneficiario(fila['Beneficiario'])
                    cuenta = BaseLogic("").limpiar_numero_cuenta(fila['Nº  Cuenta'])
                    monto = fila['Monto (S/)']
                    
                    # Escribir en excel
                    hoja.range(f'B{ultima_fila}').value = memo
                    hoja.range(f'C{ultima_fila}').value = nro_cuenta
                    hoja.range(f'D{ultima_fila}').value = beneficiario
                    hoja.range(f'E{ultima_fila}').value = cuenta
                    hoja.range(f'F{ultima_fila}').value = monto
                    wb.save()
                
                wb.close()
                libro_activo = True
                
                mensaje = "Revisar excel." if not hoja_no_util else 'Revisar excel. Algunas hojas se omitieron'
                messagebox.showinfo("Proceso terminado", mensaje)
                
            return True
            
        except FileNotFoundError:
            messagebox.showwarning("ADVERTENCIA", "Archivo movido o no encontrado")
            return False
        except Exception as ex:
            logger.logger.error(f"Error procesando archivo Ahorros: {ex}")
            messagebox.showwarning("ERROR", str(ex))
            return False
        finally:
            if not libro_activo and ruta_origen:
                try:
                    os.startfile(ruta_origen)
                except Exception:
                    pass