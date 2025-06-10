"""
Operaciones para CCE (Cámara de Compensación Electrónica)
Maneja abonos y cargos de CCE
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
from typing import Optional, Tuple, List

from src.core.base_logic import BaseLogic
from src.utils.config_manager import ConfigManager
from src.utils.file_manager import FileManager


class CCEOperations(BaseLogic):
    """Clase para manejar operaciones de CCE"""
    
    def __init__(self):
        super().__init__("CCE")
        self.config_manager = ConfigManager()
        self.file_manager = FileManager()
        
        # Definir tipos de datos para las columnas
        self.dicc_tabla = {
            'ID': str,
            'Memorandum': str,
            'Cuenta': str,
            'Beneficiario': str,
            'CCI': str,
            'Monto': float,
            'IB': float,
            'BN': float,
            'COMENTARIO': str,
            'MENSAJE_EMULADOR': str
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
            self.logger.info("Iniciando detección de botones ESC para CCE")
            
            keyboard.hook(on_key_event)
            while self.deteccion_activa and not self.detener_proceso:
                time.sleep(0.1)
                
            self.logger.info("Detección de botones finalizada")
        except Exception as e:
            self.logger.error(f"Error en detección de botones: {e}")
            messagebox.showerror("ERROR", f"Problemas con la detección de botones: {e}")
        finally:
            keyboard.unhook(on_key_event)
    
    def execute_cce(self, ventana) -> bool:
        """
        Ejecuta el proceso de abono CCE
        
        Args:
            ventana: Ventana del emulador bancario
        
        Returns:
            True si se completó correctamente
        """
        ruta_procesado = ''
        finalizado = False
        wb_cce = None
        book_cce = None
        
        try:
            self.iniciar_operacion()
            
            # Obtener configuración
            ruta_origen, ruta_destino = self.config_manager.leer_json("CCE")
            if not ruta_origen:
                messagebox.showerror("Error", "No se pudo obtener la configuración de CCE")
                return False
            
            directorio = os.path.dirname(ruta_origen)
            
            # Abrir Excel
            book_cce = xw.App(visible=False)
            wb_cce = book_cce.books.open(ruta_origen)
            hoja_cce = wb_cce.sheets['CCE']
            
            # Leer datos del Excel
            tabla_cce = pd.read_excel(ruta_origen, sheet_name='CCE', header=0, dtype=self.dicc_tabla)
            
            fecha_actual = self.get_fecha_actual()
            lista_memo_cce = set()
            cont_abonados = 0
            cont_no_abonados = 0
            
            self.logger.info(f"Procesando {len(tabla_cce)} registros de CCE")
            
            # Procesar cada fila
            for indice, fila in tabla_cce.iterrows():
                if self.detener_proceso:
                    break
                
                fila_cce = indice + 2
                
                # Extraer datos de la fila
                memorandum = str(fila['Memorandum']).strip()
                cuenta = self.limpiar_numero_cuenta(str(fila['Cuenta']))
                beneficiario = self.limpiar_texto_beneficiario(str(fila['Beneficiario']))
                cci = self.limpiar_numero_cuenta(str(fila['CCI']))
                monto = float(fila['Monto'])
                comentario = fila['COMENTARIO']
                
                # Agregar memo a la lista para el nombre del archivo
                if indice == 0:
                    lista_memo_cce.add(memorandum)
                if memorandum not in lista_memo_cce:
                    lista_memo_cce.add(memorandum)
                
                # Validaciones
                if not self._validar_registro_cce(fila_cce, hoja_cce, cci, monto, comentario):
                    cont_no_abonados += 1
                    continue
                
                # Procesar abono
                resultado = self._procesar_abono_cce(
                    ventana, hoja_cce, fila_cce, cci, beneficiario, 
                    memorandum, cuenta, monto, directorio, fecha_actual
                )
                
                if resultado:
                    cont_abonados += 1
                else:
                    cont_no_abonados += 1
                
                wb_cce.save()
            
            # Finalizar proceso
            self.finalizar_operacion()
            wb_cce.save()
            
            if self.detener_proceso:
                messagebox.showwarning(
                    "Proceso detenido",
                    "Se ha procedido a detener todos los procesos."
                )
            elif cont_abonados == 0 and cont_no_abonados == 0:
                messagebox.showinfo(
                    "Proceso no iniciado",
                    "No se ha realizado ningún abono, el excel ya está procesado o está vacío."
                )
            else:
                # Guardar archivo procesado
                ruta_procesado = self._guardar_archivo_procesado(
                    wb_cce, lista_memo_cce, directorio, fecha_actual
                )
                finalizado = True
                
                messagebox.showinfo(
                    "Proceso finalizado",
                    f"Abonos realizados = {cont_abonados}\n"
                    f"Abonos no realizados = {cont_no_abonados}"
                )
            
            return True
            
        except FileNotFoundError as fn:
            self.logger.error(f"Archivo no encontrado: {fn}")
            messagebox.showerror("Archivo no encontrado", 
                               f"Archivo excel no encontrado: {fn}")
            return False
        except Exception as e:
            self.logger.error(f"Error en ejecución CCE: {e}")
            messagebox.showerror("Error de ejecución", 
                               f"No se ha podido completar la ejecución CCE: {e}")
            return False
        finally:
            self.finalizar_operacion()
            if wb_cce and book_cce:
                try:
                    wb_cce.save()
                    wb_cce.close()
                    book_cce.quit()
                except Exception as e:
                    self.logger.warning(f"Error cerrando Excel: {e}")
            
            # Abrir archivo procesado
            if ruta_procesado and finalizado:
                try:
                    os.startfile(ruta_procesado)
                except Exception as e:
                    self.logger.warning(f"No se pudo abrir archivo procesado: {e}")
    
    def _validar_registro_cce(self, fila: int, hoja: object, cci: str, 
                            monto: float, comentario: str) -> bool:
        """Valida un registro de CCE"""
        try:
            # Verificar si ya está procesado
            if not pd.isna(comentario):
                return False
            
            # Validar monto
            if monto >= 10000:
                hoja.range(f'I{fila}').value = 'Monto superior a los S/9,999.99'
                return False
            
            # Validar CCI
            if len(cci) != 20:
                hoja.range(f'I{fila}').value = 'Formato no correcto de CCI'
                return False
            
            return True
        except Exception as e:
            self.logger.error(f"Error validando registro: {e}")
            return False
    
    def _procesar_abono_cce(self, ventana, hoja, fila: int, cci: str, beneficiario: str,
                          memorandum: str, cuenta: str, monto: float, 
                          directorio: str, fecha_actual: datetime) -> bool:
        """Procesa un abono CCE individual"""
        try:
            # Activar ventana del emulador
            if not ventana.isActive:
                ventana.maximize()
                ventana.activate()
            
            # Limpiar clipboard
            pyperclip.copy('')
            
            # Ejecutar secuencia de comandos
            pyautogui.press('f5')
            time.sleep(self.intervalo)
            
            if self.detener_proceso:
                return False
            
            # Ingresar datos
            pyautogui.write('220')  # Código de transacción
            time.sleep(self.intervalo)
            pyautogui.write(cci)
            time.sleep(self.intervalo)
            pyautogui.write(beneficiario)
            time.sleep(self.intervalo)
            pyautogui.press('Tab')
            time.sleep(self.intervalo)
            pyautogui.write(f"MEMO {memorandum}-BN-7101")
            time.sleep(self.intervalo)
            pyautogui.press('Tab')
            time.sleep(self.intervalo)
            pyautogui.write(cuenta)
            time.sleep(self.intervalo)
            
            if self.detener_proceso:
                return False
            
            pyautogui.press('Enter')
            time.sleep(self.intervalo)
            pyautogui.write('sol')
            time.sleep(self.intervalo)
            pyautogui.write(self.formatear_monto(monto))
            time.sleep(self.intervalo)
            pyautogui.press('Tab')
            time.sleep(self.intervalo)
            pyautogui.press('Tab')
            time.sleep(self.intervalo)
            pyautogui.write('1')
            time.sleep(self.intervalo)
            
            if self.detener_proceso:
                return False
            
            pyautogui.press('Enter')
            time.sleep(self.intervalo)
            
            # Capturar respuesta del emulador
            pyautogui.hotkey('ctrl', 'c')
            pyautogui.sleep(0.3)
            panel_emulacion = pyperclip.paste()
            lineas_emulacion = panel_emulacion.splitlines()
            
            if self.detener_proceso:
                return False
            
            # Procesar respuesta
            return self._procesar_respuesta_emulador(
                hoja, fila, lineas_emulacion, memorandum, beneficiario, 
                directorio, fecha_actual
            )
            
        except Exception as e:
            self.logger.error(f"Error procesando abono CCE: {e}")
            return False
    
    def _procesar_respuesta_emulador(self, hoja, fila: int, lineas: List[str],
                                   memorandum: str, beneficiario: str,
                                   directorio: str, fecha_actual: datetime) -> bool:
        """Procesa la respuesta del emulador"""
        try:
            # Extraer comisiones y mensajes
            for linea in lineas:
                # Procesar comisión IB
                if "COMISION IB" in linea:
                    comision_ib = linea[35:42].strip()
                    hoja.range(f'G{fila}').value = float(comision_ib) if comision_ib else 0
                
                # Procesar comisión BN
                if "COMISION BN" in linea:
                    comision_bn = linea[35:42].strip()
                    hoja.range(f'H{fila}').value = float(comision_bn) if comision_bn else 0
                
                # Procesar mensaje de validación
                if "MSG" in linea:
                    msg_emulacion = linea[7:54].strip()
                    
                    if '**DATOS CORRECTOS' in msg_emulacion:
                        # Grabar operación
                        return self._grabar_operacion(hoja, fila, memorandum, beneficiario, 
                                                    directorio, fecha_actual)
                    else:
                        # Error en validación
                        hoja.range(f'I{fila}').value = 'ERROR DE GRABACIÓN'
                        hoja.range(f'J{fila}').value = msg_emulacion
                        
                        # Capturar pantalla para evidencia
                        self._capturar_pantalla_error(memorandum, beneficiario, 
                                                    directorio, fecha_actual)
                        
                        pyautogui.press('f5')
                        return False
            
            return False
            
        except Exception as e:
            self.logger.error(f"Error procesando respuesta emulador: {e}")
            return False
    
    def _grabar_operacion(self, hoja, fila: int, memorandum: str, beneficiario: str,
                        directorio: str, fecha_actual: datetime) -> bool:
        """Graba la operación en el emulador"""
        try:
            pyperclip.copy('')
            pyautogui.sleep(0.5)
            pyautogui.press('f4')  # Grabar
            pyautogui.sleep(0.5)
            pyautogui.hotkey('ctrl', 'c')
            pyautogui.sleep(0.5)
            
            panel_grabacion = pyperclip.paste()
            lineas_grabacion = panel_grabacion.splitlines()
            
            for linea in lineas_grabacion:
                if "MSG" in linea:
                    if "TRANSFERENCIA GRABADA" in linea:
                        # Grabación exitosa
                        hoja.range(f'I{fila}').value = 'ABONADO'
                        hoja.range(f'J{fila}').value = linea.strip()
                        return True
                    else:
                        # Error en grabación
                        hoja.range(f'I{fila}').value = 'NO ABONADO'
                        hoja.range(f'J{fila}').value = linea.strip()
                        return False
            
            return False
            
        except Exception as e:
            self.logger.error(f"Error grabando operación: {e}")
            return False
    
    def _capturar_pantalla_error(self, memorandum: str, beneficiario: str,
                               directorio: str, fecha_actual: datetime):
        """Captura pantalla en caso de error"""
        try:
            filename = f"Memo {memorandum} - {beneficiario}.png"
            ruta_screenshot = os.path.join(
                directorio, "Procesados", "CCE", "Alerta", 
                str(fecha_actual.year), filename
            )
            
            os.makedirs(os.path.dirname(ruta_screenshot), exist_ok=True)
            
            screenshot = pyautogui.screenshot()
            screenshot.save(ruta_screenshot)
            
            self.logger.info(f"Captura de error guardada: {ruta_screenshot}")
            
        except Exception as e:
            self.logger.error(f"Error capturando pantalla: {e}")
    
    def _guardar_archivo_procesado(self, wb, lista_memos: set, directorio: str, 
                                 fecha_actual: datetime) -> str:
        """Guarda el archivo procesado"""
        try:
            # Generar nombre y ruta del archivo
            fecha_info = self.format_fecha_archivo(fecha_actual)
            
            nombre_archivo = self.file_manager.generar_nombre_archivo_procesado(
                "CCE", list(lista_memos)
            )
            
            ruta_directorio = self.file_manager.crear_directorio_procesados(
                directorio, "CCE", fecha_actual
            )
            
            ruta_procesado = os.path.join(ruta_directorio, nombre_archivo)
            
            # Guardar archivo
            wb.save(ruta_procesado)
            
            self.logger.info(f"Archivo CCE procesado guardado: {ruta_procesado}")
            return ruta_procesado
            
        except Exception as e:
            self.logger.error(f"Error guardando archivo procesado: {e}")
            return ""
    
    def execute_cargo_cce(self, ventana, archivo_xlc: str) -> bool:
        """
        Ejecuta el proceso de cargo CCE
        
        Args:
            ventana: Ventana del emulador
            archivo_xlc: Ruta del archivo Excel con historial
        
        Returns:
            True si se completó correctamente
        """
        try:
            self.logger.info(f"Iniciando cargo CCE desde: {archivo_xlc}")
            
            # Leer datos del archivo
            tabla_cargo = pd.read_excel(archivo_xlc, sheet_name='CCE', header=0, dtype=self.dicc_tabla)
            
            # Validar consistencia de datos
            memo_unicos = tabla_cargo['Memorandum'].nunique()
            ctas_unicas = tabla_cargo['Cuenta'].nunique()
            
            if memo_unicos != 1 or ctas_unicas != 1:
                messagebox.showinfo("Problema", 
                                  "Cuentas y/o números de memos diferentes en el excel")
                return False
            
            # Obtener datos únicos
            memo = tabla_cargo['Memorandum'].iloc[0]
            cuenta = tabla_cargo['Cuenta'].iloc[0]
            nro_memo = memo.split('-')[0]
            
            # Calcular totales de registros abonados
            filtro_abonados = tabla_cargo['COMENTARIO'] == 'ABONADO'
            suma_montos = tabla_cargo.loc[filtro_abonados, 'Monto'].sum()
            suma_ib = tabla_cargo.loc[filtro_abonados, 'IB'].sum()
            suma_bn = tabla_cargo.loc[filtro_abonados, 'BN'].sum()
            
            # Formatear montos
            suma_montos_str = self.formatear_monto(suma_montos)
            suma_ib_str = self.formatear_monto(suma_ib)
            suma_bn_str = self.formatear_monto(suma_bn)
            
            importe_total = suma_montos + suma_ib + suma_bn
            importe_total_str = self.formatear_monto(importe_total)
            
            # Activar ventana del emulador
            if not ventana.isActive:
                ventana.maximize()
                ventana.activate()
            
            # Ejecutar secuencia de cargo
            time.sleep(2)
            pyautogui.press('f5')
            pyautogui.write('042')  # Código de cargo
            pyautogui.write(cuenta)
            pyautogui.write(importe_total_str)
            pyautogui.press('tab')
            pyautogui.write(nro_memo)
            pyautogui.press('tab')
            pyautogui.write('84')  # Motivo
            pyautogui.write(f"MEMO {memo}-BN-7101")
            pyautogui.press('tab')
            pyautogui.write(f"CCE {suma_montos_str} IB {suma_ib_str} BN {suma_bn_str}")
            
            self.logger.info(f"Cargo CCE ejecutado - Total: {importe_total_str}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error en cargo CCE: {e}")
            messagebox.showerror("ERROR", 
                               f"El archivo seleccionado no es válido o no cumple el formato: {e}")
            return False
    
    @staticmethod
    def leer_xlc(ruta_xlc: str, memo: str, nro_cuenta: str, limpiar: bool = False) -> bool:
        """
        Lee y procesa un archivo Excel de CCE
        
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
        logger = BaseLogic("CCE_Reader")
        
        ruta_origen = ''
        libro_activo = False
        
        try:
            # Obtener configuración
            ruta_origen, _ = config_manager.leer_json('CCE')
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
            with xw.App(visible=False) as book_cce:
                wb = book_cce.books.open(ruta_origen)
                hoja = wb.sheets['CCE']
                
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
                    valor_e2 = hoja.range('E2').value
                    if ultima_fila2 != 2 or valor_e2:
                        hoja.range('2:90').delete()
                        hoja.range('G2:J2').value = None
                
                # Combinar datos válidos
                df = pd.concat(hojas_validas, ignore_index=True)
                
                # Procesar datos
                df['Nº  Cuenta'] = df['Nº  Cuenta'].str.replace('-', '', regex=False)
                df_filtro = df.dropna(subset=['N°', 'Nº  Cuenta', 'Monto (S/)'])
                df_cci_valido = df_filtro[df_filtro['Nº  Cuenta'].str.len() == 20]
                df_cce = df_cci_valido[df_cci_valido['Monto (S/)'] < 10000].copy()
                df_cce.reset_index(drop=True, inplace=True)
                
                # Encontrar última fila
                ultima_fila = hoja.range('E' + str(hoja.cells.last_cell.row)).end('up').row
                valor_d2 = hoja.range('D2').value
                if valor_d2 is None:
                    ultima_fila = 1
                
                # Procesar cada registro
                for _, fila in df_cce.iterrows():
                    ultima_fila += 1
                    
                    beneficiario = BaseLogic("").limpiar_texto_beneficiario(fila['Beneficiario'])
                    cuenta_cci = BaseLogic("").limpiar_numero_cuenta(fila['Nº  Cuenta'])
                    monto = fila['Monto (S/)']
                    
                    # Escribir en excel
                    hoja.range(f'B{ultima_fila}').value = memo
                    hoja.range(f'C{ultima_fila}').value = nro_cuenta
                    hoja.range(f'D{ultima_fila}').value = beneficiario
                    hoja.range(f'E{ultima_fila}').value = cuenta_cci
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
            logger.logger.error(f"Error procesando archivo CCE: {ex}")
            messagebox.showwarning("ERROR", str(ex))
            return False
        finally:
            if not libro_activo and ruta_origen:
                try:
                    # Intentar abrir el archivo original
                    os.startfile(ruta_origen)
                except Exception:
                    pass