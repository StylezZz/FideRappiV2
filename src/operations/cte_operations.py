"""
Operaciones para Cuentas Corrientes
Maneja transferencias entre cuentas corrientes (cargo y abono)
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
from typing import Optional, Dict

from src.core.base_logic import BaseLogic
from src.utils.config_manager import ConfigManager
from src.utils.file_manager import FileManager


class CTEOperations(BaseLogic):
    """Clase para manejar operaciones de Cuentas Corrientes"""
    
    def __init__(self):
        super().__init__("CTA_CTES")
        self.config_manager = ConfigManager()
        self.file_manager = FileManager()
        
        # Definir tipos de datos para las columnas
        self.dicc_tabla_cte = {
            'ID': str,
            'Memorandum': str,
            'Cta_cargo': str,
            'Cta_abono': str,
            'Monto': float,
            'Glosa': str,
            'Comision': str,
            'ITF_cargo': float,
            'ITF_abono': float,
            'Observacion': str,
            'Mensaje_cargo': str,
            'Mensaje_abono': str
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
            self.logger.info("Iniciando detección de botones ESC para Cuentas Corrientes")
            
            keyboard.hook(on_key_event)
            while self.deteccion_activa and not self.detener_proceso:
                time.sleep(0.1)
                
            self.logger.info("Detección de botones finalizada")
        except Exception as e:
            self.logger.error(f"Error en detección de botones: {e}")
            messagebox.showerror("ERROR", f"Problemas con la detección de botones: {e}")
        finally:
            keyboard.unhook(on_key_event)
    
    def execute_ctas_ctes(self, ventana) -> bool:
        """
        Ejecuta el proceso de transferencias entre cuentas corrientes
        
        Args:
            ventana: Ventana del emulador bancario
        
        Returns:
            True si se completó correctamente
        """
        wb_cte = None
        book_cte = None
        ruta_procesado = ''
        finalizado = False
        
        try:
            self.iniciar_operacion()
            
            # Obtener configuración
            ruta_origen, ruta_destino = self.config_manager.leer_json("CTA_CTES")
            if not ruta_origen:
                messagebox.showerror("Error", "No se pudo obtener la configuración de Cuentas Corrientes")
                return False
            
            directorio = os.path.dirname(ruta_origen)
            
            # Abrir Excel
            book_cte = xw.App(visible=False)
            wb_cte = book_cte.books.open(ruta_origen)
            hoja_cte = wb_cte.sheets['Corriente']
            
            # Leer datos del Excel
            tabla_cte = pd.read_excel(ruta_origen, sheet_name='Corriente', header=0, dtype=self.dicc_tabla_cte)
            
            fecha_actual = self.get_fecha_actual()
            lista_memo_cte = set()
            cont_abonados = 0
            cont_no_abonados = 0
            cont_cargados = 0
            cont_no_cargados = 0
            
            self.logger.info(f"Procesando {len(tabla_cte)} registros de Cuentas Corrientes")
            
            # Procesar cada fila
            for indice, fila in tabla_cte.iterrows():
                if self.detener_proceso:
                    break
                
                fila_cte = indice + 2
                
                # Extraer datos de la fila
                memorandum = str(fila['Memorandum']).strip()
                cta_cargo = self.limpiar_numero_cuenta(str(fila['Cta_cargo']))
                cta_abono = self.limpiar_numero_cuenta(str(fila['Cta_abono']))
                monto = float(fila['Monto'])
                glosa = str(fila['Glosa']).strip()
                comision = str(fila['Comision']).strip()
                itf = fila['ITF_cargo'] if not pd.isna(fila['ITF_cargo']) else 0
                observacion = fila['Observacion']
                mensaje_cargo = fila['Mensaje_cargo']
                mensaje_abono = fila['Mensaje_abono']
                
                # Agregar memo a la lista
                if indice == 0:
                    lista_memo_cte.add(memorandum)
                if memorandum not in lista_memo_cte:
                    lista_memo_cte.add(memorandum)
                
                # Validar datos
                if not self._validar_datos_cte(cta_cargo, cta_abono):
                    continue
                
                pyperclip.copy('')
                validar_cargo = False
                
                # PROCESO DE CARGO
                if self._debe_procesar_cargo(observacion, mensaje_cargo, cta_cargo, cta_abono):
                    resultado_cargo = self._procesar_cargo_cte(
                        ventana, hoja_cte, fila_cte, cta_cargo, monto, 
                        memorandum, comision, glosa, cta_abono
                    )
                    
                    if resultado_cargo['exito']:
                        cont_cargados += 1
                        validar_cargo = True
                        if resultado_cargo['itf']:
                            hoja_cte.range(f'H{fila_cte}').value = resultado_cargo['itf']
                    else:
                        cont_no_cargados += 1
                
                # PROCESO DE ABONO
                if self._debe_procesar_abono(mensaje_abono, mensaje_cargo, validar_cargo):
                    resultado_abono = self._procesar_abono_cte(
                        ventana, hoja_cte, fila_cte, cta_abono, monto,
                        memorandum, glosa, cta_cargo, itf
                    )
                    
                    if resultado_abono['exito']:
                        cont_abonados += 1
                        # Actualizar ITF total si es necesario
                        self._actualizar_itf_total(hoja_cte, fila_cte, resultado_abono['itf'], 
                                                 resultado_cargo.get('itf', 0) if 'resultado_cargo' in locals() else 0, itf)
                    else:
                        cont_no_abonados += 1
                
                wb_cte.save()
            
            # Finalizar proceso
            self.finalizar_operacion()
            wb_cte.save()
            
            if self.detener_proceso:
                messagebox.showwarning(
                    "Proceso detenido",
                    "Se ha procedido a detener todos los procesos."
                )
            elif all(count == 0 for count in [cont_abonados, cont_no_abonados, cont_cargados, cont_no_cargados]):
                messagebox.showwarning(
                    "Proceso no iniciado",
                    "Revisar excel, no hay cargos ni abonos por procesar."
                )
            else:
                # Guardar archivo procesado
                ruta_procesado = self._guardar_archivo_procesado(
                    wb_cte, lista_memo_cte, directorio, fecha_actual
                )
                finalizado = True
                
                messagebox.showinfo(
                    "Proceso finalizado",
                    f"Cargos realizados = {cont_cargados}\n"
                    f"Cargos no realizados = {cont_no_cargados}\n"
                    f"Abonos realizados = {cont_abonados}\n" 
                    f"Abonos no realizados = {cont_no_abonados}"
                )
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error en ejecución Cuentas Corrientes: {e}")
            messagebox.showerror("Error de ejecución", 
                               f"No se ha podido completar la ejecución: {e}")
            return False
        finally:
            self.finalizar_operacion()
            if wb_cte and book_cte:
                try:
                    wb_cte.save()
                    wb_cte.close()
                    book_cte.quit()
                except Exception as e:
                    self.logger.warning(f"Error cerrando Excel: {e}")
            
            # Abrir archivo procesado
            if ruta_procesado and finalizado:
                try:
                    os.startfile(ruta_procesado)
                except Exception as e:
                    self.logger.warning(f"No se pudo abrir archivo procesado: {e}")
    
    def _validar_datos_cte(self, cta_cargo: str, cta_abono: str) -> bool:
        """Valida que las cuentas tengan la longitud correcta"""
        return len(cta_cargo) == 11 and len(cta_abono) == 11
    
    def _debe_procesar_cargo(self, observacion: str, mensaje_cargo: str, 
                           cta_cargo: str, cta_abono: str) -> bool:
        """Determina si debe procesarse el cargo"""
        return (pd.isna(observacion) and pd.isna(mensaje_cargo) and 
                len(cta_cargo) == 11 and len(cta_abono) == 11)
    
    def _debe_procesar_abono(self, mensaje_abono: str, mensaje_cargo: str, validar_cargo: bool) -> bool:
        """Determina si debe procesarse el abono"""
        return (pd.isna(mensaje_abono) and 
                (not pd.isna(mensaje_cargo) and "GRABACION CORRECTA" in str(mensaje_cargo) or validar_cargo))
    
    def _procesar_cargo_cte(self, ventana, hoja, fila: int, cta_cargo: str, 
                          monto: float, memorandum: str, comision: str, 
                          glosa: str, cta_abono: str) -> Dict:
        """Procesa el cargo en cuenta corriente"""
        try:
            # Activar ventana del emulador
            if not ventana.isActive:
                ventana.maximize()
                ventana.activate()
            
            # Determinar código de transacción según tipo de cuenta
            cod_cargo = self._determinar_codigo_cargo(cta_cargo)
            
            # Ejecutar secuencia de cargo
            pyautogui.press('f5')
            if self.detener_proceso:
                return {'exito': False, 'itf': 0}
            
            pyautogui.write(cod_cargo)
            time.sleep(self.intervalo)
            pyautogui.write(cta_cargo)
            time.sleep(self.intervalo)
            pyautogui.write(self.formatear_monto(monto))
            
            if self.detener_proceso:
                return {'exito': False, 'itf': 0}
            
            pyautogui.press('enter')
            time.sleep(self.intervalo)
            pyautogui.write(memorandum)
            time.sleep(self.intervalo)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.write(comision)
            time.sleep(self.intervalo)
            pyautogui.write(glosa)
            
            if self.detener_proceso:
                return {'exito': False, 'itf': 0}
            
            time.sleep(self.intervalo)
            pyautogui.press('tab')
            pyautogui.write(f"TRANSF A CTA CTE BN {cta_abono}")
            pyautogui.press('tab')
            pyautogui.write(cta_abono)
            pyautogui.press('enter')
            
            if self.detener_proceso:
                return {'exito': False, 'itf': 0}
            
            # Procesar respuesta del emulador
            if not ventana.isActive:
                ventana.maximize()
                ventana.activate()
            
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            panel_emulacion = pyperclip.paste()
            lineas_emulacion = panel_emulacion.splitlines()
            
            if len(lineas_emulacion) > 23:
                linea_msj = lineas_emulacion[23].strip()
            else:
                linea_msj = "Error: Respuesta incompleta"
            
            if 'DATOS CORRECTOS PUEDE GRABAR' in linea_msj:
                # Extraer ITF del cargo
                itf_cargo = 0
                if len(lineas_emulacion) > 7:
                    itf_texto = lineas_emulacion[7][61:].strip()
                    if itf_texto:
                        itf_cargo = float(itf_texto)
                
                # Grabar la operación
                if self.detener_proceso:
                    return {'exito': False, 'itf': itf_cargo}
                
                pyperclip.copy('')
                time.sleep(0.3)
                pyautogui.press('f4')  # Grabar
                time.sleep(0.3)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.3)
                panel_grabacion = pyperclip.paste()
                lineas_grabacion = panel_grabacion.splitlines()
                
                if len(lineas_grabacion) > 23:
                    mensaje_grabacion = lineas_grabacion[23].strip()
                else:
                    mensaje_grabacion = "Error en grabación"
                
                if 'GRABACION CORRECTA' in mensaje_grabacion:
                    hoja.range(f'I{fila}').value = 'CARGADO'
                    hoja.range(f'J{fila}').value = mensaje_grabacion
                    return {'exito': True, 'itf': itf_cargo}
                else:
                    hoja.range(f'I{fila}').value = 'NO CARGADO'
                    hoja.range(f'J{fila}').value = mensaje_grabacion
                    return {'exito': False, 'itf': itf_cargo}
            
            elif "CUENTA SOBREGIRADA" in linea_msj:
                hoja.range(f'I{fila}').value = 'SIN FONDOS'
                hoja.range(f'J{fila}').value = linea_msj
                return {'exito': False, 'itf': 0}
            else:
                hoja.range(f'I{fila}').value = 'NO CARGADO'
                hoja.range(f'J{fila}').value = linea_msj
                return {'exito': False, 'itf': 0}
                
        except Exception as e:
            self.logger.error(f"Error procesando cargo CTE: {e}")
            return {'exito': False, 'itf': 0}
    
    def _procesar_abono_cte(self, ventana, hoja, fila: int, cta_abono: str,
                          monto: float, memorandum: str, glosa: str, 
                          cta_cargo: str, itf_previo: float) -> Dict:
        """Procesa el abono en cuenta corriente"""
        try:
            # Activar ventana del emulador
            if not ventana.isActive:
                ventana.maximize()
                ventana.activate()
            
            # Determinar código de transacción según tipo de cuenta
            cod_abono = self._determinar_codigo_abono(cta_abono)
            
            # Ejecutar secuencia de abono
            pyautogui.press('f5')
            pyperclip.copy('')
            if self.detener_proceso:
                return {'exito': False, 'itf': 0}
            
            pyautogui.write(cod_abono)
            time.sleep(self.intervalo)
            pyautogui.write(cta_abono)
            time.sleep(self.intervalo)
            pyautogui.write(self.formatear_monto(monto))
            
            if self.detener_proceso:
                return {'exito': False, 'itf': 0}
            
            pyautogui.press('enter')
            time.sleep(self.intervalo)
            pyautogui.write(memorandum)
            
            if self.detener_proceso:
                return {'exito': False, 'itf': 0}
            
            time.sleep(self.intervalo)
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            time.sleep(self.intervalo)
            pyautogui.write(glosa)
            time.sleep(self.intervalo)
            pyautogui.press('tab')
            pyautogui.write(f"TRANSF DE CTA CTE BN {cta_cargo}")
            
            if self.detener_proceso:
                return {'exito': False, 'itf': 0}
            
            pyautogui.press('tab')
            pyautogui.write("00000000000")
            pyautogui.press('enter')
            pyautogui.sleep(0.4)
            pyautogui.hotkey('ctrl', 'c')
            pyautogui.sleep(0.4)
            panel_emulacion = pyperclip.paste()
            lineas_emulacion = panel_emulacion.splitlines()
            
            if len(lineas_emulacion) > 23:
                linea_msj = lineas_emulacion[23].strip()
            else:
                linea_msj = "Error: Respuesta incompleta"
            
            if 'DATOS CORRECTOS PUEDE GRABAR' in linea_msj:
                pyperclip.copy('')
                time.sleep(0.3)
                pyautogui.press('f4')  # Grabar
                pyautogui.sleep(0.5)
                pyautogui.hotkey('ctrl', 'c')
                pyautogui.sleep(0.5)
                panel_grabacion = pyperclip.paste()
                lineas_grabacion = panel_grabacion.splitlines()
                
                # Extraer ITF del abono
                itf_abono = 0
                if len(lineas_grabacion) > 7:
                    itf_texto = lineas_grabacion[7][61:].strip()
                    if itf_texto:
                        itf_abono = float(itf_texto)
                
                if len(lineas_grabacion) > 23:
                    mensaje_grabacion = lineas_grabacion[23].strip()
                else:
                    mensaje_grabacion = "Error en grabación"
                
                if 'GRABACION CORRECTA' in mensaje_grabacion:
                    hoja.range(f'I{fila}').value = 'CARGADO Y ABONADO'
                    hoja.range(f'K{fila}').value = mensaje_grabacion
                    return {'exito': True, 'itf': itf_abono}
                else:
                    hoja.range(f'I{fila}').value = 'NO ABONADO'
                    hoja.range(f'K{fila}').value = mensaje_grabacion
                    return {'exito': False, 'itf': itf_abono}
            else:
                hoja.range(f'I{fila}').value = 'NO ABONADO'
                hoja.range(f'K{fila}').value = linea_msj
                return {'exito': False, 'itf': 0}
                
        except Exception as e:
            self.logger.error(f"Error procesando abono CTE: {e}")
            return {'exito': False, 'itf': 0}
    
    def _determinar_codigo_cargo(self, cuenta: str) -> str:
        """Determina el código de cargo según el tipo de cuenta"""
        if cuenta.startswith(('00068', '00000', '00015')):
            return '312'
        return '322'
    
    def _determinar_codigo_abono(self, cuenta: str) -> str:
        """Determina el código de abono según el tipo de cuenta"""
        if cuenta.startswith(('00068', '00000', '00015')):
            return '311'
        return '321'
    
    def _actualizar_itf_total(self, hoja, fila: int, itf_abono: float, 
                            itf_cargo: float, itf_previo: float):
        """Actualiza el ITF total en la hoja"""
        try:
            if itf_previo != 0 and itf_cargo == 0:
                # ITF previo sin ITF de cargo
                suma_itf = itf_abono + itf_previo
                hoja.range(f'H{fila}').value = suma_itf
            elif itf_cargo > 0:
                # ITF de cargo existente
                suma_itf = itf_cargo + itf_abono
                hoja.range(f'H{fila}').value = suma_itf
            else:
                # Solo ITF de abono
                hoja.range(f'H{fila}').value = itf_abono
        except Exception as e:
            self.logger.error(f"Error actualizando ITF total: {e}")
    
    def _guardar_archivo_procesado(self, wb, lista_memos: set, directorio: str, 
                                 fecha_actual: datetime) -> str:
        """Guarda el archivo procesado de cuentas corrientes"""
        try:
            nombre_archivo = self.file_manager.generar_nombre_archivo_procesado(
                "Cuentas corrientes", list(lista_memos)
            )
            
            ruta_directorio = self.file_manager.crear_directorio_procesados(
                directorio, "Cuentas corrientes", fecha_actual
            )
            
            ruta_procesado = os.path.join(ruta_directorio, nombre_archivo)
            wb.save(ruta_procesado)
            
            self.logger.info(f"Archivo Cuentas Corrientes procesado guardado: {ruta_procesado}")
            return ruta_procesado
            
        except Exception as e:
            self.logger.error(f"Error guardando archivo procesado: {e}")
            return ""
    
    @staticmethod
    def leer_xlc(ruta_xlc: str, memo: str, nro_cuenta: str, year: str, limpiar: bool = False) -> bool:
        """
        Lee y procesa un archivo Excel de Cuentas Corrientes
        
        Args:
            ruta_xlc: Ruta del archivo a procesar
            memo: Número de memorándum
            nro_cuenta: Número de cuenta de cargo
            year: Año del memorándum
            limpiar: Si debe limpiar el excel antes de procesar
        
        Returns:
            True si se procesó correctamente
        """
        config_manager = ConfigManager()
        logger = BaseLogic("CTE_Reader")
        
        ruta_origen = ''
        libro_activo = False
        
        try:
            # Obtener configuración
            ruta_origen, _ = config_manager.leer_json('CTA_CTES')
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
            with xw.App(visible=False) as book_cte:
                wb = book_cte.books.open(ruta_origen)
                hoja = wb.sheets['Corriente']
                
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
                    valor_c2 = hoja.range('C2').value
                    if ultima_fila2 != 2 or valor_c2:
                        hoja.range('2:90').delete()
                        hoja.range('H2:I2').value = None
                
                # Combinar datos válidos
                df = pd.concat(hojas_validas, ignore_index=True)
                
                # Procesar datos - filtrar solo cuentas corrientes
                df['Nº  Cuenta'] = df['Nº  Cuenta'].str.replace('-', '', regex=False)
                df_filtro = df.dropna(subset=['N°', 'Nº  Cuenta', 'Monto (S/)'])
                df_cta_valido = df_filtro[df_filtro['Tipo de Cuenta'] == 'CORRIENTE']
                df_final = df_cta_valido[df_cta_valido['Nº  Cuenta'].str.len() == 11]
                df_final.reset_index(drop=True, inplace=True)
                
                # Encontrar última fila
                ultima_fila = hoja.range('E' + str(hoja.cells.last_cell.row)).end('up').row
                valor_b2 = hoja.range('B2').value
                if valor_b2 is None:
                    ultima_fila = 1
                
                # Procesar cada registro
                for _, fila in df_final.iterrows():
                    ultima_fila += 1
                    
                    beneficiario = BaseLogic("").limpiar_texto_beneficiario(fila['Beneficiario'])
                    cuenta_abono = BaseLogic("").limpiar_numero_cuenta(fila['Nº  Cuenta'])
                    monto = fila['Monto (S/)']
                    
                    # Escribir en excel
                    hoja.range(f'B{ultima_fila}').value = memo
                    hoja.range(f'C{ultima_fila}').value = nro_cuenta
                    hoja.range(f'D{ultima_fila}').value = cuenta_abono
                    hoja.range(f'E{ultima_fila}').value = monto
                    hoja.range(f'F{ultima_fila}').value = f'MEMO {memo}-{year}-BN-7101'
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
            logger.logger.error(f"Error procesando archivo Cuentas Corrientes: {ex}")
            messagebox.showwarning("ERROR", str(ex))
            return False
        finally:
            if not libro_activo and ruta_origen:
                try:
                    os.startfile(ruta_origen)
                except Exception:
                    pass