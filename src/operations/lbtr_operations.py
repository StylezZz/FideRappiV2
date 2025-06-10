"""
Operaciones LBTR (Sistema de Liquidación Bruta en Tiempo Real)
Maneja transferencias interbancarias a través del sistema web LBTR
"""

import time
import os
import pandas as pd
import xlwings as xw
import pyautogui
import pyperclip
from tkinter import messagebox
import datetime
from typing import Optional, Dict

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import Select
    from selenium.common.exceptions import NoSuchElementException, TimeoutException
    from selenium.webdriver.common.keys import Keys
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False
    webdriver = None

from src.core.base_logic import BaseLogic
from src.utils.config_manager import ConfigManager
from src.utils.file_manager import FileManager


class LBTROperations(BaseLogic):
    """Clase para manejar operaciones LBTR"""
    
    def __init__(self):
        super().__init__("LBTR")
        self.config_manager = ConfigManager()
        self.file_manager = FileManager()
        
        # Verificar disponibilidad de Selenium
        if not SELENIUM_AVAILABLE:
            self.logger.error("Selenium no está disponible. Las operaciones LBTR no funcionarán.")
        
        # Definir tipos de datos para las columnas
        self.dicc_tabla = {
            'ID': str,
            'Cuenta_cargo': str,
            'OBS_1': str,
            'OBS_2': str,
            'Beneficiario': str,
            'CCI': str,
            'Entidad_Financiera': str,
            'Importe': float,
            'RUC': str,
            'DOMICILIO': str,
            'ESTADO': str
        }
        
        # Mapeo de entidades financieras
        self.entidades_financieras = {
            "BCRP": "0",
            "CREDITO": "1",
            "BCP": "1",
            "BCO. CREDITO": "1",
            "BCO CREDITO": "1",
            "BANCO DE CREDITO DEL PERU": "1",
            "INTERBANK": "2",
            "CITIBANK": "3",
            "SCOTIABANK": "4",
            "CONTINENTAL": "5",
            "BBVA": "5",
            "BCO. CONTINENTAL": "5",
            "COMERCIO": "6",
            "FINANCIERO": "7",
            "BIF": "8",
            "BANBIF": "8",
            "B.I.F.": "8",
            "CREDISCOTIA": "9",
            "MIBANCO": "10",
            "AGROBANCO": "11",
            "BCO GNB": "12",
            "FALABELLA": "13",
            "RIPLEY": "14",
            "SANTANDER": "15",
            "DEUTSCHE": "16",
            "AZTECA": "17",
            "BANCO CENCOSUD": "18",
            "ICBC PERU": "19",
            "BANK OF CHINA (PERU)": "20",
            "COFIDE": "21",
            "FIN.CREDINKA": "22",
            "F.CREDITO": "23",
            "FIN CMR": "24",
            "FIN TFC S.A.": "25",
            "CORDILLERA": "26",
            "FIN. EDYFICAR": "27",
            "COMPARTAMOS": "28",
            "FIN. CONFIANZA": "29",
            "FIN. UNIVERSAL": "30",
            "FIN. OH": "31",
            "AMERIKA FIN.": "32",
            "FIN. EFECTIVA": "33",
            "MITSUILEASING": "34",
            "PROEMPRESA": "35",
            "FIN.CONFIANZA S.A.A.": "36",
            "FONDO BCRP": "37",
            "F.S.D.": "38",
            "F.S.D.Cooperativo": "39",
            "CAVALI": "40",
            "M.E.F.": "41",
            "CAJA METROPOLITANA": "42",
            "CMAC. PIURA SAC": "43",
            "CAJA PIURA": "43",
            "CAJA MUNICI.TRUJILLO": "44",
            "CAJA MUNICI. AREQUIPA": "45",
            "CMAC. SULLANA": "46",
            "CMAC. CUZCO": "47",
            "CMAC SANTA": "48",
            "CMAC. HUANCAYO": "49",
            "CMAC. ICA": "50",
            "CMAC. PAITA": "51",
            "CMAC. MAYNAS": "52",
            "CMAC PISCO": "53",
            "CMAC. TACNA": "54",
            "CRAC. SN.MARTIN": "55",
            "CRAC. SR. DE LUREN": "56",
            "CRAC. TUMBAY": "57",
            "CREDINKA": "58",
            "CRAC. VALLE APU": "59",
            "CREDICHAVIN": "60",
            "CAJA NUESTRA GENTE": "61",
            "PROFINANZAS": "62",
            "CRAC.LOS LIBERT": "63",
            "CAJA SIPAN": "64",
            "CRAC. CAJAMARCA": "65",
            "CRAC SELVA PERU": "66",
            "CRAC-LOS ANDES": "67",
            "CAJA RURAL PRYMERA": "68",
            "CRAC DEL SUR": "69",
            "CRAC INCASUR S.A.": "70",
            "CCE": "71"
        }
    
    def exec_lbtr(self, usuario: str, clave: str) -> bool:
        """
        Ejecuta el proceso completo de LBTR
        
        Args:
            usuario: Usuario para login
            clave: Contraseña para login
        
        Returns:
            True si se completó correctamente
        """
        if not SELENIUM_AVAILABLE:
            messagebox.showerror(
                "Error",
                "Selenium no está disponible. Instale las dependencias necesarias:\n"
                "pip install selenium\n"
                "Y descargue el driver de Edge (msedgedriver.exe)"
            )
            return False
        
        driver = None
        wb_lbtr = None
        book_lbtr = None
        ruta_procesado = ''
        
        try:
            self.iniciar_operacion()
            
            # Obtener configuración
            ruta_origen, ruta_destino = self.config_manager.leer_json("LBTR")
            enlace = self.config_manager.lbtr_credenciales()
            
            if not ruta_origen or not enlace:
                messagebox.showerror("Error", "Configuración de LBTR incompleta")
                return False
            
            directorio = os.path.dirname(ruta_origen)
            
            # Abrir Excel
            book_lbtr = xw.App(visible=False)
            wb_lbtr = book_lbtr.books.open(ruta_origen)
            hoja_lbtr = wb_lbtr.sheets['LBTR']
            
            # Leer datos del Excel
            tabla_lbtr = pd.read_excel(ruta_origen, sheet_name='LBTR', header=0, dtype=self.dicc_tabla)
            
            fecha_actual = self.get_fecha_actual()
            lista_memo_lbtr = set()
            cont_abonados = 0
            cont_no_abonados = 0
            
            self.logger.info(f"Procesando {len(tabla_lbtr)} registros de LBTR")
            
            # Configurar e iniciar Selenium
            driver = self._init_selenium_driver(enlace)
            if not driver:
                return False
            
            # Hacer login
            if not self._login_lbtr(driver, usuario, clave):
                return False
            
            # Navegar al módulo de transferencias
            if not self._navigate_to_transfers(driver):
                return False
            
            # Procesar cada transferencia
            for indice, fila in tabla_lbtr.iterrows():
                if self.detener_proceso:
                    break
                
                fila_lbtr = indice + 2
                
                # Extraer datos de la fila
                obs_1 = str(fila['OBS_1']).strip()
                obs_2 = str(fila['OBS_2']).strip()
                beneficiario = self.limpiar_texto_beneficiario(str(fila['Beneficiario']))
                cci = self.limpiar_numero_cuenta(str(fila['CCI']))
                entidad_financiera = str(fila['Entidad_Financiera']).strip()
                importe = float(fila['Importe'])
                ruc_completo = str(fila['RUC']).strip()
                domicilio_completo = str(fila['DOMICILIO']).strip()
                estado = fila['ESTADO']
                
                # Validar si debe procesarse
                if not pd.isna(estado):
                    continue
                
                # Extraer datos de RUC y domicilio
                ruc = self._extract_after_colon(ruc_completo)
                domicilio = self._extract_after_colon(domicilio_completo)
                
                # Extraer número de memo
                titulo_memo = self._extract_memo_number(obs_1, obs_2)
                if titulo_memo:
                    if indice == 0:
                        lista_memo_lbtr.add(titulo_memo)
                    if titulo_memo not in lista_memo_lbtr:
                        lista_memo_lbtr.add(titulo_memo)
                
                # Procesar transferencia
                resultado = self._procesar_transferencia_lbtr(
                    driver, hoja_lbtr, fila_lbtr, obs_1, obs_2, beneficiario,
                    cci, entidad_financiera, importe, ruc, domicilio
                )
                
                if resultado:
                    cont_abonados += 1
                else:
                    cont_no_abonados += 1
                
                wb_lbtr.save()
                time.sleep(5)  # Pausa entre transferencias
            
            # Finalizar proceso
            wb_lbtr.save()
            
            if cont_abonados == 0 and cont_no_abonados == 0:
                messagebox.showinfo(
                    "Proceso no iniciado",
                    "No se ha realizado ningún abono, el excel ya está procesado o está vacío."
                )
            else:
                # Guardar archivo procesado
                ruta_procesado = self._guardar_archivo_procesado_lbtr(
                    wb_lbtr, lista_memo_lbtr, directorio, fecha_actual
                )
                
                messagebox.showinfo(
                    "Proceso terminado",
                    f"Transferencias realizadas = {cont_abonados}\n"
                    f"Transferencias fallidas = {cont_no_abonados}"
                )
            
            return True
            
        except TimeoutException:
            self.logger.error("Timeout en operación LBTR")
            messagebox.showerror("Error", "El elemento no fue encontrado o el tiempo de espera se agotó.")
            return False
        except Exception as e:
            self.logger.error(f"Error en ejecución LBTR: {e}")
            messagebox.showerror("Error", f"Ocurrió un error: {e}")
            return False
        finally:
            self.finalizar_operacion()
            
            # Cerrar recursos
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            
            if wb_lbtr and book_lbtr:
                try:
                    wb_lbtr.save()
                    wb_lbtr.close()
                    book_lbtr.quit()
                except:
                    pass
            
            # Abrir archivo procesado
            if ruta_procesado:
                try:
                    os.startfile(ruta_procesado)
                except:
                    pass
    
    def _init_selenium_driver(self, enlace: str):
        """Inicializa el driver de Selenium"""
        try:
            # Verificar que existe el driver
            driver_path = self._find_edge_driver()
            if not driver_path:
                messagebox.showerror(
                    "Error",
                    "No se encontró msedgedriver.exe\n"
                    "Descargue el driver desde: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/"
                )
                return None
            
            # Configurar opciones
            edge_options = Options()
            edge_options.add_argument('--ignore-certificate-errors')
            edge_options.add_argument("--inprivate")
            
            # Configurar servicio
            service = webdriver.EdgeService(executable_path=driver_path)
            
            # Crear driver
            driver = webdriver.Edge(service=service, options=edge_options)
            driver.maximize_window()
            driver.implicitly_wait(5)
            
            # Navegar a la página
            driver.get(enlace)
            
            return driver
            
        except Exception as e:
            self.logger.error(f"Error inicializando Selenium: {e}")
            messagebox.showerror("Error", f"Error inicializando navegador: {e}")
            return None
    
    def _find_edge_driver(self) -> Optional[str]:
        """Busca el driver de Edge en ubicaciones comunes"""
        possible_paths = [
            "msedgedriver.exe",  # En el directorio actual
            "drivers/msedgedriver.exe",
            "bin/msedgedriver.exe",
            os.path.join(self.config_manager.get_base_directory(), "msedgedriver.exe")
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        return None
    
    def _login_lbtr(self, driver, usuario: str, clave: str) -> bool:
        """Realiza el login en el sistema LBTR"""
        try:
            # Encontrar campos de login
            text_user = driver.find_element(By.NAME, "user")
            text_user.send_keys(usuario)
            
            text_password = driver.find_element(By.NAME, "password")
            text_password.send_keys(clave)
            
            # Hacer click en el botón de login
            submit_button = driver.find_element(By.ID, "btnSave")
            submit_button.click()
            
            time.sleep(2)
            
            # Verificar si el login fue exitoso
            return self._verificar_login_exitoso(driver)
            
        except Exception as e:
            self.logger.error(f"Error en login LBTR: {e}")
            return False
    
    def _verificar_login_exitoso(self, driver) -> bool:
        """Verifica si el login fue exitoso"""
        try:
            # Intentar encontrar mensaje de error
            try:
                mensaje_login = driver.find_element(By.XPATH, '//*[@id="principal-login"]/div[2]/div/form/div[1]/div[6]')
                mensaje_texto = mensaje_login.text
                
                if "Usuario y/o Contraseña incorrecta" in mensaje_texto:
                    messagebox.showerror("Error", "Usuario y/o clave incorrectos.")
                    return False
            except NoSuchElementException:
                # Si no encuentra el mensaje de error, el login fue exitoso
                pass
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error verificando login: {e}")
            return False
    
    def _navigate_to_transfers(self, driver) -> bool:
        """Navega al módulo de transferencias interbancarias"""
        try:
            # Localizar el menú principal
            menu = driver.find_element(By.ID, "dropdownMenu3")
            actions = ActionChains(driver)
            actions.move_to_element(menu).perform()
            
            driver.implicitly_wait(5)
            
            # Navegar a transferencia interbancaria
            submenu_transferencia = driver.find_element(By.XPATH, "//a[text()='Transferencia interbancaria']")
            actions.move_to_element(submenu_transferencia).perform()
            
            driver.implicitly_wait(5)
            time.sleep(1)
            
            # Hacer click en "Nuevo"
            nuevo = driver.find_element(By.LINK_TEXT, "Nuevo")
            actions.move_to_element(nuevo).perform()
            driver.implicitly_wait(2)
            nuevo.click()
            
            time.sleep(3)
            return True
            
        except Exception as e:
            self.logger.error(f"Error navegando a transferencias: {e}")
            return False
    
    def _procesar_transferencia_lbtr(self, driver, hoja, fila: int, obs_1: str, obs_2: str,
                                   beneficiario: str, cci: str, entidad_financiera: str,
                                   importe: float, ruc: str, domicilio: str) -> bool:
        """Procesa una transferencia LBTR individual"""
        try:
            # Seleccionar concepto
            concepto = Select(driver.find_element(By.ID, "selConcepto"))
            concepto.select_by_value("1")
            
            # Seleccionar entidad financiera
            valor_entidad = self._obtener_codigo_entidad(entidad_financiera)
            if valor_entidad is not None:
                entidad = Select(driver.find_element(By.ID, "selEntidad"))
                entidad.select_by_value(valor_entidad)
            else:
                hoja.range(f'K{fila}').value = "Error: entidad financiera no reconocida."
                return False
            
            # Ingresar monto
            text_monto = driver.find_element(By.NAME, "monto")
            text_monto.click()
            text_monto.clear()
            text_monto.send_keys(self.formatear_monto(importe))
            
            # Ingresar observación 1
            text_observacion = driver.find_element(By.NAME, "observacion")
            text_observacion.click()
            text_observacion.clear()
            text_observacion.send_keys(obs_1)
            
            # Ingresar observación ITF
            text_observacion_itf = driver.find_element(By.NAME, "observacionITF")
            text_observacion_itf.click()
            text_observacion_itf.clear()
            text_observacion_itf.send_keys(obs_2)
            
            # Ingresar CCI
            text_cci = driver.find_element(By.NAME, "numCuentaBen")
            text_cci.click()
            text_cci.clear()
            text_cci.send_keys(cci)
            
            # Ingresar beneficiario
            text_nombre = driver.find_element(By.NAME, "nombreBen")
            text_nombre.click()
            text_nombre.clear()
            text_nombre.send_keys(beneficiario)
            
            # Seleccionar tipo de documento (RUC)
            tipo_doc = Select(driver.find_element(By.ID, "selTipoDocBen"))
            tipo_doc.select_by_value("5")
            
            # Ingresar dirección
            text_direccion = driver.find_element(By.NAME, "direccionBen")
            text_direccion.click()
            text_direccion.clear()
            text_direccion.send_keys(domicilio)
            
            # Ingresar número de documento (RUC)
            text_doc = driver.find_element(By.NAME, "numDocumentoBen")
            text_doc.click()
            text_doc.clear()
            text_doc.send_keys(ruc)
            
            time.sleep(1)
            
            # Guardar transferencia
            guardar_button = driver.find_element(By.XPATH, "//button[@access='opcion.nuevointerbancaria.guardar']")
            guardar_button.click()
            
            # Esperar respuesta del modal
            wait = WebDriverWait(driver, 12)
            wait.until(EC.visibility_of_element_located((By.ID, "mdlMensajeInterbancaria")))
            
            # Obtener mensaje de respuesta
            mensaje_elemento = driver.find_element(By.XPATH, '//*[@id="mdlMensajeInterbancaria"]/div[2]/div/div[2]/div[1]')
            mensaje_texto = mensaje_elemento.text
            
            # Procesar respuesta
            if "La operación se realizó satisfactoriamente" in mensaje_texto:
                hoja.range(f'K{fila}').value = mensaje_texto
                resultado = True
            else:
                hoja.range(f'K{fila}').value = mensaje_texto
                resultado = False
            
            # Cerrar modal
            btn_mensaje = driver.find_element(By.XPATH, '//*[@id="mdlMensajeInterbancaria"]/div[2]/div/div[2]/div[2]/button')
            btn_mensaje.click()
            driver.implicitly_wait(4)
            
            if not resultado:
                # Si hay error, refrescar página para siguiente operación
                driver.refresh()
                time.sleep(5)
                # Re-navegar al módulo
                self._navigate_to_transfers(driver)
            
            return resultado
            
        except Exception as e:
            self.logger.error(f"Error procesando transferencia LBTR: {e}")
            return False
    
    def _obtener_codigo_entidad(self, entidad_financiera: str) -> Optional[str]:
        """Obtiene el código de entidad financiera"""
        entidad_upper = entidad_financiera.upper()
        for clave, valor in self.entidades_financieras.items():
            if clave in entidad_upper:
                return valor
        return None
    
    def _extract_after_colon(self, texto: str) -> str:
        """Extrae el texto después de los dos puntos"""
        if ':' in texto:
            return texto.split(':', 1)[1].strip()
        return texto.strip()
    
    def _extract_memo_number(self, obs_1: str, obs_2: str) -> Optional[str]:
        """Extrae el número de memo de las observaciones"""
        for obs in [obs_1, obs_2]:
            if "MEMO" in obs:
                palabras = obs.split()
                if len(palabras) > 1:
                    memo_completo = palabras[1]
                    if '-' in memo_completo:
                        return memo_completo.split('-')[0]
                    return memo_completo
        return None
    
    def _guardar_archivo_procesado_lbtr(self, wb, lista_memos: set, directorio: str, 
                                      fecha_actual: datetime) -> str:
        """Guarda el archivo procesado de LBTR con columna de total"""
        try:
            hoja_lbtr = wb.sheets['LBTR']
            
            # Insertar columna de total
            hoja_lbtr.range('I:I').insert(shift='right')
            hoja_lbtr.range('I1').value = "TOTAL"
            hoja_lbtr.range('I2').value = "=[@Importe]+14"  # Fórmula para sumar comisión
            
            # Generar nombre y ruta
            nombre_archivo = self.file_manager.generar_nombre_archivo_procesado(
                "LBTR", list(lista_memos)
            )
            
            ruta_directorio = self.file_manager.crear_directorio_procesados(
                directorio, "LBTR", fecha_actual
            )
            
            ruta_procesado = os.path.join(ruta_directorio, nombre_archivo)
            wb.save(ruta_procesado)
            
            self.logger.info(f"Archivo LBTR procesado guardado: {ruta_procesado}")
            return ruta_procesado
            
        except Exception as e:
            self.logger.error(f"Error guardando archivo procesado LBTR: {e}")
            return ""
    
    def execute_cargo_lbtr(self, ventana, archivo_xlc: str) -> bool:
        """
        Ejecuta el proceso de cargo LBTR desde historial
        
        Args:
            ventana: Ventana del emulador
            archivo_xlc: Ruta del archivo Excel con historial
        
        Returns:
            True si se completó correctamente
        """
        wb_lbtr = None
        book_lbtr = None
        
        try:
            # Definir estructura para archivo de historial
            dicc_tabla_historial = {
                'ID': str,
                'Cuenta_cargo': str,
                'OBS_1': str,
                'OBS_2': str,
                'Beneficiario': str,
                'CCI': str,
                'Entidad_Financiera': str,
                'Importe': float,
                'RUC': str,
                'DOMICILIO': str,
                'ESTADO': str
            }
            
            # Abrir archivo de historial
            book_lbtr = xw.App(visible=False)
            wb_lbtr = book_lbtr.books.open(archivo_xlc)
            hoja_lbtr = wb_lbtr.sheets['LBTR']
            
            # Leer datos
            tabla_lbtr = pd.read_excel(archivo_xlc, sheet_name='LBTR', header=0, dtype=dicc_tabla_historial)
            
            fecha_actual = self.get_fecha_actual()
            lista_memo_lbtr = set()
            cont_cargados = 0
            cont_no_cargados = 0
            
            self.logger.info(f"Procesando cargo LBTR desde: {archivo_xlc}")
            
            # Procesar cada fila exitosa
            for indice, fila in tabla_lbtr.iterrows():
                fila_lbtr = indice + 2
                
                cuenta = self.limpiar_numero_cuenta(str(fila['Cuenta_cargo']))
                obs_1 = str(fila['OBS_1']).strip()
                obs_2 = str(fila['OBS_2']).strip()
                importe = float(fila['Importe'])
                estado = fila['ESTADO']
                
                # Extraer número de memo
                titulo_memo = self._extract_memo_number(obs_1, obs_2)
                if titulo_memo:
                    if indice == 0:
                        lista_memo_lbtr.add(titulo_memo)
                    if titulo_memo not in lista_memo_lbtr:
                        lista_memo_lbtr.add(titulo_memo)
                
                # Procesar solo transferencias exitosas
                if "La operación se realizó satisfactoriamente" in str(estado):
                    resultado = self._procesar_cargo_lbtr_individual(
                        ventana, hoja_lbtr, fila_lbtr, cuenta, importe, titulo_memo, obs_1
                    )
                    
                    if resultado:
                        cont_cargados += 1
                    else:
                        cont_no_cargados += 1
                    
                    wb_lbtr.save()
            
            wb_lbtr.close()
            
            messagebox.showinfo(
                "FINALIZADO",
                f"Cargos terminados.\n"
                f"Exitosos: {cont_cargados}\n"
                f"Fallidos: {cont_no_cargados}\n"
                "Revisar excel."
            )
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error en cargo LBTR: {e}")
            messagebox.showerror("ERROR", f"Error procesando cargo LBTR: {e}")
            return False
        finally:
            if wb_lbtr and book_lbtr:
                try:
                    wb_lbtr.save()
                    wb_lbtr.close()
                    book_lbtr.quit()
                except:
                    pass
    
    def _procesar_cargo_lbtr_individual(self, ventana, hoja, fila: int, cuenta: str,
                                      importe: float, memorandum: str, obs_1: str) -> bool:
        """Procesa un cargo LBTR individual"""
        try:
            # Calcular monto total (importe + comisión de 14)
            monto_total = self.formatear_monto(importe + 14)
            
            # Activar ventana del emulador
            if not ventana.isActive:
                ventana.maximize()
                ventana.activate()
            
            # Limpiar clipboard
            pyperclip.copy('')
            pyautogui.press('f5')
            pyperclip.copy('')
            
            # Ejecutar secuencia de cargo
            pyautogui.write('042')  # Código de transacción
            time.sleep(self.intervalo)
            pyautogui.write(cuenta)  # Cuenta
            time.sleep(self.intervalo)
            pyautogui.write(monto_total)  # Importe total
            time.sleep(self.intervalo)
            pyautogui.press('Tab')
            time.sleep(self.intervalo)
            pyautogui.write(memorandum)  # Documento
            time.sleep(self.intervalo)
            pyautogui.press('Tab')
            time.sleep(self.intervalo)
            pyautogui.write('84')  # Motivo
            time.sleep(self.intervalo)
            
            # Glosa 1 - Memorándum
            if len(obs_1) <= 50:
                pyautogui.write(obs_1)
                pyautogui.press('Tab')
            else:
                pyautogui.write(obs_1[:50])
            
            # Glosa 2 - Detalle del importe
            glosa_importe = f'Importe {importe} comision S/14'
            pyautogui.write(glosa_importe)
            pyautogui.press('Enter')
            time.sleep(self.intervalo)
            
            # Capturar respuesta del emulador
            pyautogui.hotkey('ctrl', 'c')
            pyautogui.sleep(0.3)
            panel = pyperclip.paste()
            pyautogui.sleep(0.2)
            
            lineas = panel.splitlines()
            if len(lineas) > 23:
                msj_cargo = lineas[23].strip()
            else:
                msj_cargo = "Error: Respuesta incompleta del sistema"
            
            if "CORRECTOS" in msj_cargo:
                # Grabar operación
                pyautogui.sleep(0.2)
                pyautogui.press('f4')  # Grabar
                pyautogui.sleep(0.2)
                pyautogui.hotkey('ctrl', 'c')
                pyautogui.sleep(0.3)
                
                panel2 = pyperclip.paste()
                lineas2 = panel2.splitlines()
                
                if len(lineas2) > 23:
                    msj_cargo2 = lineas2[23].strip()
                else:
                    msj_cargo2 = "Error en grabación"
                
                hoja.range(f'K{fila}').value = msj_cargo2
                
                if "GRABACION CORRECTA" in msj_cargo2:
                    resultado = True
                else:
                    resultado = False
            else:
                hoja.range(f'K{fila}').value = msj_cargo
                resultado = False
            
            pyperclip.copy('')
            return resultado
            
        except Exception as e:
            self.logger.error(f"Error procesando cargo LBTR individual: {e}")
            return False
    
    @staticmethod
    def leer_xlc(ruta_xlc: str, memo: str, year: str, nro_cuenta: str, 
                posicion: int, glosa: str, limpiar: bool = False) -> bool:
        """
        Lee y procesa un archivo Excel de LBTR
        
        Args:
            ruta_xlc: Ruta del archivo a procesar
            memo: Número de memorándum
            year: Año del memorándum
            nro_cuenta: Número de cuenta
            posicion: Posición del memorándum (1 o 2)
            glosa: Glosa personalizada
            limpiar: Si debe limpiar el excel antes de procesar
        
        Returns:
            True si se procesó correctamente
        """
        config_manager = ConfigManager()
        logger = BaseLogic("LBTR_Reader")
        
        ruta_origen = ''
        finalizado = False
        
        try:
            # Obtener configuración
            ruta_origen, _ = config_manager.leer_json('LBTR')
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
            with xw.App(visible=True) as book:
                wb = book.books.open(ruta_origen)
                hoja = wb.sheets['LBTR']
                
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
                
                # Combinar datos válidos
                df = pd.concat(hojas_validas, ignore_index=True)
                
                # Eliminar transferencias del Banco Nación
                df = df.drop(df[df['Entidad Financiera'] == 'NACION'].index)
                
                # Limpiar excel si se solicita
                if limpiar:
                    ultima_fila2 = hoja.range('E' + str(hoja.cells.last_cell.row)).end('up').row
                    valor_e2 = hoja.range('E2').value
                    if ultima_fila2 == 2 or valor_e2:
                        hoja.range('2:90').delete()
                        hoja.range('K2').value = None
                        ultima_fila = 1
                
                # Encontrar última fila
                ultima_fila = hoja.range('E' + str(hoja.cells.last_cell.row)).end('up').row
                valor_b2 = hoja.range('D2').value
                if valor_b2 is None:
                    ultima_fila = 1
                
                es_lbtr = False
                
                # Procesar cada registro
                for indice, fila in df.iterrows():
                    numero = fila['N°']
                    beneficiario = fila['Beneficiario']
                    nro_cuenta_cci = fila['Nº  Cuenta']
                    tipo_cuenta = fila['Tipo de Cuenta']
                    entidad = fila['Entidad Financiera']
                    monto = fila['Monto (S/)']
                    
                    # Verificar si es registro válido
                    if pd.isna(beneficiario):
                        continue
                    elif numero and monto > 9999.99:
                        # Es una transferencia LBTR
                        ultima_fila += 1
                        es_lbtr = True
                        
                        hoja.range(f'B{ultima_fila}').value = nro_cuenta
                        beneficiario_limpio = BaseLogic("").limpiar_texto_beneficiario(str(beneficiario))
                        hoja.range(f'E{ultima_fila}').value = beneficiario_limpio
                        hoja.range(f'F{ultima_fila}').value = nro_cuenta_cci
                        hoja.range(f'G{ultima_fila}').value = entidad
                        hoja.range(f'H{ultima_fila}').value = monto
                        
                        # Configurar observaciones según posición
                        if posicion == 1:
                            hoja.range(f'C{ultima_fila}').value = f"MEMO {memo}-{year}-BN-7101 ANEXO {numero}"
                            hoja.range(f'D{ultima_fila}').value = glosa
                        else:
                            hoja.range(f'D{ultima_fila}').value = f"MEMO {memo}-{year}-BN-7101 ANEXO {numero}"
                            hoja.range(f'C{ultima_fila}').value = glosa
                    elif numero and monto < 10000:
                        es_lbtr = False
                    
                    # Procesar información adicional (RUC y DOMICILIO)
                    if es_lbtr and pd.isna(numero):
                        if "RUC" in str(beneficiario):
                            hoja.range(f'I{ultima_fila}').value = beneficiario
                        elif "DOMICILIO" in str(beneficiario):
                            hoja.range(f'J{ultima_fila}').value = beneficiario
                    
                    wb.save()
                    finalizado = True
                
                mensaje = "Revisar excel." if not hoja_no_util else 'Revisar excel. Algunas hojas se omitieron'
                messagebox.showinfo("Proceso terminado", mensaje)
                
                wb.save()
                wb.close()
                
            return True
            
        except FileNotFoundError:
            messagebox.showerror("Error", "Archivo movido o no encontrado")
            return False
        except Exception as ex:
            logger.logger.error(f"Error procesando archivo LBTR: {ex}")
            messagebox.showerror("Error", str(ex))
            return False
        finally:
            if not finalizado and ruta_origen:
                try:
                    os.startfile(ruta_origen)
                except:
                    pass