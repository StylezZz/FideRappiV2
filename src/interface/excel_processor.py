"""
Procesador de archivos Excel para FideRAPPI
Ventana para procesar y acomodar archivos Excel de memorándums
"""

import customtkinter
from tkinter import filedialog, messagebox
import os
import re
import datetime
from typing import Optional, Callable

from src.utils.logger import LoggerMixin
from src.utils.config_manager import ConfigManager


class ExcelProcessor(customtkinter.CTkToplevel, LoggerMixin):
    """Ventana para procesar archivos Excel"""
    
    def __init__(self, parent, tipo_operacion: str, config_manager: ConfigManager):
        """
        Inicializa el procesador de Excel
        
        Args:
            parent: Ventana padre
            tipo_operacion: Tipo de operación (CCE, AHORROS, etc.)
            config_manager: Gestor de configuración
        """
        super().__init__(parent)
        
        self.parent = parent
        self.tipo_operacion = tipo_operacion
        self.config_manager = config_manager
        
        # Variables de control
        self.memo_xlc = ""
        self.check_limpiar = None
        self.ent_nro_memo = None
        self.ent_nro_cuenta = None
        self.ent_year_memo = None
        self.combobox = None
        self.radio_var = customtkinter.IntVar(value=1)
        self.bar_carga = None
        self.btn_iniciar_limpia = None
        
        # Configurar ventana
        self.title(f"Procesar Excel - {self.tipo_operacion}")
        self.geometry("500x600")
        self.transient(parent)
        
        # Configurar protocolo de cierre
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Validaciones
        self.vcmd = self.register(self._validar_input_4_digitos)
        self.vcmd2 = self.register(self._validar_input_11_digitos)
        
        # Crear interfaz
        self._create_interface()
        
        # Bloquear ventana padre
        self._disable_parent_buttons()
        
        self.logger.info(f"Procesador de Excel abierto para {self.tipo_operacion}")
    
    def _create_interface(self):
        """Crea la interfaz según el tipo de operación"""
        # Frame principal
        self.main_frame = customtkinter.CTkFrame(self, bg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configurar grid
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(2, weight=1)
        
        if self.tipo_operacion == "LBTR":
            self._create_lbtr_interface()
        else:
            self._create_standard_interface()
    
    def _create_standard_interface(self):
        """Crea interfaz estándar para CCE, Ahorros y Cuentas Corrientes"""
        # Configurar filas
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(7, weight=1)
        
        # Título
        title_label = customtkinter.CTkLabel(
            self.main_frame,
            text=f"Procesar Excel - {self.tipo_operacion}",
            font=customtkinter.CTkFont(size=16, weight="bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))
        
        # Botón seleccionar archivo
        self.btn_buscar_memo = customtkinter.CTkButton(
            self.main_frame,
            text="Seleccionar memorándum",
            command=self._seleccionar_archivo
        )
        self.btn_buscar_memo.grid(row=1, column=0, columnspan=3, padx=15, pady=10, sticky="we")
        
        # Barra de progreso
        self.bar_carga = customtkinter.CTkProgressBar(
            self.main_frame,
            indeterminate_speed=1,
            mode='determinate'
        )
        self.bar_carga.grid(row=2, column=0, columnspan=3, padx=15, pady=10, sticky="we")
        self.bar_carga.set(0)
        
        # Número de memo
        memo_label = customtkinter.CTkLabel(self.main_frame, text="Ingresar n° de memo:")
        memo_label.grid(row=3, column=0, columnspan=3, sticky="s")
        
        # Frame para memo y año
        memo_frame = customtkinter.CTkFrame(self.main_frame, fg_color="transparent")
        memo_frame.grid(row=4, column=0, columnspan=3, pady=5)
        
        self.ent_nro_memo = customtkinter.CTkEntry(
            memo_frame,
            width=60,
            validate='key',
            validatecommand=(self.vcmd, '%P')
        )
        self.ent_nro_memo.pack(side="left", padx=5)
        
        guion_label = customtkinter.CTkLabel(memo_frame, text="-")
        guion_label.pack(side="left", padx=2)
        
        # Año actual por defecto
        año_actual = datetime.datetime.now().year
        self.ent_year_memo = customtkinter.CTkEntry(memo_frame, width=60)
        self.ent_year_memo.pack(side="left", padx=5)
        self.ent_year_memo.insert(0, str(año_actual))
        
        # Número de cuenta
        cuenta_label = customtkinter.CTkLabel(self.main_frame, text="Ingresar n° de cuenta:")
        cuenta_label.grid(row=5, column=0, columnspan=2, pady=(10, 5))
        
        self.ent_nro_cuenta = customtkinter.CTkEntry(
            self.main_frame,
            width=150,
            validate='key',
            validatecommand=(self.vcmd2, '%P')
        )
        self.ent_nro_cuenta.grid(row=6, column=0, pady=8)
        
        # Checkbox limpiar excel
        self.check_limpiar = customtkinter.CTkCheckBox(
            self.main_frame,
            text="Limpiar\nexcel"
        )
        self.check_limpiar.grid(row=6, column=1, sticky="e", padx=5)
        
        # Botón iniciar
        self.btn_iniciar_limpia = customtkinter.CTkButton(
            self.main_frame,
            text="Iniciar Procesamiento",
            command=self._procesar_archivo
        )
        self.btn_iniciar_limpia.grid(row=7, column=0, columnspan=3, padx=15, pady=10, sticky="we")
    
    def _create_lbtr_interface(self):
        """Crea interfaz específica para LBTR"""
        # Configurar filas
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(9, weight=1)
        
        # Lista de glosas predefinidas para LBTR
        self.lista_glosas = [
            "FIDEICOMISO DEL SISTEMA DE GESTION DE RESIDUOS SOLIDOS DE LA CIUDAD DE",
            "FIDEICOMISO GORE LORETO PLAN CIERRE DE BRECHAS - BN COMPONENTE GORE LORETO PCB",
            "P/O FIDEICOMISO LEY 30897-GORE LORETO",
            "P/O FIDEICOMISO MUNICIPALIDAD PROVINCIAL DE",
            "Comision de Confianza PIP 01 y PIP 03",
            "P/O FIDEICOMISO MUNICIPALIDAD DISTRITAL DE"
        ]
        
        # Título
        title_label = customtkinter.CTkLabel(
            self.main_frame,
            text="Procesar Excel - LBTR",
            font=customtkinter.CTkFont(size=16, weight="bold")
        )
        title_label.grid(row=0, column=0, columnspan=4, pady=(10, 20))
        
        # Botón seleccionar archivo
        self.btn_buscar_memo = customtkinter.CTkButton(
            self.main_frame,
            text="Seleccionar\nmemorándum",
            command=self._seleccionar_archivo
        )
        self.btn_buscar_memo.grid(row=1, column=0, columnspan=3, padx=15, pady=10)
        
        # Barra de progreso
        self.bar_carga = customtkinter.CTkProgressBar(
            self.main_frame,
            indeterminate_speed=1,
            mode='determinate'
        )
        self.bar_carga.grid(row=2, column=0, columnspan=3, padx=15, pady=10)
        self.bar_carga.set(0)
        
        # Número de memo
        memo_label = customtkinter.CTkLabel(self.main_frame, text="Ingresar n° de memo:")
        memo_label.grid(row=3, column=0, columnspan=3)
        
        # Frame para memo y año
        memo_frame = customtkinter.CTkFrame(self.main_frame, fg_color="transparent")
        memo_frame.grid(row=4, column=0, columnspan=3, pady=5)
        
        self.ent_nro_memo = customtkinter.CTkEntry(
            memo_frame,
            width=60,
            validate='key',
            validatecommand=(self.vcmd, '%P')
        )
        self.ent_nro_memo.pack(side="left", padx=5)
        
        guion_label = customtkinter.CTkLabel(memo_frame, text="-")
        guion_label.pack(side="left", padx=2)
        
        año_actual = datetime.datetime.now().year
        self.ent_year_memo = customtkinter.CTkEntry(memo_frame, width=60)
        self.ent_year_memo.pack(side="left", padx=5)
        self.ent_year_memo.insert(0, str(año_actual))
        
        # Número de cuenta
        cuenta_label = customtkinter.CTkLabel(self.main_frame, text="Ingresar n° de cuenta:")
        cuenta_label.grid(row=5, column=0, columnspan=3)
        
        self.ent_nro_cuenta = customtkinter.CTkEntry(
            self.main_frame,
            width=120,
            validate='key',
            validatecommand=(self.vcmd2, '%P')
        )
        self.ent_nro_cuenta.grid(row=6, column=0, columnspan=3, pady=8)
        
        # Checkbox limpiar excel
        self.check_limpiar = customtkinter.CTkCheckBox(
            self.main_frame,
            text="Limpiar\nexcel"
        )
        self.check_limpiar.grid(row=6, column=3, sticky="e", padx=5, pady=3)
        
        # Posición del memorándum (específico para LBTR)
        posicion_label = customtkinter.CTkLabel(
            self.main_frame,
            text="Posición del\nmemorándum:"
        )
        posicion_label.grid(row=1, column=3, padx=1, pady=1)
        
        self.radio_obs1 = customtkinter.CTkRadioButton(
            self.main_frame,
            text="Obs 1",
            variable=self.radio_var,
            value=1
        )
        self.radio_obs1.grid(row=3, column=3, padx=5, pady=3)
        
        self.radio_obs2 = customtkinter.CTkRadioButton(
            self.main_frame,
            text="Obs 2",
            variable=self.radio_var,
            value=2
        )
        self.radio_obs2.grid(row=4, column=3, padx=5, pady=3)
        
        # Selección de glosa
        glosa_label = customtkinter.CTkLabel(
            self.main_frame,
            text="Ingrese o seleccione la glosa:"
        )
        glosa_label.grid(row=7, column=0, columnspan=4, sticky="w", padx=2)
        
        self.combobox = customtkinter.CTkComboBox(
            self.main_frame,
            values=self.lista_glosas,
            width=450,
            font=customtkinter.CTkFont(family='Calibri', size=11)
        )
        self.combobox.set("--SELECCIONE O ESCRIBA--")
        self.combobox.grid(row=8, column=0, sticky="we", columnspan=4)
        
        # Bind para autocompletado
        self.combobox.bind("<KeyRelease>", self._auto_complete)
        
        # Botón iniciar
        self.btn_iniciar_limpia = customtkinter.CTkButton(
            self.main_frame,
            text="Iniciar Procesamiento",
            command=self._procesar_archivo
        )
        self.btn_iniciar_limpia.grid(row=9, column=0, columnspan=4, padx=15, pady=10, sticky="we")
    
    def _validar_input_4_digitos(self, value: str) -> bool:
        """Valida entrada de máximo 4 dígitos"""
        return (value.isdigit() or value == "") and len(value) <= 4
    
    def _validar_input_11_digitos(self, value: str) -> bool:
        """Valida entrada de máximo 11 dígitos"""
        return (value.isdigit() or value == "") and len(value) <= 11
    
    def _seleccionar_archivo(self):
        """Permite seleccionar el archivo de memorándum"""
        try:
            self.memo_xlc = filedialog.askopenfilename(
                title="Seleccione memorándum",
                filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
            )
            
            if self.memo_xlc:
                # Extraer nombre del archivo
                nombre_archivo = os.path.basename(self.memo_xlc)
                
                # Intentar extraer número de memo del nombre del archivo
                numero_memo = self._extraer_numero_memo(nombre_archivo)
                
                if numero_memo:
                    self.ent_nro_memo.delete(0, "end")
                    self.ent_nro_memo.insert(0, numero_memo)
                    self.bar_carga.set(1)
                else:
                    self.bar_carga.set(1)
                    
                self.logger.info(f"Archivo seleccionado: {self.memo_xlc}")
            else:
                self.ent_nro_memo.delete(0, "end")
                self.bar_carga.set(0)
                
        except Exception as e:
            self.logger.error(f"Error seleccionando archivo: {e}")
            messagebox.showerror("Error", f"Error seleccionando archivo: {e}")
    
    def _extraer_numero_memo(self, nombre_archivo: str) -> Optional[str]:
        """Extrae el número de memo del nombre del archivo"""
        try:
            # Limpiar caracteres especiales
            nombre_limpio = nombre_archivo.replace("N°", " ")
            
            # Buscar patrones de memo
            patrones = [
                r'(memo|MEMORANDO)[-\s]*(\d+)',
                r'MEMO[-\s]*(\d+)',
                r'memo[-\s]*(\d+)'
            ]
            
            for patron in patrones:
                coincidencias = re.search(patron, nombre_limpio, re.IGNORECASE)
                if coincidencias:
                    # Retornar el último grupo capturado (el número)
                    return coincidencias.group(coincidencias.lastindex)
            
            return None
            
        except Exception as e:
            self.logger.error(f"Error extrayendo número de memo: {e}")
            return None
    
    def _auto_complete(self, event):
        """Autocompletado para el combobox de glosas (solo LBTR)"""
        try:
            if self.tipo_operacion != "LBTR":
                return
            
            # Obtener valor tecleado
            valor_tecleado = self.combobox.get()
            
            # Filtrar la lista basándose en el valor tecleado
            opciones_filtradas = [
                item for item in self.lista_glosas 
                if valor_tecleado.lower() in item.lower()
            ]
            
            # Actualizar el combobox con las opciones filtradas
            self.combobox.configure(values=opciones_filtradas)
            
            # Si solo hay una coincidencia, seleccionarla automáticamente
            if len(opciones_filtradas) == 1:
                self.combobox.set(opciones_filtradas[0])
                
        except Exception as e:
            self.logger.error(f"Error en autocompletado: {e}")
    
    def _procesar_archivo(self):
        """Procesa el archivo Excel según el tipo de operación"""
        try:
            # Validar campos obligatorios
            if not self.memo_xlc:
                messagebox.showwarning("Advertencia", "Por favor seleccione un archivo")
                return
            
            if not self.ent_nro_memo.get():
                messagebox.showwarning("Advertencia", "Por favor ingrese el número de memo")
                return
            
            if not self.ent_nro_cuenta.get():
                messagebox.showwarning("Advertencia", "Por favor ingrese el número de cuenta")
                return
            
            # Deshabilitar botón durante procesamiento
            self.btn_iniciar_limpia.configure(state="disabled")
            
            # Obtener valores
            memo = self.ent_nro_memo.get()
            nro_cuenta = self.ent_nro_cuenta.get()
            year = self.ent_year_memo.get()
            limpiar = self.check_limpiar.get() == 1
            
            # Procesar según tipo de operación
            if self.tipo_operacion == "CCE":
                from src.operations.cce_operations import CCEOperations
                conca_memo = f'{memo}-{year}'
                resultado = CCEOperations.leer_xlc(self.memo_xlc, conca_memo, nro_cuenta, limpiar)
                
            elif self.tipo_operacion == "AHORROS":
                from src.operations.ahorros_operations import AhorrosOperations
                resultado = AhorrosOperations.leer_xlc(self.memo_xlc, memo, nro_cuenta, limpiar)
                
            elif self.tipo_operacion == "CTA_CTES":
                from src.operations.cte_operations import CTEOperations
                resultado = CTEOperations.leer_xlc(self.memo_xlc, memo, nro_cuenta, year, limpiar)
                
            elif self.tipo_operacion == "LBTR":
                from src.operations.lbtr_operations import LBTROperations
                posicion = self.radio_var.get()
                glosa = self.combobox.get()
                
                if glosa == "--SELECCIONE O ESCRIBA--":
                    messagebox.showwarning("Advertencia", "Por favor seleccione o escriba una glosa")
                    self.btn_iniciar_limpia.configure(state="normal")
                    return
                
                resultado = LBTROperations.leer_xlc(
                    self.memo_xlc, memo, year, nro_cuenta, posicion, glosa, limpiar
                )
            else:
                messagebox.showwarning("Error", "Tipo de operación no reconocido")
                self.btn_iniciar_limpia.configure(state="normal")
                return
            
            if resultado:
                self.logger.info(f"Archivo procesado exitosamente: {self.tipo_operacion}")
                self.on_closing()  # Cerrar ventana después del procesamiento exitoso
            else:
                self.logger.warning(f"Error procesando archivo: {self.tipo_operacion}")
            
        except Exception as e:
            self.logger.error(f"Error procesando archivo: {e}")
            messagebox.showerror("Error", f"Error procesando archivo: {e}")
        finally:
            # Rehabilitar botón
            if hasattr(self, 'btn_iniciar_limpia'):
                self.btn_iniciar_limpia.configure(state="normal")
    
    def _disable_parent_buttons(self):
        """Deshabilita botones en la ventana padre"""
        try:
            if hasattr(self.parent, 'operation_frames') and self.tipo_operacion.lower() in self.parent.operation_frames:
                frame = self.parent.operation_frames[self.tipo_operacion.lower()]
                for widget in frame.winfo_children():
                    if isinstance(widget, customtkinter.CTkButton):
                        widget.configure(state="disabled")
            
            # Deshabilitar botones de navegación
            if hasattr(self.parent, 'nav_buttons'):
                for btn in self.parent.nav_buttons.values():
                    btn.configure(state="disabled")
            
            if hasattr(self.parent, 'boton_salir'):
                self.parent.boton_salir.configure(state="disabled")
                
        except Exception as e:
            self.logger.error(f"Error deshabilitando botones padre: {e}")
    
    def _enable_parent_buttons(self):
        """Habilita botones en la ventana padre"""
        try:
            if hasattr(self.parent, 'operation_frames') and self.tipo_operacion.lower() in self.parent.operation_frames:
                frame = self.parent.operation_frames[self.tipo_operacion.lower()]
                for widget in frame.winfo_children():
                    if isinstance(widget, customtkinter.CTkButton):
                        widget.configure(state="normal")
            
            # Habilitar botones de navegación
            if hasattr(self.parent, 'nav_buttons'):
                for btn in self.parent.nav_buttons.values():
                    btn.configure(state="normal")
            
            if hasattr(self.parent, 'boton_salir'):
                self.parent.boton_salir.configure(state="normal")
                
        except Exception as e:
            self.logger.error(f"Error habilitando botones padre: {e}")
    
    def on_closing(self):
        """Maneja el cierre de la ventana"""
        try:
            self._enable_parent_buttons()
            self.logger.info("Procesador de Excel cerrado")
            self.destroy()
        except Exception as e:
            self.logger.error(f"Error cerrando procesador de Excel: {e}")
            self.destroy()