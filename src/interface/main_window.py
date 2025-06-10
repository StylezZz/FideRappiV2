"""
Ventana principal de FideRAPPI
Interfaz de usuario principal usando CustomTkinter
"""

import customtkinter
from customtkinter import CTk, CTkRadioButton, IntVar, CTkLabel
import re
import os
import sys
import threading
import pyautogui
import pyperclip
from PIL import Image
import datetime
from tkinter import messagebox, filedialog
import traceback

from src.utils.config_manager import ConfigManager
from src.utils.logger import LoggerMixin
from src.interface.config_window import ConfigWindow
from src.interface.excel_processor import ExcelProcessor
from src.interface.operation_validator import OperationValidator
#from src.operations import cce_operations, ahorros_operations, cte_operations
from src.operations import lbtr_operations, cargo_operations, extra_operations

# Configurar tema de CustomTkinter
customtkinter.set_default_color_theme("blue")
customtkinter.set_appearance_mode("dark")


class FideRappiApp(customtkinter.CTk, LoggerMixin):
    """Ventana principal de la aplicación FideRAPPI"""
    
    def __init__(self):
        super().__init__()
        
        # Configuración inicial
        self.title("FideRAPPI v2.0")
        self.geometry("390x410+720+120")
        
        # Configurar icono si está disponible
        try:
            self.iconbitmap(self._get_resource_path('assets/logo-banco-nacion.ico'))
        except Exception:
            self.logger.warning("No se pudo cargar el icono de la aplicación")
        
        # Inicializar componentes
        self.config_manager = ConfigManager()
        self.tipo_operacion = "CCE"  # Operación por defecto
        self.operation_validator = OperationValidator(self.tipo_operacion)
        
        # Variables de control
        self.tipo_operacion = "CCE"
        self.radio_var = IntVar(value=1)
        self.vcmd = self.register(self._validar_input_4_digitos)
        self.vcmd2 = self.register(self._validar_input_11_digitos)
        
        # Configurar layout
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        # Cargar imágenes
        self._load_images()
        
        # Crear interfaz
        self._create_navigation_frame()
        self._create_operation_frames()
        
        # Seleccionar frame por defecto
        self.select_frame_by_name("CCE")
        self.button_frame_cce()
        
        self.logger.info("Ventana principal inicializada correctamente")
    
    def _get_resource_path(self, relative_path: str) -> str:
        """Obtiene la ruta de un recurso de forma portable"""
        try:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)
        except Exception:
            # Fallback al directorio actual
            return relative_path
    
    def _load_images(self):
        """Carga las imágenes necesarias para la interfaz"""
        try:
            image_path = self._get_resource_path("assets/img")
            
            # Definir imágenes con tamaños
            images_config = {
                'imagen_config': ('configuraciones.png', (26, 26)),
                'imagen_bank': ('banco.png', (45, 45)),
                'img_save': ('salvar.png', (25, 25)),
                'img_exit': ('cerrar-sesion.png', (25, 25)),
                'img_play': ('play.png', (25, 25)),
                'img_abonar': ('depositar.png', (30, 30)),
                'img_depositar': ('retirar.png', (30, 30)),
                'img_transa': ('historial-de-transacciones.png', (30, 30)),
                'img_lbtr_abonar': ('lbtr_abonos.png', (30, 30)),
                'img_enlace': ('enlace-web.png', (25, 25)),
                'img_en_progreso': ('trabajo-en-progreso.png', (125, 125))
            }
            
            # Cargar cada imagen
            for attr_name, (filename, size) in images_config.items():
                try:
                    img_path = os.path.join(image_path, filename)
                    if os.path.exists(img_path):
                        setattr(self, attr_name, customtkinter.CTkImage(
                            Image.open(img_path), size=size
                        ))
                    else:
                        # Crear imagen placeholder si no existe
                        setattr(self, attr_name, None)
                        self.logger.warning(f"Imagen no encontrada: {img_path}")
                except Exception as e:
                    setattr(self, attr_name, None)
                    self.logger.error(f"Error cargando imagen {filename}: {e}")
                    
        except Exception as e:
            self.logger.error(f"Error general cargando imágenes: {e}")
            # Inicializar todas las imágenes como None
            for attr_name in ['imagen_config', 'imagen_bank', 'img_save', 'img_exit',
                            'img_play', 'img_abonar', 'img_depositar', 'img_transa',
                            'img_lbtr_abonar', 'img_enlace', 'img_en_progreso']:
                setattr(self, attr_name, None)
    
    def _create_navigation_frame(self):
        """Crea el frame de navegación lateral"""
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(7, weight=1)
        
        # Título con logo
        self.nav_frame_title = customtkinter.CTkLabel(
            self.navigation_frame, 
            text="  FideRAPPI v2.0", 
            image=self.imagen_bank,
            compound="left", 
            font=customtkinter.CTkFont(size=15, weight="bold")
        )
        self.nav_frame_title.grid(row=0, column=0, padx=20, pady=20)
        
        # Botones de navegación
        nav_buttons = [
            ("CCE", self.button_frame_cce),
            ("Ahorros", self.button_frame_ahorros),
            ("Corrientes", self.button_frame_ctes),
            ("LBTR", self.button_frame_lbtr),
            ("Cargo", self.button_frame_cargo),
            ("Extras", self.button_frame_xtras)
        ]
        
        self.nav_buttons = {}
        for i, (text, command) in enumerate(nav_buttons, 1):
            btn = customtkinter.CTkButton(
                self.navigation_frame,
                corner_radius=0,
                height=40,
                border_spacing=10,
                text=text,
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray70", "gray30"),
                image=None,
                anchor="w",
                command=command
            )
            btn.grid(row=i, column=0, sticky="ew")
            self.nav_buttons[text.lower().replace(" ", "_")] = btn
        
        # Botón salir
        self.boton_salir = customtkinter.CTkButton(
            self.navigation_frame,
            text="",
            image=self.img_exit,
            command=self.salir_aplicacion,
            width=45,
            height=45
        )
        self.boton_salir.grid(row=7, column=0, padx=5, pady=5, sticky="ws")
    
    def _create_operation_frames(self):
        """Crea los frames para cada tipo de operación"""
        self.operation_frames = {}
        
        # Configuración de operaciones
        operations_config = {
            'CCE': {
                'buttons': [
                    ('Acomodar excel', None, self.arreglar),
                    ('Abonar CCE', 'img_abonar', self.validacion_host),
                    ('Cargar', 'img_depositar', self.validacion_host_cargo),
                    ('Ver historial', 'img_transa', lambda: self.abrir_historial('CCE'))
                ]
            },
            'Ahorros': {
                'buttons': [
                    ('Acomodar excel', None, self.arreglar),
                    ('Abonar\nahorros', 'img_abonar', self.validacion_host),
                    ('Cargar', 'img_depositar', self.validacion_host_cargo),
                    ('Ver historial', 'img_transa', lambda: self.abrir_historial('AHORROS'))
                ]
            },
            'Corrientes': {
                'buttons': [
                    ('Acomodar excel', None, self.arreglar),
                    ('Abonar Y cargar', 'img_abonar', self.validacion_host),
                    ('Ver historial', 'img_transa', lambda: self.abrir_historial('CTA_CTES'))
                ]
            },
            'LBTR': {
                'buttons': [
                    ('Acomodar excel', None, self.arreglar),
                    ('Abonar LBTR', 'img_lbtr_abonar', self.iniciar_sesion_lbtr),
                    ('Cargar 1x1', 'img_depositar', self.validacion_host_cargo)
                ]
            },
            'Cargo': {
                'buttons': [
                    ('Cargar 1X1', 'img_abonar', self.validacion_host),
                    ('Historial', 'img_transa', None)
                ]
            },
            'Extras': {
                'buttons': [
                    ('UNIR PDF\'s', None, self.exec_union)
                ]
            }
        }
        
        # Crear frames para cada operación
        for op_name, config in operations_config.items():
            frame = self._create_operation_frame(op_name, config['buttons'])
            self.operation_frames[op_name.lower()] = frame
    
    def _create_operation_frame(self, operation_name: str, buttons_config: list) -> customtkinter.CTkFrame:
        """Crea un frame para una operación específica"""
        frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        
        # Configurar grid
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_rowconfigure(len(buttons_config) + 1, weight=1)
        
        # Crear botones
        for i, (text, image_attr, command) in enumerate(buttons_config, 1):
            image = getattr(self, image_attr) if image_attr else None
            
            btn = customtkinter.CTkButton(
                frame,
                text=text,
                image=image,
                width=75,
                height=45,
                anchor="center",
                command=command,
                font=("Berlin Sans FB", 17)
            )
            
            # Configurar grid según el tipo de operación
            if operation_name == "Extras":
                btn.grid(row=i-1, column=0, columnspan=2, padx=20, pady=20, sticky="nwse")
            elif operation_name == "Cargo" and i == len(buttons_config):
                # Botón de configuración para cargo
                btn.grid(row=i, column=0, columnspan=2, padx=20, pady=20, sticky="we")
            else:
                btn.grid(row=i, column=0, columnspan=2, padx=20, pady=20, sticky="we")
        
        # Añadir botón de configuración (excepto para Extras)
        if operation_name != "Extras":
            config_row = len(buttons_config) + 1
            btn_conf = customtkinter.CTkButton(
                frame,
                text="",
                image=self.imagen_config,
                width=45,
                height=45,
                command=self.configurar
            )
            btn_conf.grid(row=config_row, column=1, padx=15, pady=10, sticky="es")
        
        return frame
    
    def _validar_input_4_digitos(self, value: str) -> bool:
        """Valida entrada de máximo 4 dígitos"""
        return value.isdigit() or value == "" and len(value) <= 4
    
    def _validar_input_11_digitos(self, value: str) -> bool:
        """Valida entrada de máximo 11 dígitos"""
        return value.isdigit() or value == "" and len(value) <= 11
    
    def salir_aplicacion(self):
        """Cierra la aplicación"""
        try:
            self.logger.info("Cerrando aplicación")
            self.destroy()
        except Exception as e:
            self.logger.error(f"Error al cerrar aplicación: {e}")
    
    def select_frame_by_name(self, name: str):
        """Selecciona y muestra el frame de una operación específica"""
        # Resetear colores de botones
        for btn in self.nav_buttons.values():
            btn.configure(fg_color="transparent")
        
        # Activar botón seleccionado
        button_map = {
            "CCE": "cce",
            "Ahorros": "ahorros", 
            "Corrientes": "corrientes",
            "Lbtr": "lbtr",
            "Cargo": "cargo",
            "Extras": "extras"
        }
        
        if name in button_map and button_map[name] in self.nav_buttons:
            self.nav_buttons[button_map[name]].configure(fg_color=("gray75", "gray25"))
        
        # Ocultar todos los frames
        for frame in self.operation_frames.values():
            frame.grid_forget()
        
        # Mostrar frame seleccionado
        frame_name = name.lower()
        if frame_name in self.operation_frames:
            self.operation_frames[frame_name].grid(row=0, column=1, sticky="nsew")
    
    def button_frame(self, tipo_operacion: str, frame_name: str):
        """Configura el contexto para un tipo de operación"""
        try:
            self.tipo_operacion = tipo_operacion
            self.ruta_origen, self.ruta_destino = self.config_manager.leer_json(tipo_operacion)
            self.select_frame_by_name(frame_name)
            
            if tipo_operacion == "LBTR":
                self.link = self.config_manager.lbtr_credenciales()
            
            self.logger.info(f"Contexto configurado para: {tipo_operacion}")
        except Exception as ex:
            self.logger.error(f"Error configurando contexto para {tipo_operacion}: {ex}")
            messagebox.showerror("Error", f"Error configurando operación: {ex}")
    
    # Métodos para cambiar entre operaciones
    def button_frame_cce(self):
        self.button_frame("CCE", "CCE")
    
    def button_frame_ahorros(self):
        self.button_frame("AHORROS", "Ahorros")
    
    def button_frame_ctes(self):
        self.button_frame("CTA_CTES", "Corrientes")
    
    def button_frame_lbtr(self):
        self.button_frame("LBTR", "Lbtr")
    
    def button_frame_cargo(self):
        self.button_frame("Cargo", "Cargo")
    
    def button_frame_xtras(self):
        self.select_frame_by_name("Extras")
    
    def configurar(self):
        """Abre la ventana de configuración"""
        try:
            config_window = ConfigWindow(self, self.tipo_operacion, self.config_manager)
            config_window.grab_set()
        except Exception as e:
            self.logger.error(f"Error abriendo configuración: {e}")
            messagebox.showerror("Error", f"Error abriendo configuración: {e}")
    
    def arreglar(self):
        """Abre el procesador de Excel"""
        try:
            excel_window = ExcelProcessor(self, self.tipo_operacion, self.config_manager)
            excel_window.grab_set()
        except Exception as e:
            self.logger.error(f"Error abriendo procesador de Excel: {e}")
            messagebox.showerror("Error", f"Error abriendo procesador: {e}")
    
    def validacion_host(self):
        """Valida y ejecuta operaciones en el host"""
        try:
            validator = OperationValidator(self.tipo_operacion)
            validator.validar_y_ejecutar_operacion(False)
        except Exception as e:
            self.logger.error(f"Error en validación de host: {e}")
            messagebox.showerror("Error", f"Error en operación: {e}")
    
    def validacion_host_cargo(self):
        """Valida y ejecuta operaciones de cargo en el host"""
        try:
            validator = OperationValidator(self.tipo_operacion)
            validator.validar_y_ejecutar_operacion(True)
        except Exception as e:
            self.logger.error(f"Error en validación de cargo: {e}")
            messagebox.showerror("Error", f"Error en operación de cargo: {e}")
    
    def iniciar_sesion_lbtr(self):
        """Inicia el proceso de login para LBTR"""
        try:
            from src.interface.lbtr_login import LBTRLoginWindow
            login_window = LBTRLoginWindow(self)
            login_window.grab_set()
        except Exception as e:
            self.logger.error(f"Error iniciando sesión LBTR: {e}")
            messagebox.showerror("Error", f"Error iniciando LBTR: {e}")
    
    def exec_union(self):
        """Ejecuta la unión de PDFs"""
        try:
            carpeta = filedialog.askdirectory(title="Escoja la carpeta con PDFs")
            if carpeta:
                extra_operations.ExtraOperations.combinar_pdfs(carpeta)
        except Exception as e:
            self.logger.error(f"Error uniendo PDFs: {e}")
            messagebox.showerror("Error", f"Error uniendo PDFs: {e}")
    
    def abrir_historial(self, tipo_operacion: str):
        """Abre el historial de una operación"""
        try:
            from src.utils.file_manager import FileManager
            file_manager = FileManager()
            file_manager.abrir_historial(tipo_operacion, self.config_manager)
        except Exception as e:
            self.logger.error(f"Error abriendo historial: {e}")
            messagebox.showerror("Error", f"Error abriendo historial: {e}")