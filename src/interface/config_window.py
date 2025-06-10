"""
Ventana de configuración para FideRAPPI
Permite configurar rutas de archivos y enlaces
"""

import customtkinter
from tkinter import filedialog, messagebox
import os
from typing import Optional

from src.utils.logger import LoggerMixin
from src.utils.config_manager import ConfigManager


class ConfigWindow(customtkinter.CTkToplevel, LoggerMixin):
    """Ventana de configuración de la aplicación"""
    
    def __init__(self, parent, tipo_operacion: str, config_manager: ConfigManager):
        """
        Inicializa la ventana de configuración
        
        Args:
            parent: Ventana padre
            tipo_operacion: Tipo de operación a configurar
            config_manager: Gestor de configuración
        """
        super().__init__(parent)
        
        self.parent = parent
        self.tipo_operacion = tipo_operacion
        self.config_manager = config_manager
        
        # Configurar ventana
        self.title(f"Configuración {self.tipo_operacion}")
        self.geometry("600x400")
        self.transient(parent)
        
        # Obtener configuración actual
        try:
            self.ruta_origen, self.ruta_destino = self.config_manager.leer_json(self.tipo_operacion)
            if self.tipo_operacion == "LBTR":
                self.link_lbtr = self.config_manager.lbtr_credenciales()
        except Exception as e:
            self.logger.error(f"Error obteniendo configuración: {e}")
            self.ruta_origen = ""
            self.ruta_destino = ""
            self.link_lbtr = ""
        
        # Configurar protocolo de cierre
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Crear interfaz
        self._create_interface()
        
        # Bloquear ventana padre
        self._disable_parent_buttons()
        
        self.logger.info(f"Ventana de configuración abierta para {self.tipo_operacion}")
    
    def _create_interface(self):
        """Crea la interfaz de configuración"""
        # Frame principal
        self.main_frame = customtkinter.CTkFrame(self, bg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configurar grid
        self.main_frame.grid_columnconfigure(0, weight=1)
        if self.tipo_operacion == "LBTR":
            self.main_frame.grid_rowconfigure(6, weight=1)
        else:
            self.main_frame.grid_rowconfigure(4, weight=1)
        
        # Título
        title_label = customtkinter.CTkLabel(
            self.main_frame,
            text=f"Configuración de {self.tipo_operacion}",
            font=customtkinter.CTkFont(size=18, weight="bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))
        
        # Sección de archivo de origen
        self._create_file_origin_section()
        
        # Sección de ruta de destino
        self._create_destination_section()
        
        # Sección específica para LBTR
        if self.tipo_operacion == "LBTR":
            self._create_lbtr_section()
        
        # Botones de acción
        self._create_action_buttons()
    
    def _create_file_origin_section(self):
        """Crea la sección de configuración del archivo de origen"""
        # Label
        origin_label = customtkinter.CTkLabel(
            self.main_frame,
            text="Archivo de origen:",
            font=customtkinter.CTkFont(size=12, weight="bold")
        )
        origin_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 5))
        
        # Entry para mostrar ruta
        self.entry_ruta_archivo = customtkinter.CTkEntry(
            self.main_frame,
            corner_radius=5,
            width=450,
            state="readonly"
        )
        self.entry_ruta_archivo.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        
        # Insertar ruta actual
        if self.ruta_origen:
            self.entry_ruta_archivo.configure(state="normal")
            self.entry_ruta_archivo.delete(0, "end")
            self.entry_ruta_archivo.insert(0, self.ruta_origen)
            self.entry_ruta_archivo.xview_moveto(1.0)
            self.entry_ruta_archivo.configure(state="readonly")
        
        # Botón para cambiar archivo
        btn_change_file = customtkinter.CTkButton(
            self.main_frame,
            text="Cambiar",
            width=80,
            command=self._change_origin_file
        )
        btn_change_file.grid(row=2, column=1, padx=(10, 10), pady=5)
    
    def _create_destination_section(self):
        """Crea la sección de configuración de ruta de destino"""
        # Label
        dest_label = customtkinter.CTkLabel(
            self.main_frame,
            text="Carpeta de guardado:",
            font=customtkinter.CTkFont(size=12, weight="bold")
        )
        dest_label.grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=(15, 5))
        
        # Entry para mostrar ruta
        self.entry_dest_archivo = customtkinter.CTkEntry(
            self.main_frame,
            corner_radius=5,
            width=450,
            state="readonly"
        )
        self.entry_dest_archivo.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
        
        # Insertar ruta actual
        if self.ruta_destino:
            self.entry_dest_archivo.configure(state="normal")
            self.entry_dest_archivo.delete(0, "end")
            self.entry_dest_archivo.insert(0, self.ruta_destino)
            self.entry_dest_archivo.xview_moveto(1.0)
            self.entry_dest_archivo.configure(state="readonly")
        
        # Botón para cambiar carpeta
        btn_change_dest = customtkinter.CTkButton(
            self.main_frame,
            text="Cambiar",
            width=80,
            command=self._change_destination_folder
        )
        btn_change_dest.grid(row=4, column=1, padx=(10, 10), pady=5)
    
    def _create_lbtr_section(self):
        """Crea la sección específica para configuración de LBTR"""
        # Label
        lbtr_label = customtkinter.CTkLabel(
            self.main_frame,
            text="Enlace a LBTR:",
            font=customtkinter.CTkFont(size=12, weight="bold")
        )
        lbtr_label.grid(row=5, column=0, columnspan=2, sticky="w", padx=10, pady=(15, 5))
        
        # Entry para enlace
        self.entry_link = customtkinter.CTkEntry(
            self.main_frame,
            corner_radius=5,
            width=450
        )
        self.entry_link.grid(row=6, column=0, sticky="ew", padx=10, pady=5)
        
        # Insertar enlace actual
        if hasattr(self, 'link_lbtr') and self.link_lbtr:
            self.entry_link.insert(0, self.link_lbtr)
            self.entry_link.xview_moveto(1.0)
        
        # Botón para guardar enlace
        btn_save_link = customtkinter.CTkButton(
            self.main_frame,
            text="Guardar",
            width=80,
            command=self._save_lbtr_link
        )
        btn_save_link.grid(row=6, column=1, padx=(10, 10), pady=5)
    
    def _create_action_buttons(self):
        """Crea los botones de acción"""
        # Frame para botones
        button_frame = customtkinter.CTkFrame(self.main_frame, fg_color="transparent")
        if self.tipo_operacion == "LBTR":
            button_frame.grid(row=7, column=0, columnspan=2, pady=(20, 10))
        else:
            button_frame.grid(row=5, column=0, columnspan=2, pady=(20, 10))
        
        # Botón cerrar
        btn_close = customtkinter.CTkButton(
            button_frame,
            text="Cerrar",
            width=120,
            fg_color="gray",
            command=self.on_closing
        )
        btn_close.pack(side="right", padx=10)
        
        # Botón aplicar (si hay cambios pendientes)
        btn_apply = customtkinter.CTkButton(
            button_frame,
            text="Aplicar cambios",
            width=120,
            command=self._apply_changes
        )
        btn_apply.pack(side="right", padx=10)
    
    def _change_origin_file(self):
        """Cambia el archivo de origen"""
        try:
            # Tipos de archivo según operación
            if self.tipo_operacion in ["CCE", "AHORROS", "CTA_CTES", "LBTR"]:
                filetypes = [
                    ("Archivos Excel con macros", "*.xlsm"),
                    ("Archivos Excel", "*.xlsx"),
                    ("Todos los archivos", "*.*")
                ]
            else:
                filetypes = [
                    ("Archivos Excel", "*.xlsx"),
                    ("Archivos Excel con macros", "*.xlsm"),
                    ("Todos los archivos", "*.*")
                ]
            
            # Solicitar archivo
            nueva_ruta = filedialog.askopenfilename(
                title="Seleccionar archivo de origen",
                filetypes=filetypes,
                initialdir=os.path.dirname(self.ruta_origen) if self.ruta_origen else None
            )
            
            if nueva_ruta:
                # Validar archivo
                if self._validate_excel_file(nueva_ruta):
                    # Actualizar configuración
                    if self.config_manager.modificar_json(self.tipo_operacion, nueva_ruta):
                        # Actualizar interfaz
                        self.entry_ruta_archivo.configure(state="normal")
                        self.entry_ruta_archivo.delete(0, "end")
                        self.entry_ruta_archivo.insert(0, nueva_ruta)
                        self.entry_ruta_archivo.xview_moveto(1.0)
                        self.entry_ruta_archivo.configure(state="readonly")
                        
                        # Actualizar ruta de destino automáticamente
                        directorio = os.path.dirname(nueva_ruta)
                        self._update_destination_path(directorio)
                        
                        self.logger.info(f"Archivo de origen actualizado: {nueva_ruta}")
                        
        except Exception as e:
            self.logger.error(f"Error cambiando archivo de origen: {e}")
            messagebox.showerror("Error", f"Error al cambiar archivo: {e}")
    
    def _change_destination_folder(self):
        """Cambia la carpeta de destino"""
        try:
            nueva_carpeta = filedialog.askdirectory(
                title="Seleccionar carpeta de destino",
                initialdir=self.ruta_destino if self.ruta_destino else None
            )
            
            if nueva_carpeta:
                self._update_destination_path(nueva_carpeta)
                
        except Exception as e:
            self.logger.error(f"Error cambiando carpeta de destino: {e}")
            messagebox.showerror("Error", f"Error al cambiar carpeta: {e}")
    
    def _update_destination_path(self, nueva_ruta: str):
        """Actualiza la ruta de destino"""
        try:
            if self.config_manager.modificar_ruta_final_json(self.tipo_operacion, nueva_ruta):
                # Actualizar interfaz
                self.entry_dest_archivo.configure(state="normal")
                self.entry_dest_archivo.delete(0, "end")
                self.entry_dest_archivo.insert(0, nueva_ruta)
                self.entry_dest_archivo.xview_moveto(1.0)
                self.entry_dest_archivo.configure(state="readonly")
                
                self.logger.info(f"Ruta de destino actualizada: {nueva_ruta}")
                
        except Exception as e:
            self.logger.error(f"Error actualizando ruta de destino: {e}")
            messagebox.showerror("Error", f"Error al actualizar ruta: {e}")
    
    def _save_lbtr_link(self):
        """Guarda el enlace de LBTR"""
        try:
            nuevo_link = self.entry_link.get().strip()
            if nuevo_link:
                if self.config_manager.save_link_lbtr(nuevo_link):
                    self.logger.info(f"Enlace LBTR actualizado: {nuevo_link}")
            else:
                messagebox.showwarning("Advertencia", "Por favor ingrese un enlace válido")
                
        except Exception as e:
            self.logger.error(f"Error guardando enlace LBTR: {e}")
            messagebox.showerror("Error", f"Error al guardar enlace: {e}")
    
    def _validate_excel_file(self, ruta_archivo: str) -> bool:
        """
        Valida que el archivo Excel tenga el formato correcto
        
        Args:
            ruta_archivo: Ruta del archivo a validar
        
        Returns:
            True si es válido
        """
        try:
            import xlwings as xw
            
            # Validar que el archivo existe
            if not os.path.exists(ruta_archivo):
                messagebox.showerror("Error", "El archivo seleccionado no existe")
                return False
            
            # Validar formato específico según tipo de operación
            cabeceras_esperadas = {
                "CCE": ("CCE", ["ID", "Memorandum", "Cuenta", "Beneficiario", "CCI", "Monto", "IB", "BN", "COMENTARIO", "MENSAJE_EMULADOR"]),
                "CTA_CTES": ("Corriente", ["ID", "Memorandum", "Cta_cargo", "Cta_abono", "Monto", "Glosa", "Comision", "ITF", "Observacion", "Mensaje_cargo", "Mensaje_abono"]),
                "AHORROS": ("Ahorros", ["ID", "Memo", "Cuenta_cargo", "Beneficiario", "Cuenta_abono", "Monto", "ITF", "Msj_abono", "Beneficiario_final", "Secuencia", "Estado"]),
                "Cargo": ("Cargo", ["Id", "COD", "Cuenta", "Importe", "Memorandum", "Motivo", "Glosa1", "Glosa2", "Glosa3", "Mensaje_emulacion", "Observacion"]),
                "LBTR": ("LBTR", ["ID", "Cuenta_cargo", "OBS_1", "OBS_2", "Beneficiario", "CCI", "Entidad_Financiera", "Importe", "RUC", "DOMICILIO", "ESTADO"])
            }
            
            if self.tipo_operacion not in cabeceras_esperadas:
                # No validar tipos no definidos
                return True
            
            hoja_esperada, cabecera_esperada = cabeceras_esperadas[self.tipo_operacion]
            
            # Abrir archivo para validación
            app = xw.App(visible=False)
            try:
                book = app.books.open(ruta_archivo)
                
                # Verificar que existe la hoja
                if hoja_esperada not in [sheet.name for sheet in book.sheets]:
                    messagebox.showerror("Error", f"La hoja '{hoja_esperada}' no está presente en el archivo")
                    return False
                
                sheet = book.sheets[hoja_esperada]
                
                # Obtener primera fila
                primera_fila = sheet.range("1:1").options(ndim=1, expand='right').value[:len(cabecera_esperada)]
                
                # Verificar cabecera
                if primera_fila != cabecera_esperada:
                    messagebox.showwarning(
                        "Advertencia",
                        f"La cabecera del archivo no coincide exactamente con el formato esperado.\n"
                        f"El archivo podría funcionar, pero se recomienda verificar el formato."
                    )
                
                book.close()
                return True
                
            finally:
                app.quit()
                
        except Exception as e:
            self.logger.error(f"Error validando archivo Excel: {e}")
            messagebox.showerror("Error", f"Error validando archivo: {e}")
            return False
    
    def _apply_changes(self):
        """Aplica cambios pendientes"""
        try:
            # Actualizar configuración del padre
            if hasattr(self.parent, 'ruta_origen'):
                self.parent.ruta_origen, self.parent.ruta_destino = self.config_manager.leer_json(self.tipo_operacion)
            
            messagebox.showinfo("Información", "Cambios aplicados correctamente")
            
        except Exception as e:
            self.logger.error(f"Error aplicando cambios: {e}")
            messagebox.showerror("Error", f"Error aplicando cambios: {e}")
    
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
            self.logger.info("Ventana de configuración cerrada")
            self.destroy()
        except Exception as e:
            self.logger.error(f"Error cerrando ventana de configuración: {e}")
            self.destroy()