"""
Ventana de login para LBTR
Permite al usuario ingresar credenciales para el sistema LBTR
"""

import customtkinter
from tkinter import messagebox
import threading

from src.utils.logger import LoggerMixin


class LBTRLoginWindow(customtkinter.CTkToplevel, LoggerMixin):
    """Ventana de login para operaciones LBTR"""
    
    def __init__(self, parent):
        """
        Inicializa la ventana de login LBTR
        
        Args:
            parent: Ventana padre
        """
        super().__init__(parent)
        
        self.parent = parent
        
        # Variables de control
        self.entry_user = None
        self.entry_pass = None
        self.check_mostrar_clave = None
        self.btn_iniciar_lbtr = None
        self.btn_salir_login = None
        
        # Configurar ventana
        self.title("Login LBTR")
        self.geometry("400x300")
        self.transient(parent)
        self.resizable(False, False)
        
        # Configurar protocolo de cierre
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Crear interfaz
        self._create_interface()
        
        # Bloquear ventana padre
        self._disable_parent_buttons()
        
        # Enfocar en el campo de usuario
        self.after(100, lambda: self.entry_user.focus())
        
        self.logger.info("Ventana de login LBTR abierta")
    
    def _create_interface(self):
        """Crea la interfaz de login"""
        # Frame principal
        self.main_frame = customtkinter.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Configurar grid
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=0)
        for i in range(6):
            self.main_frame.grid_rowconfigure(i, weight=1)
        
        # Título
        title_label = customtkinter.CTkLabel(
            self.main_frame,
            text="Iniciar Sesión LBTR",
            font=customtkinter.CTkFont(size=18, weight="bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(10, 20))
        
        # Campo Usuario
        user_label = customtkinter.CTkLabel(self.main_frame, text="Usuario:")
        user_label.grid(row=1, column=0, columnspan=2, pady=(5, 0), padx=5, sticky="w")
        
        self.entry_user = customtkinter.CTkEntry(
            self.main_frame,
            placeholder_text="Ingrese su usuario"
        )
        self.entry_user.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        
        # Campo Contraseña
        pass_label = customtkinter.CTkLabel(self.main_frame, text="Contraseña:")
        pass_label.grid(row=3, column=0, columnspan=2, pady=(10, 0), padx=5, sticky="w")
        
        self.entry_pass = customtkinter.CTkEntry(
            self.main_frame,
            placeholder_text="Ingrese su contraseña",
            show="*"
        )
        self.entry_pass.grid(row=4, column=0, padx=10, pady=5, sticky="ew")
        
        # Checkbox para mostrar/ocultar contraseña
        var_mostrar_clave = customtkinter.IntVar()
        
        def toggle_password():
            """Alterna la visibilidad de la contraseña"""
            if var_mostrar_clave.get():
                self.entry_pass.configure(show='')
            else:
                self.entry_pass.configure(show='*')
        
        self.check_mostrar_clave = customtkinter.CTkCheckBox(
            self.main_frame,
            text="Mostrar",
            variable=var_mostrar_clave,
            command=toggle_password,
            width=80
        )
        self.check_mostrar_clave.grid(row=4, column=1, padx=10, pady=5)
        
        # Frame para botones
        button_frame = customtkinter.CTkFrame(self.main_frame, fg_color="transparent")
        button_frame.grid(row=5, column=0, columnspan=2, pady=(20, 10), sticky="ew")
        
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        
        # Botón Iniciar
        self.btn_iniciar_lbtr = customtkinter.CTkButton(
            button_frame,
            text="Iniciar Sesión",
            command=self._iniciar_lbtr,
            width=120
        )
        self.btn_iniciar_lbtr.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        # Botón Cancelar
        self.btn_salir_login = customtkinter.CTkButton(
            button_frame,
            text="Cancelar",
            command=self.on_closing,
            fg_color="gray",
            width=120
        )
        self.btn_salir_login.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        # Bind Enter key para iniciar sesión
        self.bind('<Return>', lambda event: self._iniciar_lbtr())
        self.entry_user.bind('<Return>', lambda event: self.entry_pass.focus())
        self.entry_pass.bind('<Return>', lambda event: self._iniciar_lbtr())
        
        # Información adicional
        info_label = customtkinter.CTkLabel(
            self.main_frame,
            text="Ingrese sus credenciales del sistema LBTR\nEste proceso puede tomar varios minutos",
            font=customtkinter.CTkFont(size=11),
            text_color="gray"
        )
        info_label.grid(row=6, column=0, columnspan=2, pady=(10, 5))
    
    def _iniciar_lbtr(self):
        """Inicia el proceso LBTR con las credenciales ingresadas"""
        try:
            # Validar campos
            usuario = self.entry_user.get().strip()
            clave = self.entry_pass.get().strip()
            
            if not usuario:
                messagebox.showwarning("Advertencia", "Por favor ingrese el usuario")
                self.entry_user.focus()
                return
            
            if not clave:
                messagebox.showwarning("Advertencia", "Por favor ingrese la contraseña")
                self.entry_pass.focus()
                return
            
            # Confirmar inicio de sesión
            respuesta = messagebox.askyesno(
                "Confirmar",
                f"¿Iniciar proceso LBTR con el usuario '{usuario}'?\n\n"
                "Este proceso puede tomar varios minutos y abrirá un navegador web."
            )
            
            if not respuesta:
                return
            
            # Deshabilitar botones durante el proceso
            self.btn_iniciar_lbtr.configure(state="disabled", text="Iniciando...")
            self.btn_salir_login.configure(state="disabled")
            
            # Ejecutar LBTR en hilo separado
            self._ejecutar_lbtr_async(usuario, clave)
            
        except Exception as e:
            self.logger.error(f"Error iniciando LBTR: {e}")
            messagebox.showerror("Error", f"Error iniciando sesión LBTR: {e}")
            self._rehabilitar_botones()
    
    def _ejecutar_lbtr_async(self, usuario: str, clave: str):
        """Ejecuta LBTR de forma asíncrona"""
        def ejecutar():
            try:
                # Importar y crear instancia de LBTR
                from src.operations.lbtr_operations import LBTROperations
                lbtr = LBTROperations()
                
                # Cerrar ventana de login antes de iniciar
                self.after(0, self.on_closing)
                
                # Ejecutar proceso LBTR
                lbtr.exec_lbtr(usuario, clave)
                
            except Exception as e:
                self.logger.error(f"Error en hilo LBTR: {e}")
                # Mostrar error en el hilo principal
                self.after(0, lambda: self._mostrar_error_async(str(e)))
        
        # Crear y iniciar hilo
        hilo_lbtr = threading.Thread(target=ejecutar, daemon=True)
        hilo_lbtr.start()
        
        self.logger.info(f"Proceso LBTR iniciado en hilo separado para usuario: {usuario}")
    
    def _mostrar_error_async(self, error_msg: str):
        """Muestra error en el hilo principal"""
        messagebox.showerror("Error LBTR", f"Error ejecutando LBTR: {error_msg}")
        self._rehabilitar_botones()
    
    def _rehabilitar_botones(self):
        """Rehabilita los botones después de un error"""
        try:
            if hasattr(self, 'btn_iniciar_lbtr') and self.btn_iniciar_lbtr.winfo_exists():
                self.btn_iniciar_lbtr.configure(state="normal", text="Iniciar Sesión")
            if hasattr(self, 'btn_salir_login') and self.btn_salir_login.winfo_exists():
                self.btn_salir_login.configure(state="normal")
        except Exception:
            pass
    
    def _validar_credenciales(self, usuario: str, clave: str) -> bool:
        """
        Valida que las credenciales cumplan con requisitos básicos
        
        Args:
            usuario: Usuario a validar
            clave: Contraseña a validar
        
        Returns:
            True si las credenciales son válidas
        """
        # Validaciones básicas
        if len(usuario) < 3:
            messagebox.showwarning("Advertencia", "El usuario debe tener al menos 3 caracteres")
            return False
        
        if len(clave) < 4:
            messagebox.showwarning("Advertencia", "La contraseña debe tener al menos 4 caracteres")
            return False
        
        # Verificar caracteres especiales que podrían causar problemas
        caracteres_problematicos = ['<', '>', '"', "'", '&']
        for char in caracteres_problematicos:
            if char in usuario or char in clave:
                messagebox.showwarning(
                    "Advertencia", 
                    f"Las credenciales no pueden contener el carácter: {char}"
                )
                return False
        
        return True
    
    def _disable_parent_buttons(self):
        """Deshabilita botones en la ventana padre"""
        try:
            if hasattr(self.parent, 'operation_frames') and 'lbtr' in self.parent.operation_frames:
                frame = self.parent.operation_frames['lbtr']
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
            if hasattr(self.parent, 'operation_frames') and 'lbtr' in self.parent.operation_frames:
                frame = self.parent.operation_frames['lbtr']
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
            self.logger.info("Ventana de login LBTR cerrada")
            self.destroy()
        except Exception as e:
            self.logger.error(f"Error cerrando ventana de login LBTR: {e}")
            self.destroy()