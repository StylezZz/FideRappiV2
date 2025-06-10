"""
Gestor de archivos para FideRAPPI
Maneja operaciones relacionadas con archivos y directorios
"""

import os
import datetime
from pathlib import Path
from tkinter import messagebox
from typing import Optional

from src.core.base_logic import BaseLogic
from src.utils.logger import LoggerMixin


class FileManager(LoggerMixin):
    """Clase para manejar operaciones de archivos"""
    
    def __init__(self):
        """Inicializa el gestor de archivos"""
        self.base_logic = BaseLogic("FileManager")
    
    def abrir_historial(self, tipo_operacion: str, config_manager) -> bool:
        """
        Abre el historial de una operación específica
        
        Args:
            tipo_operacion: Tipo de operación (CCE, AHORROS, etc.)
            config_manager: Instancia del gestor de configuración
        
        Returns:
            True si se pudo abrir el historial
        """
        try:
            # Mapeo de operaciones a nombres de carpeta
            operaciones_map = {
                "CCE": "CCE",
                "CTA_CTES": "Cuentas corrientes", 
                "AHORROS": "Ahorros",
                "LBTR": "LBTR"
            }
            
            if tipo_operacion not in operaciones_map:
                messagebox.showwarning("Error", f"Tipo de operación no válido: {tipo_operacion}")
                return False
            
            nombre_carpeta = operaciones_map[tipo_operacion]
            
            # Obtener rutas de configuración
            ruta_origen, ruta_destino = config_manager.leer_json(tipo_operacion)
            if not ruta_origen:
                messagebox.showwarning("Error", "No se pudo obtener la configuración")
                return False
            
            directorio_base = os.path.dirname(ruta_origen)
            
            # Obtener fecha actual
            fecha_actual = datetime.datetime.now()
            fecha_info = self.base_logic.format_fecha_archivo(fecha_actual)
            
            # Construir rutas de historial
            ruta_hoy = Path(directorio_base) / "Procesados" / nombre_carpeta / fecha_info['year'] / f"{fecha_info['mes']}_{fecha_info['nombre_mes']}" / fecha_info['dia']
            ruta_mes = Path(directorio_base) / "Procesados" / nombre_carpeta / fecha_info['year'] / f"{fecha_info['mes']}_{fecha_info['nombre_mes']}"
            ruta_año = Path(directorio_base) / "Procesados" / nombre_carpeta / fecha_info['year']
            ruta_base = Path(directorio_base) / "Procesados" / nombre_carpeta
            
            # Intentar abrir en orden de prioridad
            rutas_a_probar = [ruta_hoy, ruta_mes, ruta_año, ruta_base]
            
            for ruta in rutas_a_probar:
                if ruta.exists() and ruta.is_dir():
                    self.logger.info(f"Abriendo historial en: {ruta}")
                    os.startfile(str(ruta))
                    return True
            
            # Si no existe ninguna ruta, mostrar mensaje
            messagebox.showwarning(
                "Historial no encontrado", 
                f"No existe historial para {nombre_carpeta} o no se ha usado el programa aún.\n\n"
                f"Se buscó en:\n{ruta_base}"
            )
            return False
            
        except Exception as e:
            self.logger.error(f"Error abriendo historial: {e}")
            messagebox.showerror("Error", f"Error abriendo historial: {e}")
            return False
    
    def crear_directorio_procesados(self, ruta_base: str, tipo_operacion: str, 
                                  fecha: Optional[datetime.datetime] = None) -> str:
        """
        Crea la estructura de directorios para archivos procesados
        
        Args:
            ruta_base: Directorio base
            tipo_operacion: Tipo de operación
            fecha: Fecha para la estructura (usa fecha actual si no se proporciona)
        
        Returns:
            Ruta del directorio creado
        """
        try:
            if fecha is None:
                fecha = datetime.datetime.now()
            
            fecha_info = self.base_logic.format_fecha_archivo(fecha)
            
            ruta_completa = Path(ruta_base) / "Procesados" / tipo_operacion / fecha_info['year'] / f"{fecha_info['mes']}_{fecha_info['nombre_mes']}" / fecha_info['dia']
            
            # Crear directorio si no existe
            ruta_completa.mkdir(parents=True, exist_ok=True)
            
            self.logger.info(f"Directorio creado/verificado: {ruta_completa}")
            return str(ruta_completa)
            
        except Exception as e:
            self.logger.error(f"Error creando directorio: {e}")
            raise
    
    def generar_nombre_archivo_procesado(self, tipo_operacion: str, memos: list, 
                                       prefijo: str = "MEMO", extension: str = ".xlsx") -> str:
        """
        Genera un nombre de archivo para archivos procesados
        
        Args:
            tipo_operacion: Tipo de operación
            memos: Lista de números de memo
            prefijo: Prefijo del archivo
            extension: Extensión del archivo
        
        Returns:
            Nombre del archivo generado
        """
        try:
            # Limpiar y ordenar memos
            numeros_memos = sorted(set(
                memo.split("-")[0] if isinstance(memo, str) and "-" in memo else str(memo)
                for memo in memos if memo
            ))
            
            if not numeros_memos:
                numeros_memos = ["SIN_MEMO"]
            
            nombre = f"{prefijo} {tipo_operacion} {'-'.join(numeros_memos)}{extension}"
            
            self.logger.info(f"Nombre de archivo generado: {nombre}")
            return nombre
            
        except Exception as e:
            self.logger.error(f"Error generando nombre de archivo: {e}")
            return f"{prefijo}_{tipo_operacion}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}{extension}"
    
    def verificar_archivo_existe(self, ruta: str) -> bool:
        """
        Verifica si un archivo existe
        
        Args:
            ruta: Ruta del archivo
        
        Returns:
            True si el archivo existe
        """
        try:
            return Path(ruta).exists()
        except Exception:
            return False
    
    def crear_backup_archivo(self, ruta_archivo: str) -> Optional[str]:
        """
        Crea un backup de un archivo
        
        Args:
            ruta_archivo: Ruta del archivo original
        
        Returns:
            Ruta del archivo de backup creado, None si no se pudo crear
        """
        try:
            archivo = Path(ruta_archivo)
            if not archivo.exists():
                return None
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"{archivo.stem}_backup_{timestamp}{archivo.suffix}"
            ruta_backup = archivo.parent / backup_name
            
            # Copiar archivo
            import shutil
            shutil.copy2(ruta_archivo, ruta_backup)
            
            self.logger.info(f"Backup creado: {ruta_backup}")
            return str(ruta_backup)
            
        except Exception as e:
            self.logger.error(f"Error creando backup: {e}")
            return None
    
    def limpiar_archivos_temporales(self, directorio: str, dias_antiguedad: int = 7):
        """
        Limpia archivos temporales antiguos
        
        Args:
            directorio: Directorio a limpiar
            dias_antiguedad: Días de antigüedad para considerar archivos como temporales
        """
        try:
            directorio_path = Path(directorio)
            if not directorio_path.exists():
                return
            
            limite_fecha = datetime.datetime.now() - datetime.timedelta(days=dias_antiguedad)
            archivos_eliminados = 0
            
            for archivo in directorio_path.rglob("*temp*"):
                if archivo.is_file():
                    fecha_modificacion = datetime.datetime.fromtimestamp(archivo.stat().st_mtime)
                    if fecha_modificacion < limite_fecha:
                        try:
                            archivo.unlink()
                            archivos_eliminados += 1
                        except Exception as e:
                            self.logger.warning(f"No se pudo eliminar {archivo}: {e}")
            
            if archivos_eliminados > 0:
                self.logger.info(f"Archivos temporales eliminados: {archivos_eliminados}")
                
        except Exception as e:
            self.logger.error(f"Error limpiando archivos temporales: {e}")
    
    def obtener_espacio_disco(self, ruta: str) -> dict:
        """
        Obtiene información del espacio en disco
        
        Args:
            ruta: Ruta para verificar el espacio
        
        Returns:
            Diccionario con información del espacio (total, usado, libre)
        """
        try:
            import shutil
            total, usado, libre = shutil.disk_usage(ruta)
            
            return {
                'total': total,
                'usado': usado, 
                'libre': libre,
                'porcentaje_usado': (usado / total) * 100 if total > 0 else 0
            }
        except Exception as e:
            self.logger.error(f"Error obteniendo espacio en disco: {e}")
            return {'total': 0, 'usado': 0, 'libre': 0, 'porcentaje_usado': 0}