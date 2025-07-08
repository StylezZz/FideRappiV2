"""
Operaciones Extra
Funcionalidades adicionales como unión de PDFs
"""

import os
from pathlib import Path
from tkinter import messagebox
from typing import List

try:
    from PyPDF2 import PdfWriter, PdfReader
except ImportError:
    try:
        from pypdf import PdfWriter, PdfReader
    except ImportError:
        PdfWriter = None
        PdfReader = None

from src.utils.logger import LoggerMixin


class ExtraOperations(LoggerMixin):
    """Clase para manejar operaciones extra"""
    
    @staticmethod
    def combinar_pdfs(carpeta: str) -> bool:
        """
        Combina todos los archivos PDF de una carpeta en un solo archivo
        
        Args:
            carpeta: Ruta de la carpeta con los PDFs
        
        Returns:
            True si se combinaron correctamente
        """
        logger = ExtraOperations()
        
        try:
            if not PdfWriter or not PdfReader:
                messagebox.showerror(
                    "Error",
                    "La librería PyPDF2 o pypdf no está disponible.\n"
                    "Instale una de estas librerías para usar esta funcionalidad."
                )
                return False
            
            # Validar que la carpeta existe
            carpeta_path = Path(carpeta)
            if not carpeta_path.exists() or not carpeta_path.is_dir():
                messagebox.showerror("Error", "La carpeta seleccionada no existe o no es válida.")
                return False
            
            # Buscar archivos PDF
            archivos_pdf = []
            for extension in ['*.pdf', '*.PDF']:
                archivos_pdf.extend(list(carpeta_path.glob(extension)))
            
            # Verificar si hay archivos PDF
            if not archivos_pdf:
                messagebox.showinfo(
                    title="Sin archivos PDF",
                    message="No se encontraron archivos PDF en la carpeta seleccionada."
                )
                return False
            
            # Ordenar archivos por nombre
            archivos_pdf.sort(key=lambda x: x.name.lower())
            
            logger.logger.info(f"Encontrados {len(archivos_pdf)} archivos PDF para combinar")
            
            # Crear el merger
            merger = PdfWriter()
            archivos_procesados = 0
            archivos_con_error = []
            
            # Procesar cada archivo PDF
            for archivo_pdf in archivos_pdf:
                try:
                    logger.logger.info(f"Procesando: {archivo_pdf.name}")
                    
                    with open(archivo_pdf, 'rb') as pdf_file:
                        pdf_reader = PdfReader(pdf_file)
                        
                        # Verificar que el PDF no esté corrupto
                        if len(pdf_reader.pages) == 0:
                            logger.logger.warning(f"PDF vacío ignorado: {archivo_pdf.name}")
                            archivos_con_error.append(archivo_pdf.name)
                            continue
                        
                        # Añadir todas las páginas del PDF
                        for pagina_num in range(len(pdf_reader.pages)):
                            try:
                                pagina = pdf_reader.pages[pagina_num]
                                merger.add_page(pagina)
                            except Exception as e:
                                logger.logger.error(f"Error procesando página {pagina_num + 1} de {archivo_pdf.name}: {e}")
                        
                        archivos_procesados += 1
                        
                except Exception as e:
                    logger.logger.error(f"Error procesando {archivo_pdf.name}: {e}")
                    archivos_con_error.append(archivo_pdf.name)
                    continue
            
            if archivos_procesados == 0:
                messagebox.showerror(
                    "Error",
                    "No se pudieron procesar ningún archivo PDF. Verifique que los archivos no estén corruptos."
                )
                return False
            
            # Guardar el archivo combinado
            archivo_salida = carpeta_path / "PDFs_Combinados.pdf"
            contador = 1
            
            # Si ya existe, crear uno con numeración
            while archivo_salida.exists():
                archivo_salida = carpeta_path / f"PDFs_Combinados_{contador}.pdf"
                contador += 1
            
            try:
                with open(archivo_salida, 'wb') as archivo_output:
                    merger.write(archivo_output)
                
                # Cerrar el merger
                merger.close()
                
                # Mensaje de éxito
                mensaje_exito = f"PDF combinado guardado en:\n{archivo_salida}\n\n"
                mensaje_exito += f"Archivos procesados: {archivos_procesados}"
                
                if archivos_con_error:
                    mensaje_exito += f"\nArchivos con errores: {len(archivos_con_error)}"
                    if len(archivos_con_error) <= 5:
                        mensaje_exito += f"\n{', '.join(archivos_con_error)}"
                
                messagebox.showinfo(title="Proceso Completado", message=mensaje_exito)
                
                logger.logger.info(f"PDF combinado creado exitosamente: {archivo_salida}")
                return True
                
            except Exception as e:
                logger.logger.error(f"Error guardando archivo combinado: {e}")
                messagebox.showerror(
                    "Error",
                    f"Error al guardar el archivo combinado:\n{str(e)}"
                )
                return False
            
        except Exception as e:
            logger.logger.error(f"Error general en combinación de PDFs: {e}")
            messagebox.showerror(
                "Error",
                f"Error general al combinar PDFs:\n{str(e)}"
            )
            return False
    
    @staticmethod
    def verificar_dependencias_pdf() -> bool:
        """
        Verifica si las dependencias para PDF están disponibles
        
        Returns:
            True si están disponibles
        """
        return PdfWriter is not None and PdfReader is not None
    
    @staticmethod
    def combinar_pdfs_seleccionados(archivos_pdf: List[str]) -> bool:
        """Combina archivos PDF seleccionados individualmente"""
        logger = ExtraOperations()
        
        try:
            if not archivos_pdf:
                messagebox.showwarning("Advertencia", "No se seleccionaron archivos PDF.")
                return False
            if not PdfWriter or not PdfReader:
                messagebox.showerror("Error", "No se encontraron las librerías necesarias para PDF.")
                return False
            
            merger = PdfWriter()
            archivos_con_error = []
            archivos_procesados = 0
            
            for archivo in archivos_pdf:
                try:
                    with open(archivo, 'rb') as f:
                        reader = PdfReader(f)
                        for page in reader.pages:
                            merger.add_page(page)
                    archivos_procesados += 1
                except Exception as e:
                    archivos_con_error.append(os.path.basename(archivo))
                    logger.logger.error(f"Error procesando {archivo}: {e}")
            
            if archivos_procesados == 0:
                messagebox.showerror("Error", "No se pudieron procesar los archivos PDF.")
                return False
            
            carpeta_destino = os.path.dirname(archivos_pdf[0])
            salida = os.path.join(carpeta_destino, "PDFs_Combinados_Sueltos.pdf")
            contador = 1
            while os.path.exists(salida):
                salida = os.path.join(carpeta_destino, f"PDFs_Combinados_Sueltos_{contador}.pdf")
                contador += 1
            
            with open(salida, "wb") as f:
                merger.write(f)
            
            mensaje = f"PDF combinado guardado en:\n{salida}\n\nArchivos procesados: {archivos_procesados}"
            if archivos_con_error:
                mensaje += f"\nArchivos con errores: {len(archivos_con_error)}"
            
            messagebox.showinfo("Éxito", mensaje)
            return True
        
        except Exception as e:
            logger.logger.error(f"Error general combinando PDF seleccionados: {e}")
            messagebox.showerror("Error", f"Error combinando PDF seleccionados: {e}")
            return False
    
    @staticmethod
    def obtener_info_pdfs(carpeta: str) -> dict:
        """
        Obtiene información sobre los PDFs en una carpeta
        
        Args:
            carpeta: Ruta de la carpeta
        
        Returns:
            Diccionario con información de los PDFs
        """
        logger = ExtraOperations()
        
        try:
            carpeta_path = Path(carpeta)
            if not carpeta_path.exists():
                return {'error': 'Carpeta no existe'}
            
            archivos_pdf = []
            for extension in ['*.pdf', '*.PDF']:
                archivos_pdf.extend(list(carpeta_path.glob(extension)))
            
            info = {
                'total_archivos': len(archivos_pdf),
                'archivos': [],
                'tamaño_total': 0
            }
            
            for archivo in archivos_pdf:
                try:
                    tamaño = archivo.stat().st_size
                    info['archivos'].append({
                        'nombre': archivo.name,
                        'tamaño': tamaño,
                        'tamaño_mb': round(tamaño / (1024 * 1024), 2)
                    })
                    info['tamaño_total'] += tamaño
                except Exception as e:
                    logger.logger.error(f"Error obteniendo info de {archivo}: {e}")
            
            info['tamaño_total_mb'] = round(info['tamaño_total'] / (1024 * 1024), 2)
            return info
            
        except Exception as e:
            logger.logger.error(f"Error obteniendo información de PDFs: {e}")
            return {'error': str(e)}
    
    @staticmethod
    def limpiar_nombres_archivos(carpeta: str) -> bool:
        """
        Limpia nombres de archivos eliminando caracteres especiales
        
        Args:
            carpeta: Ruta de la carpeta
        
        Returns:
            True si se completó correctamente
        """
        logger = ExtraOperations()
        
        try:
            import re
            
            carpeta_path = Path(carpeta)
            if not carpeta_path.exists():
                messagebox.showerror("Error", "La carpeta no existe")
                return False
            
            archivos_renombrados = 0
            caracteres_prohibidos = r'[<>:"/\\|?*]'
            
            for archivo in carpeta_path.iterdir():
                if archivo.is_file():
                    nombre_original = archivo.name
                    nombre_limpio = re.sub(caracteres_prohibidos, '_', nombre_original)
                    nombre_limpio = re.sub(r'_{2,}', '_', nombre_limpio)  # Reemplazar múltiples _ con uno solo
                    
                    if nombre_limpio != nombre_original:
                        try:
                            nuevo_archivo = archivo.parent / nombre_limpio
                            archivo.rename(nuevo_archivo)
                            archivos_renombrados += 1
                            logger.logger.info(f"Renombrado: {nombre_original} -> {nombre_limpio}")
                        except Exception as e:
                            logger.logger.error(f"Error renombrando {nombre_original}: {e}")
            
            if archivos_renombrados > 0:
                messagebox.showinfo(
                    "Proceso Completado",
                    f"Se renombraron {archivos_renombrados} archivos."
                )
            else:
                messagebox.showinfo(
                    "Proceso Completado",
                    "No se encontraron archivos que necesiten ser renombrados."
                )
            
            return True
            
        except Exception as e:
            logger.logger.error(f"Error limpiando nombres de archivos: {e}")
            messagebox.showerror("Error", f"Error limpiando nombres: {str(e)}")
            return False