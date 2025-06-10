"""
Lógica base compartida para todas las operaciones de FideRAPPI
"""

from datetime import datetime
from typing import Dict
from src.utils.logger import LoggerMixin

class BaseLogic(LoggerMixin):
    """Clase base con lógica compartida para todas las operaciones"""
    
    # Meses en español para usar en nombres de archivos
    MESES_ESPANOL = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo',
        4: 'Abril', 5: 'Mayo', 6: 'Junio',
        7: 'Julio', 8: 'Agosto', 9: 'Septiembre',
        10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    
    def __init__(self, tipo_operacion: str):
        """
        Inicializa la lógica base
        
        Args:
            tipo_operacion: Tipo de operación (CCE, AHORROS, etc.)
        """
        self.tipo_operacion = tipo_operacion
        self.ejecucion_en_progreso = True
        self.seleccion = None
        self.intervalo = 0.8  # Intervalo de espera para automatización
        self.detener_proceso = False
        self.deteccion_activa = False
        
        self.logger.info(f"Inicializando lógica para operación: {tipo_operacion}")
    
    def get_fecha_actual(self) -> datetime:
        """Obtiene la fecha actual"""
        return datetime.now()
    
    def get_mes_espanol(self, mes: int = None) -> str:
        """
        Obtiene el nombre del mes en español
        
        Args:
            mes: Número del mes (1-12). Si no se proporciona, usa el mes actual
        
        Returns:
            Nombre del mes en español
        """
        if mes is None:
            mes = self.get_fecha_actual().month
        return self.MESES_ESPANOL.get(mes, "Desconocido")
    
    def format_fecha_archivo(self, fecha: datetime = None) -> Dict[str, str]:
        """
        Formatea una fecha para usar en nombres de archivos
        
        Args:
            fecha: Fecha a formatear. Si no se proporciona, usa la fecha actual
        
        Returns:
            Diccionario con componentes de fecha formateados
        """
        if fecha is None:
            fecha = self.get_fecha_actual()
        
        return {
            'dia': str(fecha.day).zfill(2),
            'mes': str(fecha.month).zfill(2),
            'year': str(fecha.year),
            'nombre_mes': self.get_mes_espanol(fecha.month)
        }
    
    def generar_ruta_procesado(self, directorio_base: str, memos: list, 
                             prefijo: str = "MEMO", extension: str = ".xlsx") -> str:
        """
        Genera la ruta completa para guardar un archivo procesado
        
        Args:
            directorio_base: Directorio base donde guardar
            memos: Lista de números de memo
            prefijo: Prefijo para el nombre del archivo
            extension: Extensión del archivo
        
        Returns:
            Ruta completa del archivo
        """
        fecha_info = self.format_fecha_archivo()
        
        # Limpiar y ordenar memos
        numeros_memos = sorted(set(
            memo.split("-")[0] if isinstance(memo, str) and "-" in memo else str(memo)
            for memo in memos
        ))
        
        nombre_archivo = f"{prefijo} {self.tipo_operacion} {'-'.join(numeros_memos)}{extension}"
        
        ruta_completa = (
            f"{directorio_base}/Procesados/{self.tipo_operacion}/"
            f"{fecha_info['year']}/{fecha_info['mes']}_{fecha_info['nombre_mes']}/"
            f"{fecha_info['dia']}/{nombre_archivo}"
        )
        
        return ruta_completa
    
    def validar_entrada_numerica(self, valor: str, max_digitos: int = 4) -> bool:
        """
        Valida que una entrada sea numérica y tenga máximo cierta cantidad de dígitos
        
        Args:
            valor: Valor a validar
            max_digitos: Máximo número de dígitos permitidos
        
        Returns:
            True si es válido
        """
        if valor == "" or (valor.isdigit() and len(valor) <= max_digitos):
            return True
        return False
    
    def limpiar_texto_beneficiario(self, texto: str) -> str:
        """
        Limpia el texto de un beneficiario para cumplir con los estándares bancarios
        
        Args:
            texto: Texto a limpiar
        
        Returns:
            Texto limpiado
        """
        if not texto:
            return ""
        
        # Convertir a mayúsculas y limpiar caracteres especiales
        texto_limpio = texto.strip().upper()
        texto_limpio = texto_limpio.replace('\n', ' ')
        texto_limpio = texto_limpio.replace('Ñ', 'N')
        texto_limpio = texto_limpio.replace('&', 'Y')
        
        # Limpiar espacios múltiples
        while '  ' in texto_limpio:
            texto_limpio = texto_limpio.replace('  ', ' ')
        
        return texto_limpio.strip()
    
    def limpiar_numero_cuenta(self, numero: str) -> str:
        """
        Limpia un número de cuenta eliminando guiones y espacios
        
        Args:
            numero: Número de cuenta a limpiar
        
        Returns:
            Número limpio
        """
        if not numero:
            return ""
        
        return str(numero).replace('-', '').replace(' ', '').strip()
    
    def validar_longitud_cuenta(self, numero: str, longitud_esperada: int) -> bool:
        """
        Valida que un número de cuenta tenga la longitud esperada
        
        Args:
            numero: Número de cuenta
            longitud_esperada: Longitud esperada
        
        Returns:
            True si tiene la longitud correcta
        """
        numero_limpio = self.limpiar_numero_cuenta(numero)
        return len(numero_limpio) == longitud_esperada
    
    def formatear_monto(self, monto: float, decimales: int = 2) -> str:
        """
        Formatea un monto con el número especificado de decimales
        
        Args:
            monto: Monto a formatear
            decimales: Número de decimales
        
        Returns:
            Monto formateado como string
        """
        return f"{monto:.{decimales}f}"
    
    def detener_operacion(self):
        """Marca la operación para ser detenida"""
        self.detener_proceso = True
        self.ejecucion_en_progreso = False
        self.logger.info("Operación marcada para detener")
    
    def iniciar_operacion(self):
        """Inicia una operación"""
        self.detener_proceso = False
        self.ejecucion_en_progreso = True
        self.deteccion_activa = True
        self.logger.info("Operación iniciada")
    
    def finalizar_operacion(self):
        """Finaliza una operación"""
        self.ejecucion_en_progreso = False
        self.deteccion_activa = False
        self.logger.info("Operación finalizada")