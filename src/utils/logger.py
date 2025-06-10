"""
Sistema de logging para FideRAPPI
Proporciona logging configurable para la aplicación
"""

import logging
import sys
from pathlib import Path
from datetime import datetime
from typing import Optional

def setup_logger(name: str = "FideRAPPI", level: int = logging.INFO) -> logging.Logger:
    """
    Configura el sistema de logging de la aplicación
    
    Args:
        name: Nombre del logger
        level: Nivel de logging
    
    Returns:
        Logger configurado
    """
    # Obtener directorio base
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent.parent.parent
    
    # Crear directorio de logs
    logs_dir = base_dir / "logs"
    logs_dir.mkdir(exist_ok=True)
    
    # Crear logger
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # Evitar duplicar handlers
    if logger.handlers:
        return logger
    
    # Formato de logging
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Handler para archivo
    log_file = logs_dir / f"fiderapp_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Handler para consola (solo errores críticos)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.ERROR)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger

def get_logger(name: str = "FideRAPPI") -> logging.Logger:
    """
    Obtiene el logger de la aplicación
    
    Args:
        name: Nombre del logger
    
    Returns:
        Logger existente o nuevo
    """
    return logging.getLogger(name)

class LoggerMixin:
    """Mixin para añadir logging a cualquier clase"""
    
    @property
    def logger(self) -> logging.Logger:
        """Obtiene el logger para la clase"""
        if not hasattr(self, '_logger'):
            self._logger = get_logger(f"FideRAPPI.{self.__class__.__name__}")
        return self._logger