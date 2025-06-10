# src/operations/__init__.py
"""
Operaciones bancarias de FideRAPPI
Módulos para cada tipo de operación
"""

from . import cce_operations
from . import ahorros_operations
from . import cte_operations
from . import lbtr_operations
from . import cargo_operations
from . import extra_operations

__all__ = [
    'cce_operations',
    'ahorros_operations', 
    'cte_operations',
    'lbtr_operations',
    'cargo_operations',
    'extra_operations'
]