"""
PDM-SW GUI Module
Interfaccia grafica principale suddivisa per tab.
"""

from .base_tab import BaseTab, warn, info, ask
from .tab_generatore import TabGeneratore
from .tab_codifica import TabCodifica
from .tab_gestione_codifica import TabGestioneCodifica
from .tab_manuale import TabManuale
from .tab_solidworks import TabSolidWorks
from .tab_gerarchia import TabGerarchia
from .tab_monitor import TabMonitor
from .tab_operativo import TabOperativo

__all__ = [
    "BaseTab",
    "warn",
    "info",
    "ask",
    "TabGeneratore",
    "TabCodifica",
    "TabGestioneCodifica",
    "TabManuale",
    "TabSolidWorks",
    "TabGerarchia",
    "TabMonitor",
    "TabOperativo",
]
