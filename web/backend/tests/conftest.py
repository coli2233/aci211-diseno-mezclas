"""Configuración de pytest para la API de ACI 211."""
import os
import sys

# Permitimos importar api.py desde web/backend
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def pytest_configure(config):
    # Evitamos que pytest recoja archivos generados por el cerebro
    config.addinivalue_line("norecursedirs", "fixtures")