"""
Módulo centralizado de logging para o Sistema Numerador.

Grava logs em arquivo (error.log) e no console (se disponível).
Níveis: DEBUG, INFO, WARNING, ERROR, CRITICAL.
"""
import logging
import os
from pathlib import Path

# Diretório raiz do projeto (onde está o run.bat)
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_LOG_FILE = _PROJECT_ROOT / "error.log"


def setup_logging(console: bool = False) -> logging.Logger:
    """
    Configura o logging global do sistema.

    Args:
        console: Se True, também exibe logs no console (para debug.bat).
    
    Returns:
        Logger raiz configurado.
    """
    root_logger = logging.getLogger()

    # Evita duplicar handlers se chamar duas vezes
    if root_logger.handlers:
        return root_logger

    root_logger.setLevel(logging.DEBUG)

    # Formato com data, nível, módulo e mensagem
    fmt = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%d/%m/%Y %H:%M:%S",
    )

    # Handler de arquivo (sempre ativo)
    try:
        file_handler = logging.FileHandler(
            _LOG_FILE, mode="a", encoding="utf-8"
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(fmt)
        root_logger.addHandler(file_handler)
    except Exception as e:
        print(f"Aviso: não foi possível criar arquivo de log: {e}")

    # Handler de console (ativado pelo debug.bat)
    if console:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)
        console_handler.setFormatter(fmt)
        root_logger.addHandler(console_handler)

    return root_logger


def get_logger(name: str) -> logging.Logger:
    """Retorna um logger nomeado para uso nos módulos."""
    return logging.getLogger(name)
