"""
Gerenciamento centralizado de configurações e caminhos do sistema.
Adaptado do Sistema Credencial.
"""
from __future__ import annotations

import configparser
import json
import shutil
from io import StringIO
from pathlib import Path
from typing import Optional


PROJECT_ROOT = Path(__file__).resolve().parents[1]

LOCAL_DATA_DIR = PROJECT_ROOT / "data"
CONFIG_FILE = PROJECT_ROOT / "config.ini"
SETTINGS_FILE = LOCAL_DATA_DIR / "settings.json"
LOCAL_USERS_FILE = LOCAL_DATA_DIR / "users.json"
LOCAL_DB_FILE = LOCAL_DATA_DIR / "numerador.sqlite"
BACKUP_DIR = PROJECT_ROOT / "backup"

# Raiz padrão do Numerador em rede.
NUMERADOR_ROOT = Path("G:/NUMERADOR DADOS")

def _read_config_value(section: str, key: str) -> Optional[str]:
    """Lê uma chave de um arquivo INI."""
    if not CONFIG_FILE.exists():
        return None

    try:
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE, encoding="utf-8")
        if config.has_option(section, key):
            return config.get(section, key)

        content = CONFIG_FILE.read_text(encoding="utf-8")
        if not content:
            return None
        if "[" not in content.splitlines()[0]:
            content = "[DEFAULT]\n" + content

        config = configparser.ConfigParser()
        config.read_file(StringIO(content))
        return config.get("DEFAULT", key, fallback=None)
    except Exception:
        return None


def _configured_network_data_dir() -> Optional[Path]:
    path_str = (_read_config_value("DEFAULT", "REDE_PATH") or "").strip()
    if not path_str:
        return None
    return Path(path_str)


def get_numerador_root() -> Path:
    """Retorna a raiz lógica da rede."""
    configured = _configured_network_data_dir()
    if configured:
        if configured.name.upper() == "DATA":
            return configured.parent
        return configured.parent
    return NUMERADOR_ROOT


def get_network_data_dir() -> Path:
    configured = _configured_network_data_dir()
    if configured:
        return configured
    return get_numerador_root() / "DATA"


def get_network_backup_dir() -> Path:
    """Retorna o caminho de backup na rede."""
    return get_numerador_root() / "BACKUPS"


def get_rede_path() -> Optional[Path]:
    """Retorna o diretório de dados da rede quando ele estiver acessível."""
    data_dir = get_network_data_dir()
    try:
        if data_dir.exists():
            return data_dir

        parent = data_dir.parent
        if parent.exists():
            data_dir.mkdir(parents=True, exist_ok=True)
            return data_dir
    except Exception:
        pass
    return None


def get_data_dir() -> Path:
    """Retorna o diretório base para dados (rede se disponível, senão local)."""
    rede = get_rede_path()
    if rede:
        try:
            rede.mkdir(parents=True, exist_ok=True)
            return rede
        except Exception:
            pass

    LOCAL_DATA_DIR.mkdir(parents=True, exist_ok=True)
    return LOCAL_DATA_DIR


def get_db_path() -> Path:
    return get_data_dir() / "numerador.sqlite"


def get_users_file_path() -> Path:
    """Retorna o caminho local do users.json unificado. Nós optamos por gravar no G: para que todos vejam as mesmas contas, mas com fallback local."""
    rede = get_rede_path()
    if rede:
        try:
            rede.mkdir(parents=True, exist_ok=True)
            return rede / "users.json"
        except:
            pass
    LOCAL_DATA_DIR.mkdir(parents=True, exist_ok=True)
    return LOCAL_DATA_DIR / "users.json"
