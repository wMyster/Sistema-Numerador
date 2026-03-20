import os
from pathlib import Path

# Diretório raiz do projeto (onde está o run.bat e a pasta app)
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Pastas de Dados Locais
LOCAL_DATA_DIR = PROJECT_ROOT / "data"
BACKUP_DIR = PROJECT_ROOT / "backup"

# Configuração de Rede (Unidade G:)
# O sistema tenta priorizar a rede para backup e base de dados,
# permitindo que múltiplos computadores vejam os mesmos números.
def get_rede_path() -> Path | None:
    """Verifica se a unidade G: (rede) está acessível."""
    p = Path(r"G:\NUMERADOR DADOS")
    if p.exists():
        return p
    return None

def get_network_data_dir() -> Path:
    """Retorna o diretório de dados na rede ou local se indisponível."""
    rede = get_rede_path()
    if rede:
        return rede / "DATA"
    return LOCAL_DATA_DIR

def get_network_backup_dir() -> Path:
    """Retorna o diretório de backup na rede ou local se indisponível."""
    rede = get_rede_path()
    if rede:
        return rede / "BACKUPS"
    return BACKUP_DIR

def get_db_path() -> Path:
    """Retorna o caminho do banco de dados ativo (Rede > Local)."""
    # Se a rede estiver disponível e o DB existir lá, usa ele
    db_rede = get_network_data_dir() / "numerador.sqlite"
    if db_rede.exists():
        return db_rede
    # Caso contrário, usa o local
    return LOCAL_DATA_DIR / "numerador.sqlite"

def get_users_file_path() -> Path:
    """Retorna o caminho do arquivo de usuários (Rede > Local)."""
    users_rede = get_network_data_dir() / "users.json"
    if users_rede.exists():
        return users_rede
    return LOCAL_DATA_DIR / "users.json"
