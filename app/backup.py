import shutil
import time
import threading
import os
import sqlite3
from contextlib import closing
from pathlib import Path
from datetime import datetime

import settings

_last_db_mtime: float = 0.0

def perform_backup(is_manual: bool = False, max_files: int = 20) -> tuple[bool, str]:
    """
    Executa o backup usando SQLite Backup API nativa, protegendo contra locks.
    """
    global _last_db_mtime
    
    db_path = settings.get_db_path()
    if not db_path.exists():
        return False, "Banco de dados local não encontrado."
        
    current_mtime = os.path.getmtime(db_path)
    
    if not is_manual and current_mtime <= _last_db_mtime:
        return False, "Nenhum novo registro ou modificação detectada no banco."
        
    dest_dir = settings.get_network_backup_dir() if settings.get_rede_path() else settings.BACKUP_DIR
    try:
        dest_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        dest_dir = settings.BACKUP_DIR
        dest_dir.mkdir(parents=True, exist_ok=True)
        
    timestamp = datetime.now().strftime("%Y_%m_%d_%H%M%S")
    dest_file = dest_dir / f"numerador_backup_{timestamp}.sqlite"
    
    local_backup_dir = settings.BACKUP_DIR
    local_backup_dir.mkdir(parents=True, exist_ok=True)
    local_dest_file = local_backup_dir / f"numerador_backup_{timestamp}.sqlite"
    
    try:
        with closing(sqlite3.connect(db_path)) as source:
            with closing(sqlite3.connect(dest_file)) as dest:
                source.backup(dest)
            
            # Executa a cópia tbm para a máquina local se dest_file for rede
            if dest_file != local_dest_file:
                with closing(sqlite3.connect(local_dest_file)) as local_dest:
                    source.backup(local_dest)
                
        _last_db_mtime = current_mtime
        
        _cleanup_old_backups(dest_dir, max_files=max_files)
        if dest_file != local_dest_file:
            _cleanup_old_backups(local_backup_dir, max_files=max_files)
            
            # Mover antigos locais órfãos para rede
            for of_file in local_backup_dir.glob("numerador_backup_*.sqlite"):
                if of_file.name != local_dest_file.name:
                    try:
                        shutil.move(str(of_file), str(dest_dir))
                    except:
                        pass
        
        return True, f"Backup salvo com sucesso!\n\nArquivo: {dest_file.name}"
    except Exception as e:
        return False, f"Erro crítico ao copiar arquivo para o backup: {e}"

def _cleanup_old_backups(dest_dir: Path, max_files: int = 20):
    """
    Retenção inteligente de backups.
    """
    try:
        files = list(dest_dir.glob("numerador_backup_*.sqlite"))
        if len(files) <= max_files:
            return
            
        files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        protected = set(files[:max_files])
        
        now = datetime.now()
        found_yesterday = False
        found_last_week = False
        
        for f in files[max_files:]:
            mtime = datetime.fromtimestamp(f.stat().st_mtime)
            days_old = (now - mtime).days
            
            if 1 <= days_old <= 3 and not found_yesterday:
                protected.add(f)
                found_yesterday = True
            elif 5 <= days_old <= 15 and not found_last_week:
                protected.add(f)
                found_last_week = True

        for f in files:
            if f not in protected:
                try:
                    f.unlink()
                except:
                    pass
    except Exception as e:
        pass

def start_auto_backup(interval_minutes: int = 15, max_files: int = 20):
    """
    Inicia uma Thread em segundo plano pra gerar backups.
    """
    def _backup_loop():
        # Delay do ambiente do SO pra app respirar instantes iniciais
        time.sleep(10)
        while True:
            try:
                perform_backup(is_manual=False, max_files=max_files)
            except Exception:
                pass
            time.sleep(interval_minutes * 60)

    t = threading.Thread(target=_backup_loop, daemon=True)
    t.start()
