import tkinter as tk
import sys
import os

# Adiciona o diretorio atual no path para importacoes
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Detecta modo debug (passado pelo debug.bat via --debug)
_DEBUG_MODE = "--debug" in sys.argv

# Inicializa logging ANTES de tudo
from app_logger import setup_logging, get_logger
setup_logging(console=_DEBUG_MODE)
logger = get_logger(__name__)

from ui import NumeradorApp, LoginDialog
import db

if __name__ == "__main__":
    logger.info("=" * 60)
    logger.info("Iniciando Sistema Único de Numeradores")
    logger.info("Modo debug: %s", "ATIVADO" if _DEBUG_MODE else "DESATIVADO")

    # Prevenção de múltiplas instâncias via Mutex do Windows
    try:
        import ctypes
        import time

        _mutex_handle = None
        mutex_name = "Local\\SistemaNumerador_Mutex_Unico"
        already_running = False

        for _ in range(5):
            handle = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
            if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
                ctypes.windll.kernel32.CloseHandle(handle)
                already_running = True
                time.sleep(0.5)
            else:
                _mutex_handle = handle
                already_running = False
                break

        if already_running:
            logger.warning("Outra instância já está em execução. Encerrando.")
            try:
                from tkinter import messagebox
                root = tk.Tk()
                root.withdraw()
                messagebox.showwarning(
                    "Atenção",
                    "O Sistema Numerador já está aberto neste computador!\n\n"
                    "Verifique se o ícone dele já não está ali embaixo na barra de tarefas."
                )
                root.destroy()
            except Exception:
                pass
            sys.exit(0)

        logger.info("Mutex adquirido — instância única confirmada")
    except Exception as e:
        logger.error("Erro ao verificar mutex: %s", e, exc_info=True)

    try:
        db.init_db()
        logger.info("Banco de dados inicializado com sucesso")
    except Exception as e:
        logger.critical("Falha ao inicializar banco de dados: %s", e, exc_info=True)
        sys.exit(1)

    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal inicial para previnir fantasmas
    
    # Aplicando tema visual, se disponivel
    try:
        from tkinter import ttk
        style = ttk.Style()
        if 'vista' in style.theme_names():
            style.theme_use('vista')
        elif 'winnative' in style.theme_names():
            style.theme_use('winnative')
    except Exception as e:
        logger.debug("Tema visual não disponível: %s", e)

    # Exibe o dialogo de login travando o fluxo
    login = LoginDialog(root)
    root.wait_window(login)

    # Verifica se o usuario de fato selecionou um e fechou a janela ou se ele apenas deu (X)
    if login.usuario_selecionado:
        logger.info("Usuário logado: %s", login.usuario_selecionado)
        try:
            app = NumeradorApp(root, login.usuario_selecionado)
            root.mainloop()
        except Exception as e:
            logger.critical("Erro fatal na aplicação: %s", e, exc_info=True)
            raise
        finally:
            logger.info("Aplicação encerrada")
    else:
        logger.info("Login cancelado pelo usuário. Encerrando.")
        sys.exit()

