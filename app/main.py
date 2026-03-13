import tkinter as tk
import sys
import os

# Adiciona o diretorio atual no path para importacoes
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ui import NumeradorApp, LoginDialog
import db

if __name__ == "__main__":
    db.init_db()

    try:
        from backup import start_auto_backup
        start_auto_backup(interval_minutes=15)
    except: pass

    root = tk.Tk()
    root.withdraw() # Esconde a janela principal inicial para previnir fantasmas
    
    # Aplicando tema visual, se disponivel
    try:
        from tkinter import ttk
        style = ttk.Style()
        if 'vista' in style.theme_names():
            style.theme_use('vista')
        elif 'winnative' in style.theme_names():
            style.theme_use('winnative')
    except:
        pass

    # Exibe o dialogo de login travando o fluxo
    login = LoginDialog(root)
    root.wait_window(login)

    # Verifica se o usuario de fato selecionou um e fechou a janela ou se ele apenas deu (X)
    if login.usuario_selecionado:
        try:
            app = NumeradorApp(root, login.usuario_selecionado)
            root.deiconify() # restore a janela principal do tk.Tk (withdraw inicial)
            
            # Maximizar a janela na inicialização
            try:
                root.state('zoomed')
            except:
                pass # Tratamento para SOs onde zoom não é suportado
                
            root.mainloop()

            try:
                from network_lock import release_lock
                release_lock(login.usuario_selecionado)
            except: pass
        except Exception as e:
            import traceback
            from tkinter import messagebox
            root.deiconify()
            messagebox.showerror("Erro Fatal de Inicialização", f"Crash detectado: {e}\n\n{traceback.format_exc()}")
            sys.exit(1)
    else:
        # Encerra se apenas clicou no (X) do login
        sys.exit()
