import tkinter as tk
import sys
import os

# Adiciona o diretorio atual no path para importacoes
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ui import NumeradorApp, LoginDialog
import db

if __name__ == "__main__":
    db.init_db()

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
        app = NumeradorApp(root, login.usuario_selecionado)
        root.mainloop()
    else:
        # Encerra se apenas clicou no (X) do login
        sys.exit()
