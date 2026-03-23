import os
import sys
import threading
import sqlite3
import subprocess
import tkinter as tk
from tkinter import messagebox, simpledialog
from datetime import datetime
from typing import Optional, List, Tuple

import customtkinter as ctk
from PIL import Image, ImageTk

import db
import export_docx
import settings
from app_logger import get_logger

logger = get_logger(__name__)

# ======================================================================
# Cores e Estilos Modernos (Palette: Slate & Azure)
# ======================================================================
_C = {
    "bg":           "#f1f5f9",   # Slate 100
    "card":         "#ffffff",   # White
    "border":       "#e2e8f0",   # Slate 200
    "accent":       "#2563eb",   # Blue 600
    "sidebar":      "#1e293b",   # Slate 800
    "sidebar_btn":  "#334155",   # Slate 700 (hover)
    "text":         "#1e293b",   # Slate 800
    "text_muted":   "#64748b",   # Slate 500
    "success":      "#10b981",   # Emerald 500
    "danger":       "#ef4444",   # Red 500
}

# Configurações globais do CTK
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

TIPOS_NUMERADOR = [
    ("OFICIO", "Ofício"),
    ("MEMORANDO", "Memorando"),
    ("CIRCULAR_INTERNA", "Circular Interna"),
    ("NOTIFICACAO", "Notificação"),
    ("PORTARIA", "Portaria"),
    ("AUTORIZACAO_VEICULO", "Autorização de Veículo Oficial"),
    ("CERTIDAO", "Certidão")
]

class LoginDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Acesso ao Sistema")
        self.geometry("450x650")
        self.resizable(False, False)
        self.after(200, lambda: self.iconbitmap("assets/logo.png") if os.path.exists("assets/logo.png") else None)
        
        # Centralizar
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.winfo_screenheight() // 2) - (650 // 2)
        self.geometry(f'+{x}+{y}')

        self.usuario_selecionado = None
        self.configure(fg_color=_C["bg"])
        
        # -- Layout --
        # Header Blue Gradient (simulado com frame)
        self.header = ctk.CTkFrame(self, fg_color=_C["sidebar"], height=200, corner_radius=0)
        self.header.pack(fill="x", side="top")
        
        try:
            logo_img = ctk.CTkImage(Image.open("assets/logo.png"), size=(100, 100))
            self.logo_lbl = ctk.CTkLabel(self.header, image=logo_img, text="")
            self.logo_lbl.pack(pady=(40, 10))
        except:
            pass
            
        self.title_lbl = ctk.CTkLabel(
            self.header, text="SISTEMA NUMERADOR", 
            font=ctk.CTkFont(size=24, weight="bold"), text_color="white"
        )
        self.title_lbl.pack()
        
        self.subtitle_lbl = ctk.CTkLabel(
            self.header, text="Gestão de Documentos Oficiais", 
            font=ctk.CTkFont(size=13), text_color="#94a3b8"
        )
        self.subtitle_lbl.pack(pady=(0, 20))

        # Card de Login
        self.card = ctk.CTkFrame(self, fg_color="white", corner_radius=20, border_width=1, border_color=_C["border"])
        self.card.place(relx=0.5, rely=0.65, anchor="center", relwidth=0.85, relheight=0.6)
        
        ctk.CTkLabel(
            self.card, text="Identifique-se para continuar", 
            font=ctk.CTkFont(size=14, weight="bold"), text_color=_C["text"]
        ).pack(pady=(25, 15))

        # Campo de Busca (Filtro)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.recarregar_lista())
        self.search_entry = ctk.CTkEntry(
            self.card, textvariable=self.search_var, placeholder_text="Buscar seu nome...",
            height=35, corner_radius=10, border_width=1, fg_color="#f8fafc"
        )
        self.search_entry.pack(fill="x", padx=25, pady=(0, 10))

        # Lista de Usuarios
        self.scroll_frame = ctk.CTkScrollableFrame(
            self.card, fg_color="transparent", 
            scrollbar_button_color=_C["border"], scrollbar_button_hover_color="#cbd5e1"
        )
        self.scroll_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        self.user_buttons = []
        self.usuarios_all = db.get_all_usuarios()
        self.recarregar_lista()

        # Botão principal
        self.btn_entrar = ctk.CTkButton(
            self.card, text="ENTRAR NO SISTEMA", 
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=_C["accent"], hover_color="#1d4ed8",
            height=45, corner_radius=12, command=self.entrar
        )
        self.btn_entrar.pack(fill="x", padx=25, pady=(15, 20))
        
        # Rodapé com ações secundárias
        self.footer = ctk.CTkFrame(self.card, fg_color="transparent")
        self.footer.pack(fill="x", padx=25, pady=(0, 15))
        
        ctk.CTkButton(
            self.footer, text="+ Novo Usuário", width=100, height=28,
            fg_color="transparent", text_color=_C["text_muted"], hover_color=_C["bg"],
            command=self.novo_usuario
        ).pack(side="left")

        ctk.CTkButton(
            self.footer, text="Remover", width=80, height=28,
            fg_color="transparent", text_color="#fca5a5", hover_color="#fff1f2",
            command=self.excluir_usuario
        ).pack(side="right")

        self.focus_force()
        self.grab_set()
        self.bind('<Return>', lambda e: self.entrar())

    def recarregar_lista(self):
        # Limpar
        for btn in self.user_buttons:
            btn.destroy()
        self.user_buttons = []
        
        filtro = self.search_var.get().upper()
        usuarios = [u for u in self.usuarios_all if filtro in u.upper()]
        
        for u in usuarios:
            is_selected = (u == self.usuario_selecionado)
            bg = _C["accent"] if is_selected else "transparent"
            fg = "white" if is_selected else _C["text"]
            
            btn = ctk.CTkButton(
                self.scroll_frame, text=f"👤  {u}", 
                font=ctk.CTkFont(size=13, weight="bold" if is_selected else "normal"),
                fg_color=bg, text_color=fg,
                hover_color=_C["bg"] if not is_selected else None, 
                corner_radius=10, anchor="w", height=38,
                command=lambda name=u: self.selecionar_user(name)
            )
            btn.pack(fill="x", pady=3, padx=5)
            self.user_buttons.append(btn)
        
        if not self.usuario_selecionado and usuarios:
            self.selecionar_user(usuarios[0])

    def selecionar_user(self, name):
        self.usuario_selecionado = name
        # Atualizar visual dos botões sem destruir tudo (otimização)
        for btn in self.user_buttons:
            btn_text = btn.cget("text").replace("👤  ", "")
            if btn_text == name:
                btn.configure(fg_color=_C["accent"], text_color="white", font=ctk.CTkFont(size=13, weight="bold"))
            else:
                btn.configure(fg_color="transparent", text_color=_C["text"], font=ctk.CTkFont(size=13, weight="normal"))

    def entrar(self):
        if self.usuario_selecionado:
            self.destroy()
        else:
            messagebox.showwarning("Aviso", "Por favor, selecione um usuário.", parent=self)

    def novo_usuario(self):
        novo = simpledialog.askstring("Novo Usuário", "Digite o nome do novo usuário:", parent=self)
        if novo and novo.strip():
            nome = novo.strip().upper()
            try:
                db.add_usuario(nome)
                self.usuarios_all = db.get_all_usuarios()
                self.usuario_selecionado = nome
                self.recarregar_lista()
            except Exception as e:
                messagebox.showerror("Erro", "Falha ao criar usuário.")

    def excluir_usuario(self):
        if not self.usuario_selecionado: return
        if self.usuario_selecionado in ["DIRETORIA", "VIA DCT"]:
            messagebox.showwarning("Aviso", "Usuários do sistema não podem ser removidos.")
            return
            
        if messagebox.askyesno("Confirmar", f"Excluir definitivamente o usuário {self.usuario_selecionado}?", parent=self):
            db.delete_usuario(self.usuario_selecionado)
            self.usuarios_all = db.get_all_usuarios()
            self.usuario_selecionado = self.usuarios_all[0] if self.usuarios_all else None
            self.recarregar_lista()


def show_login():
    """Exibe a tela de login e retorna o usuário selecionado."""
    root = ctk.CTk()
    root.withdraw()
    login = LoginDialog(root)
    root.wait_window(login)
    user = login.usuario_selecionado
    root.destroy()
    return user

class TabNumerador(ctk.CTkFrame):
    def __init__(self, parent, app, tipo_db, titulo_aba):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.tipo_db = tipo_db
        self.titulo_aba = titulo_aba
        self.selected_id = None
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        self.build_ui()
        self.action_novo()
        self.load_data()

    def build_ui(self):
        # Card Superior: Formulário
        self.form_card = ctk.CTkFrame(self, fg_color="white", corner_radius=15, border_width=1, border_color=_C["border"])
        self.form_card.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        self.form_card.grid_columnconfigure((0,1,2,3,4,5), weight=1)

        # Labels e Entradas Modernas
        ctk.CTkLabel(self.form_card, text=f"Numerando: {self.titulo_aba}", font=ctk.CTkFont(size=16, weight="bold"), text_color=_C["accent"]).grid(row=0, column=0, columnspan=2, padx=20, pady=15, sticky="w")
        
        # Linha 1
        ctk.CTkLabel(self.form_card, text="Nº Registro", font=ctk.CTkFont(size=11, weight="bold")).grid(row=1, column=0, padx=20, sticky="w")
        self.var_numero = tk.StringVar()
        self.entry_numero = ctk.CTkEntry(self.form_card, textvariable=self.var_numero, state="readonly", font=ctk.CTkFont(weight="bold"), fg_color="#f8fafc")
        self.entry_numero.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")

        ctk.CTkLabel(self.form_card, text="Data", font=ctk.CTkFont(size=11, weight="bold")).grid(row=1, column=1, padx=10, sticky="w")
        self.var_data = tk.StringVar()
        ctk.CTkEntry(self.form_card, textvariable=self.var_data).grid(row=2, column=1, padx=10, pady=(0, 10), sticky="ew")

        if self.tipo_db == "CERTIDAO":
            ctk.CTkLabel(self.form_card, text="Placa", font=ctk.CTkFont(size=11, weight="bold")).grid(row=1, column=2, padx=10, sticky="w")
            self.var_placa = tk.StringVar()
            ctk.CTkEntry(self.form_card, textvariable=self.var_placa).grid(row=2, column=2, padx=10, pady=(0, 10), sticky="ew")

        # Linha 2
        ctk.CTkLabel(self.form_card, text="Assunto / Detalhes", font=ctk.CTkFont(size=11, weight="bold")).grid(row=3, column=0, padx=20, sticky="w")
        self.var_assunto = tk.StringVar()
        ctk.CTkEntry(self.form_card, textvariable=self.var_assunto).grid(row=4, column=0, columnspan=6, padx=20, pady=(0, 10), sticky="ew")

        # Linha 3
        ctk.CTkLabel(self.form_card, text="Destino / Interessado", font=ctk.CTkFont(size=11, weight="bold")).grid(row=5, column=0, padx=20, sticky="w")
        self.var_destino = tk.StringVar()
        ctk.CTkEntry(self.form_card, textvariable=self.var_destino).grid(row=6, column=0, columnspan=3, padx=20, pady=(0, 15), sticky="ew")

        ctk.CTkLabel(self.form_card, text="Observações", font=ctk.CTkFont(size=11, weight="bold")).grid(row=5, column=3, padx=10, sticky="w")
        self.var_obs = tk.StringVar()
        ctk.CTkEntry(self.form_card, textvariable=self.var_obs).grid(row=6, column=3, columnspan=3, padx=10, pady=(0, 15), sticky="ew")

        # Botões de Ação
        self.actions_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.actions_frame.grid(row=0, column=0, sticky="e", padx=40, pady=20)
        
        self.btn_save = ctk.CTkButton(self.actions_frame, text="Salvar Registro", fg_color=_C["success"], hover_color="#059669", command=self.action_salvar)
        self.btn_save.pack(side="right", padx=5)
        
        self.btn_docx = ctk.CTkButton(self.actions_frame, text="Abrir DOCX", fg_color="transparent", text_color=_C["accent"], border_width=1, border_color=_C["accent"], command=self.action_abrir_numerador)
        self.btn_docx.pack(side="right", padx=5)

        self.btn_new = ctk.CTkButton(self.actions_frame, text="Limpar", fg_color=_C["sidebar_btn"], command=self.action_novo)
        self.btn_new.pack(side="right", padx=5)

        # Tabela Inferior
        self.table_card = ctk.CTkFrame(self, fg_color="white", corner_radius=15, border_width=1, border_color=_C["border"])
        self.table_card.grid(row=1, column=0, sticky="nsew", padx=20, pady=(10, 20))
        self.table_card.grid_columnconfigure(0, weight=1)
        self.table_card.grid_rowconfigure(1, weight=1)

        # Busca
        self.search_bar = ctk.CTkFrame(self.table_card, fg_color="transparent")
        self.search_bar.grid(row=0, column=0, sticky="ew", padx=20, pady=15)
        
        self.var_busca = tk.StringVar()
        self.entry_busca = ctk.CTkEntry(self.search_bar, textvariable=self.var_busca, placeholder_text="Pesquisar registros...", width=300)
        self.entry_busca.pack(side="left")
        self.entry_busca.bind("<Return>", lambda e: self.load_data())
        
        ctk.CTkButton(self.search_bar, text="Buscar", width=80, command=self.load_data).pack(side="left", padx=10)
        ctk.CTkButton(self.search_bar, text="Limpar", width=60, fg_color="transparent", text_color=_C["text_muted"], command=self.limpar_busca).pack(side="left")

        # Treeview (via tkinter.ttk)
        from tkinter import ttk
        cols = ("id", "numero", "data", "assunto", "destino", "usuario")
        if self.tipo_db == "CERTIDAO": cols = ("id", "numero", "placa", "data", "assunto", "destino", "usuario")
        
        self.tree = ttk.Treeview(self.table_card, columns=cols, show="headings")
        self.tree.heading("numero", text="Nº")
        self.tree.heading("data", text="DATA")
        self.tree.heading("assunto", text="ASSUNTO")
        self.tree.heading("destino", text="DESTINO")
        self.tree.heading("usuario", text="POR")
        
        self.tree.column("id", width=0, stretch=False)
        self.tree.column("numero", width=60, anchor="center")
        self.tree.column("data", width=100, anchor="center")
        self.tree.column("assunto", width=350)
        self.tree.column("destino", width=200)
        
        self.tree.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
    def action_abrir_numerador(self):
        try:
            arquivo = export_docx.get_docx_path(self.tipo_db)
            if os.path.exists(arquivo):
                os.startfile(arquivo)
            else:
                messagebox.showerror("Erro", "Arquivo DOCX não existe ou ainda não foi gerado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir DOCX: {e}")

    def action_abrir_pasta(self):
        try:
            pasta = db.NUMERADORES_DIR
            if os.path.exists(pasta):
                os.startfile(pasta)
            else:
                os.makedirs(pasta, exist_ok=True)
                os.startfile(pasta)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir pasta: {e}")

    def action_novo(self):
        self.selected_id = None
        prox = db.get_proximo_numero(self.tipo_db)
        self.var_numero.set(f"{prox:03d}")
        self.var_data.set(datetime.now().strftime("%d/%m/%Y"))
        self.var_assunto.set("")
        self.var_destino.set("")
        self.var_obs.set("")
        if self.tipo_db == "CERTIDAO": self.var_placa.set("")

    def action_salvar(self):
        if not self.var_assunto.get().strip():
            messagebox.showwarning("Aviso", "Preencha o assunto.")
            return
            
        try:
            if self.selected_id:
                db.update_registro(self.selected_id, self.var_placa.get() if self.tipo_db=="CERTIDAO" else "", self.var_data.get(), self.var_assunto.get(), self.var_destino.get(), self.var_obs.get(), self.app.usuario_atual)
            else:
                db.insert_registro(self.tipo_db, None, self.var_placa.get() if self.tipo_db=="CERTIDAO" else "", self.var_data.get(), self.var_assunto.get(), self.var_destino.get(), self.var_obs.get(), self.app.usuario_atual)
            
            messagebox.showinfo("Sucesso", "Registro salvo com sucesso!")
            self.action_novo()
            self.load_data()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar: {e}")

    def load_data(self, silent=False):
        registros = db.get_all_registros(self.tipo_db, self.var_busca.get())
        for r in self.tree.get_children(): self.tree.delete(r)
        for r in registros:
            # r: (id, numero, placa, data, assunto, destino, obs, usuario)
            if self.tipo_db == "CERTIDAO":
                self.tree.insert("", "end", values=(r[0], f"{r[1]:03d}", r[2], r[3], r[4], r[5], r[7]))
            else:
                self.tree.insert("", "end", values=(r[0], f"{r[1]:03d}", r[3], r[4], r[5], r[7]))

    def limpar_busca(self):
        self.var_busca.set("")
        self.load_data()

    def on_tree_select(self, e):
        sel = self.tree.selection()
        if not sel: return
        
        item_vals = self.tree.item(sel[0])['values']
        self.selected_id = item_vals[0]
        
        # Buscar registro completo do banco para garantir que temos as Obs (que podem não estar na Tree)
        r = db.get_registro_by_id(self.tipo_db, self.selected_id)
        if r:
            # r: (id, numero, placa, data, assunto, destino, obs, usuario)
            self.var_numero.set(f"{r[1]:03d}")
            if self.tipo_db == "CERTIDAO": self.var_placa.set(r[2] if r[2] else "")
            self.var_data.set(r[3])
            self.var_assunto.set(r[4])
            self.var_destino.set(r[5])
            self.var_obs.set(r[6] if r[6] else "")

class TabRelatorios(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Header
        self.header = ctk.CTkFrame(self, fg_color="white", corner_radius=15, border_width=1, border_color=_C["border"])
        self.header.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        ctk.CTkLabel(
            self.header, text="Estatísticas de Volumes (2026)", 
            font=ctk.CTkFont(size=18, weight="bold"), text_color=_C["accent"]
        ).pack(side="left", padx=20, pady=20)
        
        self.btn_refresh = ctk.CTkButton(
            self.header, text="Atualizar Dados", width=120, height=32, 
            fg_color=_C["sidebar"], command=self.load_data
        )
        self.btn_refresh.pack(side="right", padx=20)

        # Main Content (Card with Table)
        self.main_card = ctk.CTkFrame(self, fg_color="white", corner_radius=15, border_width=1, border_color=_C["border"])
        self.main_card.grid(row=1, column=0, sticky="nsew", padx=20, pady=(10, 20))
        self.main_card.grid_columnconfigure(0, weight=1)
        self.main_card.grid_rowconfigure(0, weight=1)

        from tkinter import ttk
        style = ttk.Style()
        style.configure("Report.Treeview", background="white", foreground=_C["text"], rowheight=40, fieldbackground="white", font=("Segoe UI", 11))
        style.configure("Report.Treeview.Heading", background="#f8fafc", foreground=_C["text_muted"], font=("Segoe UI", 10, "bold"))
        
        self.tree = ttk.Treeview(self.main_card, columns=("tipo", "quantidade"), show="headings", style="Report.Treeview")
        self.tree.heading("tipo", text="TIPO DO DOCUMENTO")
        self.tree.heading("quantidade", text="QUANTIDADE")
        self.tree.column("tipo", width=400, anchor="w")
        self.tree.column("quantidade", width=150, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        
        self.load_data()

    def load_data(self, silent=False):
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        historico = db.get_estatisticas()
        total_geral = 0
        importes_nominais = dict(TIPOS_NUMERADOR)
        
        for i, (tipo_db, qtd) in enumerate(historico):
            total_geral += qtd
            nome = importes_nominais.get(tipo_db, tipo_db)
            self.tree.insert("", tk.END, values=(nome, qtd))
            
        self.tree.insert("", tk.END, values=("-"*30, "-"*10))
        self.tree.insert("", tk.END, values=("TOTAIS (GERAL)", total_geral))

class TabAuditoria(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        self.header = ctk.CTkFrame(self, fg_color="white", corner_radius=15, border_width=1, border_color=_C["border"])
        self.header.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        ctk.CTkLabel(
            self.header, text="Log de Atividades do Sistema", 
            font=ctk.CTkFont(size=18, weight="bold"), text_color=_C["accent"]
        ).pack(side="left", padx=20, pady=20)
        
        ctk.CTkButton(
            self.header, text="Limpar Log", width=100, height=32, 
            fg_color="transparent", text_color=_C["danger"], hover_color="#fee2e2",
            command=lambda: messagebox.showinfo("Aviso", "Função restrita ao administrador.")
        ).pack(side="right", padx=20)

        self.main_card = ctk.CTkFrame(self, fg_color="white", corner_radius=15, border_width=1, border_color=_C["border"])
        self.main_card.grid(row=1, column=0, sticky="nsew", padx=20, pady=(10, 20))
        self.main_card.grid_columnconfigure(0, weight=1)
        self.main_card.grid_rowconfigure(0, weight=1)

        from tkinter import ttk
        self.tree = ttk.Treeview(self.main_card, columns=("data_hora", "usuario", "acao", "detalhes"), show="headings")
        self.tree.heading("data_hora", text="DATA/HORA")
        self.tree.heading("usuario", text="USUÁRIO")
        self.tree.heading("acao", text="AÇÃO")
        self.tree.heading("detalhes", text="DETALHES")
        
        for col in ("data_hora", "usuario", "acao"):
            self.tree.column(col, width=150, anchor="center")
        self.tree.column("detalhes", width=400, anchor="w")
        
        self.tree.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        
        scroll = ctk.CTkScrollbar(self.main_card, command=self.tree.yview)
        scroll.pack(side="right", fill="y", pady=20, padx=(0, 10))
        self.tree.configure(yscrollcommand=scroll.set)
        
        self.load_data()

    def load_data(self, silent=False):
        for row in self.tree.get_children():
            self.tree.delete(row)
        try:
            historico = db.get_todas_auditorias(500)
            for r in historico:
                self.tree.insert("", tk.END, values=(r[0], r[1], r[2], r[3]))
        except:
            pass

class NumeradorApp(ctk.CTk):
    def __init__(self, usuario_atual):
        super().__init__()
        self.usuario_atual = usuario_atual
        self.title("Sistema Numerador 2026")
        self.geometry("1200x800")
        self.minsize(1100, 700)
        self.configure(fg_color=_C["bg"])
        
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Layout principal: Sidebar e Conteúdo
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Carregar Ícones (Sincronizado com os nomes gerados no plano)
        self.load_icons()
        
        # Sidebar
        self.sidebar = ctk.CTkFrame(self, fg_color=_C["sidebar"], width=240, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        
        # Logo na Sidebar
        self.logo_label = ctk.CTkLabel(
            self.sidebar, text=" NUMERADOR", image=self.icon_logo, 
            compound="left", font=ctk.CTkFont(size=20, weight="bold"), text_color="white"
        )
        self.logo_label.pack(pady=(30, 40), padx=20)
        
        # Navegação
        self.nav_buttons = {}
        nav_items = [
            ("DASHBOARD", "dashboard", self.icon_dash),
            ("OFICIO", "Ofícios", self.icon_new),
            ("MEMORANDO", "Memorandos", self.icon_new),
            ("COMUNICADO", "Comunicados", self.icon_new),
            ("CERTIDAO", "Certidões (Placas)", self.icon_new),
            ("REGISTRO", "Atividades de Log", self.icon_search),
        ]
        
        for key, label, icon in nav_items:
            btn = ctk.CTkButton(
                self.sidebar, text=f"  {label}", image=icon, 
                compound="left", anchor="w", height=45,
                fg_color="transparent", text_color="#cbd5e1",
                hover_color=_C["sidebar_btn"], corner_radius=8,
                command=lambda k=key: self.show_page(k)
            )
            btn.pack(fill="x", padx=15, pady=4)
            self.nav_buttons[key] = btn

        # Espaço Flexível
        ctk.CTkLabel(self.sidebar, text="").pack(expand=True)
        
        # Rodapé Sidebar (Usuário)
        ui_info = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        ui_info.pack(fill="x", padx=15, pady=20)
        
        ctk.CTkLabel(
            ui_info, text=f"● {self.usuario_atual}", 
            font=ctk.CTkFont(size=12, weight="bold"), text_color=_C["success"]
        ).pack(anchor="w")
        
        self.btn_logout = ctk.CTkButton(
            ui_info, text="Trocar Usuário", font=ctk.CTkFont(size=11),
            fg_color="transparent", text_color="#94a3b8", hover_color=_C["sidebar_btn"],
            height=25, command=self.action_trocar_usuario
        )
        self.btn_logout.pack(anchor="w", pady=(5, 0))

        # Container de Conteúdo
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.content_frame.grid(row=0, column=1, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)
        
        self.pages = {}
        self.show_page("DASHBOARD")
        
        self.auto_sync_loop()
        self.auto_backup_loop()

    def load_icons(self):
        try:
            self.icon_logo   = ctk.CTkImage(Image.open("assets/logo.png"), size=(32, 32))
            self.icon_dash   = ctk.CTkImage(Image.open("assets/dashboard.png"), size=(20, 20))
            self.icon_new    = ctk.CTkImage(Image.open("assets/new_record.png"), size=(20, 20))
            self.icon_search = ctk.CTkImage(Image.open("assets/search.png"), size=(20, 20))
            self.icon_trash  = ctk.CTkImage(Image.open("assets/trash.png"), size=(20, 20))
            self.icon_settings = ctk.CTkImage(Image.open("assets/settings.png"), size=(20, 20))
        except:
             # Fallback icons if assets not found
             self.icon_logo = self.icon_dash = self.icon_new = self.icon_search = self.icon_trash = self.icon_settings = None

    def show_page(self, key):
        # Limpar página anterior se necessário ou trazer para frente
        for p in self.pages.values():
            p.grid_forget()
            
        # Highlight Button
        for k, btn in self.nav_buttons.items():
            if k == key:
                btn.configure(fg_color=_C["accent"], text_color="white")
            else:
                btn.configure(fg_color="transparent", text_color="#cbd5e1")

        if key == "DASHBOARD":
            if key not in self.pages: self.pages[key] = TabRelatorios(self.content_frame, self)
        elif key == "REGISTRO":
            if key not in self.pages: self.pages[key] = TabAuditoria(self.content_frame, self)
        else:
            # Abas de Numeradores (Oficio, Memorando, etc)
            if key not in self.pages:
                titulo = dict(TIPOS_NUMERADOR).get(key, key)
                self.pages[key] = TabNumerador(self.content_frame, self, key, titulo)
        
        self.pages[key].grid(row=0, column=0, sticky="nsew")
        if hasattr(self.pages[key], 'load_data'):
            self.pages[key].load_data(silent=True)

    def on_closing(self):
        if messagebox.askyesno("Sair", "Deseja encerrar o sistema?"):
            self.destroy()
            sys.exit()

    def sincro_manual(self):
        try:
            db.sincronizar_bancos()
            for p in self.pages.values():
                if hasattr(p, 'load_data'): p.load_data(silent=True)
            messagebox.showinfo("Sucesso", "Banco sincronizado!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha na sincronização: {e}")

    def auto_sync_loop(self):
        try:
            db.sincronizar_bancos()
            for p in self.pages.values():
                if hasattr(p, 'load_data'): p.load_data(silent=True)
        except: pass
        self.after(60000, self.auto_sync_loop)

    def auto_backup_loop(self):
        try: db.fazer_backup()
        except: pass
        self.after(900000, self.auto_backup_loop)

    def action_trocar_usuario(self):
        if messagebox.askyesno("Logout", "Deseja trocar de usuário?"):
            self.destroy()
            try:
                bat_path = os.path.join(os.getcwd(), "run.bat")
                if os.path.exists(bat_path):
                    subprocess.Popen([bat_path], shell=True)
                else:
                    subprocess.Popen([sys.executable, "app/main.py"])
            except: pass
            sys.exit()

