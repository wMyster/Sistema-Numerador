import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
import db
import export_docx

TIPOS_NUMERADOR = [
    ("OFICIO", "Ofício"),
    ("MEMORANDO", "Memorando"),
    ("CIRCULAR_INTERNA", "Circular Interna"),
    ("NOTIFICACAO", "Notificação"),
    ("PORTARIA", "Portaria"),
    ("AUTORIZACAO_VEICULO", "Autorização de Veículo Oficial"),
    ("CERTIDAO", "Certidão")
]

class LoginDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Acesso ao Sistema Único")
        self.geometry("450x420")
        self.resizable(False, False)
        
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

        self.usuario_selecionado = None
        
        container = ttk.Frame(self, padding="25 25 25 25")
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="SELECIONE O SEU USUÁRIO:", font=("Segoe UI", 12, "bold")).pack(pady=(0, 10))

        self.usuarios = db.get_all_usuarios()
        
        # Frame da Lista
        list_frame = ttk.Frame(container)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox_usuarios = tk.Listbox(list_frame, font=("Segoe UI", 12, "bold"), 
                                           yscrollcommand=scroll.set,
                                           selectbackground="#0078D7",
                                           selectforeground="white",
                                           activestyle="none",
                                           height=6)
        
        for u in self.usuarios:
            self.listbox_usuarios.insert(tk.END, u)
            
        self.listbox_usuarios.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.config(command=self.listbox_usuarios.yview)
        
        if self.usuarios:
            self.listbox_usuarios.selection_set(0)

        style = ttk.Style()
        style.configure("Login.TButton", font=("Segoe UI", 11, "bold"), padding=6)

        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill=tk.X)
        
        btn_entrar = ttk.Button(btn_frame, text="✅ ENTRAR", style="Login.TButton", command=self.entrar)
        btn_entrar.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        
        btn_novo = ttk.Button(btn_frame, text="➕ NOVO", style="Login.TButton", command=self.novo_usuario)
        btn_novo.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 2))

        btn_excluir = ttk.Button(btn_frame, text="🗑️ EXCLUIR", style="Login.TButton", command=self.excluir_usuario)
        btn_excluir.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))

        self.focus_force()
        self.grab_set()
        self.bind('<Return>', lambda e: self.entrar())
        self.listbox_usuarios.bind('<Double-1>', lambda e: self.entrar())

    def entrar(self):
        try:
            selecionado = self.listbox_usuarios.get(self.listbox_usuarios.curselection())
            if selecionado:
                self.usuario_selecionado = selecionado
                self.destroy()
        except:
            messagebox.showwarning("Aviso", "Selecione um usuário para entrar.", parent=self)

    def novo_usuario(self):
        novo = simpledialog.askstring("Novo Usuário", "Digite o nome do novo usuário:", parent=self)
        if novo and novo.strip():
            novo = novo.strip().upper()
            try:
                db.add_usuario(novo)
                self.usuarios = db.get_all_usuarios()
                self.listbox_usuarios.delete(0, tk.END)
                for u in self.usuarios:
                    self.listbox_usuarios.insert(tk.END, u)
                try: 
                    idx = self.usuarios.index(novo)
                    self.listbox_usuarios.selection_set(idx)
                    self.listbox_usuarios.see(idx)
                except: pass
                messagebox.showinfo("Sucesso", f"Usuário '{novo}' cadastrado com sucesso!", parent=self)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível criar. Usuário já existente ou erro interno:\n{e}", parent=self)

    def excluir_usuario(self):
        try:
            selecionado = self.listbox_usuarios.get(self.listbox_usuarios.curselection())
            if selecionado in ['VIA DCT', 'DIRETORIA']:
                messagebox.showwarning("Aviso", "Usuários nativos não podem ser excluídos.", parent=self)
                return
            if messagebox.askyesno("Excluir", f"Deseja realmente EXCLUIR permanentemente o usuário '{selecionado}'?", parent=self):
                db.delete_usuario(selecionado)
                self.usuarios = db.get_all_usuarios()
                self.listbox_usuarios.delete(0, tk.END)
                for u in self.usuarios:
                    self.listbox_usuarios.insert(tk.END, u)
                if self.usuarios:
                    self.listbox_usuarios.selection_set(0)
                messagebox.showinfo("Sucesso", "Usuário excluído com sucesso!", parent=self)
        except:
            messagebox.showwarning("Aviso", "Selecione um usuário para excluir.", parent=self)


class TabNumerador(ttk.Frame):
    def __init__(self, parent, app, tipo_db, titulo_aba):
        super().__init__(parent)
        self.app = app
        self.tipo_db = tipo_db
        self.titulo_aba = titulo_aba
        self.selected_id = None
        
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        
        self.build_ui()
        self.action_novo()
        self.load_data()

    def build_ui(self):
        top_container = ttk.Frame(self, padding="10 10 10 0")
        top_container.grid(row=0, column=0, sticky="nsew")
        top_container.columnconfigure(0, weight=1)

        form_frame = ttk.LabelFrame(top_container, text=f"  Cadastro: {self.titulo_aba}  ", padding="15 15 15 15")
        form_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        
        for p in range(5):
             form_frame.columnconfigure(p, weight=1)

        # Linha 1
        ttk.Label(form_frame, text="Nº Registro:", foreground="#444").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        self.var_numero = tk.StringVar()
        self.entry_numero = ttk.Entry(form_frame, textvariable=self.var_numero, state='readonly', takefocus=False, width=12, font=("Segoe UI", 11, "bold"))
        self.entry_numero.grid(row=0, column=1, sticky=tk.W, pady=(0, 10))
        self.entry_numero.bind("<FocusIn>", lambda e: self.entry_numero.selection_clear())
        
        # Placa (Condicional)
        self.var_placa = tk.StringVar()
        if self.tipo_db == "CERTIDAO":
            ttk.Label(form_frame, text="Placa:").grid(row=0, column=2, sticky=tk.E, padx=(10, 5), pady=(0, 10))
            self.entry_placa = ttk.Entry(form_frame, textvariable=self.var_placa, width=15, font=("Segoe UI", 10))
            self.entry_placa.grid(row=0, column=3, sticky=tk.W, pady=(0, 10))

        ttk.Label(form_frame, text="Data:").grid(row=0, column=4, sticky=tk.E, padx=(10, 5), pady=(0, 10))
        self.var_data = tk.StringVar()
        self.entry_data = ttk.Entry(form_frame, textvariable=self.var_data, width=15, font=("Segoe UI", 10))
        self.entry_data.grid(row=0, column=5, sticky=tk.W, pady=(0, 10))

        ttk.Label(form_frame, text="Responsável:").grid(row=0, column=6, sticky=tk.E, padx=(10, 5), pady=(0, 10))
        self.var_usuario = tk.StringVar(value=self.app.usuario_atual)
        self.entry_usuario = ttk.Entry(form_frame, textvariable=self.var_usuario, state='readonly', width=20, font=("Segoe UI", 10))
        self.entry_usuario.grid(row=0, column=7, sticky=tk.W, pady=(0, 10))
        
        # Linha 2
        ttk.Label(form_frame, text="Assunto:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.var_assunto = tk.StringVar()
        self.entry_assunto = ttk.Entry(form_frame, textvariable=self.var_assunto, font=("Segoe UI", 10))
        self.entry_assunto.grid(row=1, column=1, columnspan=7, sticky="we", pady=5)
        
        # Linha 3
        ttk.Label(form_frame, text="Destino:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.var_destino = tk.StringVar()
        self.entry_destino = ttk.Entry(form_frame, textvariable=self.var_destino, font=("Segoe UI", 10))
        self.entry_destino.grid(row=2, column=1, columnspan=7, sticky="we", pady=5)
        
        # Linha 4
        ttk.Label(form_frame, text="Observações:").grid(row=3, column=0, sticky=tk.W, pady=(5, 0))
        self.var_obs = tk.StringVar()
        self.entry_obs = ttk.Entry(form_frame, textvariable=self.var_obs, font=("Segoe UI", 10))
        self.entry_obs.grid(row=3, column=1, columnspan=7, sticky="we", pady=(5, 0))
        
        # Botoes
        btn_container = ttk.Frame(top_container)
        btn_container.grid(row=1, column=0, sticky="we", pady=(5, 5))
        
        grupo_registro = ttk.LabelFrame(btn_container, text="Registro", padding="5")
        grupo_registro.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        ttk.Button(grupo_registro, text="📄 Limpar / Novo", command=self.action_novo).pack(fill=tk.X, pady=2)
        ttk.Button(grupo_registro, text="💾 Salvar Alterações", command=self.action_salvar).pack(fill=tk.X, pady=2)
        
        grupo_docs = ttk.LabelFrame(btn_container, text="Documentos DOCX", padding="5")
        grupo_docs.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Button(grupo_docs, text="👁️ Ver Documento Individual", command=self.action_abrir_selecionado).pack(fill=tk.X, pady=2)
        ttk.Button(grupo_docs, text=f"📖 Abrir Livro: {self.titulo_aba}", command=self.action_abrir_numerador).pack(fill=tk.X, pady=2)

        # --- Gerenciamento (Lado Direito) ---
        grupo_sys = ttk.LabelFrame(btn_container, text="Gerenciamento de Registros e Contas", padding="5")
        grupo_sys.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        ttk.Button(grupo_sys, text="🔄 Sincronizar", command=self.app.sincro_manual).pack(side=tk.LEFT, fill=tk.X, padx=2, pady=2)
        ttk.Button(grupo_sys, text="👤 Trocar Usuário", command=self.app.action_trocar_usuario).pack(side=tk.LEFT, fill=tk.X, padx=2, pady=2)
        ttk.Button(grupo_sys, text="🗑️ Excluir Selecionado", command=self.action_excluir).pack(side=tk.LEFT, fill=tk.X, padx=2, pady=2)
        
        # --- TABELA Inferior ---
        bottom_container = ttk.Frame(self, padding="10 0 10 10")
        bottom_container.grid(row=1, column=0, sticky="nsew")
        bottom_container.rowconfigure(1, weight=1)
        bottom_container.columnconfigure(0, weight=1)

        search_frame = ttk.Frame(bottom_container)
        search_frame.grid(row=0, column=0, sticky="we", pady=(10, 5))
        
        ttk.Label(search_frame, text="🔍 Pesquisar:").pack(side=tk.LEFT, padx=(0, 5))
        self.var_busca = tk.StringVar()
        self.entry_busca = ttk.Entry(search_frame, textvariable=self.var_busca, width=50, font=("Segoe UI", 10))
        self.entry_busca.pack(side=tk.LEFT)
        self.entry_busca.bind('<Return>', lambda e: self.load_data())
        ttk.Button(search_frame, text="Buscar", command=self.load_data, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(search_frame, text="✕", command=self.limpar_busca, width=3).pack(side=tk.LEFT)

        banco_mode = "REDE (G:)" if db.get_active_db_path() == db.REDE_DB_PATH else "LOCAL (Fallback)"
        ttk.Label(search_frame, text=f"Status: {banco_mode}", foreground="gray").pack(side=tk.RIGHT, padx=10)

        tree_frame = ttk.Frame(bottom_container)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        
        if self.tipo_db == "CERTIDAO":
            columns = ("id", "numero", "placa", "data", "assunto", "destino", "obs", "usuario")
        else:
            columns = ("id", "numero", "data", "assunto", "destino", "obs", "usuario")
            
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        self.tree.heading("id", text="ID")
        self.tree.heading("numero", text="Nº")
        
        if self.tipo_db == "CERTIDAO":
            self.tree.heading("placa", text="Placa")
            self.tree.column("placa", width=80, anchor=tk.CENTER, stretch=tk.NO)
            
        self.tree.heading("data", text="Data")
        self.tree.heading("assunto", text="Assunto")
        self.tree.heading("destino", text="Destino")
        self.tree.heading("obs", text="Observações")
        self.tree.heading("usuario", text="Registrado Por")
        
        self.tree.column("id", width=0, stretch=tk.NO)
        self.tree.column("numero", width=60, anchor=tk.CENTER, stretch=tk.NO)
        self.tree.column("data", width=90, anchor=tk.CENTER, stretch=tk.NO)
        self.tree.column("assunto", width=250)
        self.tree.column("destino", width=200)
        self.tree.column("obs", width=150)
        self.tree.column("usuario", width=120, anchor=tk.W, stretch=tk.NO)
        
        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=scroll_y.set, xscroll=scroll_x.set)
        
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.tag_configure('oddrow', background='#F9F9F9')
        self.tree.tag_configure('evenrow', background='#FFFFFF')

    def action_novo(self):
        self.selected_id = None
        prox_num = db.get_proximo_numero(self.tipo_db)
        self.var_numero.set(f"{prox_num:03d}")
        if self.tipo_db == "CERTIDAO":
            self.var_placa.set("")
        self.var_data.set(datetime.now().strftime("%d/%m/%Y"))
        self.var_assunto.set("")
        self.var_destino.set("")
        self.var_obs.set("")
        self.var_usuario.set(self.app.usuario_atual)
        for item in self.tree.selection():
            self.tree.selection_remove(item)
            
    def action_salvar(self):
        num_str = self.var_numero.get()
        if not num_str: return
            
        numero = int(num_str)
        placa = self.var_placa.get().strip() if self.tipo_db == "CERTIDAO" else ""
        data = self.var_data.get()
        assunto = self.var_assunto.get().strip()
        destino = self.var_destino.get().strip()
        obs = self.var_obs.get().strip()
        usuario = self.var_usuario.get()
        
        if not assunto:
            messagebox.showwarning("Aviso", "Recomendamos preencher o Assunto.")
            
        # POP-UP de CONFIRMAÇÃO
        acao = "Atualizar" if self.selected_id else "Adicionar"
        if not messagebox.askyesno("Confirmar", f"Deseja {acao} o Número {numero:03d}?"):
            return
            
        try:
            if self.selected_id is None:
                db.insert_registro(self.tipo_db, numero, placa, data, assunto, destino, obs, usuario)
            else:
                db.update_registro(self.selected_id, placa, data, assunto, destino, obs, usuario)
                
            db.log_auditoria(self.app.usuario_atual, acao.upper(), f"{self.tipo_db} Nº {numero:03d}")
                
            self.action_novo()
            self.load_data()
            try:
                export_docx.exportar_para_docx(self.tipo_db)
            except:
                pass
            messagebox.showinfo("Sucesso", "Registrado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao registrar:\n{e}")
            
    def action_excluir(self):
        if not self.selected_id:
            messagebox.showwarning("Aviso", "Selecione um registro na grade abaixo para excluir.")
            return
            
        if messagebox.askyesno("Excluir Documento", f"ALERTA! Tem certeza que deseja apagar DEFINITIVAMENTE o registro {self.var_numero.get()}?"):
            try:
                numero_apagado = self.var_numero.get()
                db.delete_registro(self.selected_id)
                db.log_auditoria(self.app.usuario_atual, "EXCLUIR", f"{self.tipo_db} Nº {numero_apagado}")
                self.action_novo()
                self.load_data()
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")

    def action_abrir_selecionado(self):
        if not self.selected_id:
            messagebox.showwarning("Aviso", "Selecione um registro para visualizar.")
            return
            
        try:
            numero = int(self.var_numero.get())
            placa = self.var_placa.get() if self.tipo_db == "CERTIDAO" else ""
            data = self.var_data.get()
            assunto = self.var_assunto.get()
            destino = self.var_destino.get()
            obs = self.var_obs.get()
            usuario = self.var_usuario.get()
            
            caminho = export_docx.exportar_unico_docx(self.tipo_db, numero, placa, data, assunto, destino, obs, usuario)
            os.startfile(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possivel abrir DOCX:\n{e}")
                
    def action_abrir_numerador(self):
        try:
            caminho = export_docx.exportar_para_docx(self.tipo_db)
            os.startfile(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar e abrir o DOCX:\n{e}")
            
    def load_data(self, silent=False):
        busca = self.var_busca.get().strip()
        registros = db.get_all_registros(self.tipo_db, busca)
        
        if silent:
            current_ids = [self.tree.item(item)['values'][0] for item in self.tree.get_children()]
            new_ids = [r[0] for r in registros]
            if current_ids == new_ids:
                return

        for row in self.tree.get_children():
            self.tree.delete(row)
            
        for i, r in enumerate(registros):
            p_id, numero, placa, data, assunto, destino, obs, usuario = r
            numero_fmt = f"{numero:03d}"
            tags = ('evenrow',) if i % 2 == 0 else ('oddrow',)
            
            if self.tipo_db == "CERTIDAO":
                self.tree.insert("", tk.END, values=(p_id, numero_fmt, placa, data, assunto, destino, obs, usuario), tags=tags)
            else:
                self.tree.insert("", tk.END, values=(p_id, numero_fmt, data, assunto, destino, obs, usuario), tags=tags)
            
    def limpar_busca(self):
        self.var_busca.set("")
        self.load_data()
        
    def on_tree_select(self, event):
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])
            valores = item['values']
            if valores:
                self.selected_id = valores[0]
                self.var_numero.set(f"{int(valores[1]):03d}")
                
                if self.tipo_db == "CERTIDAO":
                    self.var_placa.set(valores[2] if valores[2] != 'None' else '')
                    self.var_data.set(valores[3])
                    self.var_assunto.set(valores[4])
                    self.var_destino.set(valores[5])
                    self.var_obs.set(valores[6] if valores[6] != 'None' else '')
                    if len(valores) > 7:
                        self.var_usuario.set(valores[7] if str(valores[7]) != 'None' else '')
                else:
                    self.var_data.set(valores[2])
                    self.var_assunto.set(valores[3])
                    self.var_destino.set(valores[4])
                    self.var_obs.set(valores[5] if valores[5] != 'None' else '')
                    if len(valores) > 6:
                        self.var_usuario.set(valores[6] if str(valores[6]) != 'None' else '')

class TabRelatorios(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        container = ttk.Frame(self, padding="20 20 20 20")
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)
        
        ttk.Label(container, text="Relatório Consolidado de Volumes (2026)", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, pady=(0, 20))
        
        tree_frame = ttk.Frame(container)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        
        self.tree = ttk.Treeview(tree_frame, columns=("tipo", "quantidade"), show="headings", height=10)
        self.tree.heading("tipo", text="Tipo do Documento")
        self.tree.heading("quantidade", text="Total de Registros Base")
        
        self.tree.column("tipo", width=400, anchor=tk.W)
        self.tree.column("quantidade", width=150, anchor=tk.CENTER)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=2, column=0, pady=20, sticky="ew")
        
        ttk.Button(btn_frame, text="Atualizar Dados", command=self.load_data).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="👤 Trocar Usuário / Sair", command=self.app.action_trocar_usuario).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="🔄 Sincronizar Agora", command=self.app.sincro_manual).pack(side=tk.RIGHT, padx=5)

        self.load_data()

    def load_data(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        historico = db.get_estatisticas()
        total_geral = 0
        importes_nominais = dict(TIPOS_NUMERADOR)
        
        for i, (tipo_db, qtd) in enumerate(historico):
            total_geral += qtd
            nome = importes_nominais.get(tipo_db, tipo_db)
            tags = ('evenrow',) if i % 2 == 0 else ('oddrow',)
            self.tree.insert("", tk.END, values=(nome, qtd), tags=tags)
            
        self.tree.insert("", tk.END, values=("----------", "----------"))
        self.tree.insert("", tk.END, values=("TOTAIS (GERAL)", total_geral), tags=('evenrow',))
        
        self.tree.tag_configure('oddrow', background='#F9F9F9')
        self.tree.tag_configure('evenrow', background='#FFFFFF')

class TabAuditoria(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        container = ttk.Frame(self, padding="20 20 20 20")
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)
        
        ttk.Label(container, text="LOG DE ATIVIDADES DE USUÁRIO", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, pady=(0, 20))
        
        tree_frame = ttk.Frame(container)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        
        self.tree = ttk.Treeview(tree_frame, columns=("data_hora", "usuario", "acao", "detalhes"), show="headings", height=15)
        self.tree.heading("data_hora", text="Data e Hora")
        self.tree.heading("usuario", text="Usuário Agente")
        self.tree.heading("acao", text="Ação Executada")
        self.tree.heading("detalhes", text="Alvo/Detalhes")
        
        self.tree.column("data_hora", width=150, anchor=tk.CENTER)
        self.tree.column("usuario", width=120, anchor=tk.CENTER)
        self.tree.column("acao", width=120, anchor=tk.CENTER)
        self.tree.column("detalhes", width=400, anchor=tk.W)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scroll_y.set)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=2, column=0, pady=20, sticky="ew")
        
        ttk.Button(btn_frame, text="Recarregar Histórico", command=self.load_data).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="👤 Trocar Usuário / Sair", command=self.app.action_trocar_usuario).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="🔄 Sincronizar Agora", command=self.app.sincro_manual).pack(side=tk.RIGHT, padx=5)

        self.load_data()

    def load_data(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        try:
            historico = db.get_todas_auditorias(1000)
            for i, r in enumerate(historico):
                tags = ('evenrow',) if i % 2 == 0 else ('oddrow',)
                self.tree.insert("", tk.END, values=(r[0], r[1], r[2], r[3]), tags=tags)
        except: pass
            
        self.tree.tag_configure('oddrow', background='#F9F9F9')
        self.tree.tag_configure('evenrow', background='#FFFFFF')

class NumeradorApp:
    def __init__(self, root, usuario_atual):
        self.root = root
        self.usuario_atual = usuario_atual
        self.root.title(f"Sistema Único de Numeradores 2026  |  Usuário: {self.usuario_atual}")
        self.root.geometry("1050x750")
        self.root.minsize(1050, 750)
        
        # INTERCEPTAR O FECHAMENTO DO X DA JANELA (POPUP)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        style = ttk.Style()
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10), padding=5)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28)
        
        # O Notebook Multiplexador de Abas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        self.abas = {}
        # Abas de Edição
        for (tipo_db, titulo_aba) in TIPOS_NUMERADOR:
            tab = TabNumerador(self.notebook, self, tipo_db, titulo_aba)
            self.notebook.add(tab, text=f" {titulo_aba} ")
            self.abas[tipo_db] = tab
            
        # Aba de Relatorios Isolada
        self.tab_relatorios = TabRelatorios(self.notebook, self)
        self.notebook.add(self.tab_relatorios, text=" RELATÓRIOS ")
        
        # Aba de Auditoria / Log de ações
        self.tab_auditoria = TabAuditoria(self.notebook, self)
        self.notebook.add(self.tab_auditoria, text=" REGISTRO ")
            
        self.auto_sync_loop()
        self.auto_backup_loop()
        
        # Elimina as seleções das células azuis passivas ao desenhar Interface inicial
        self.root.focus_set()
        
        # Monitora cliques nas abas, tirando instantâneamente o Foco Azul que os Elementos Tkinter tentam forçar
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self.root.focus_set())
        
        # Exibe a janela principal APENAS após concluir todo o recheio do Grid 
        # Resolvendo o problema de janela fantasma do Login.
        self.root.deiconify()
        
    def on_closing(self):
        if messagebox.askyesno("Encerrar Sistema", "Deseja realmente sair e fechar o aplicativo?"):
            self.root.destroy()
            try: sys.exit()
            except: pass
            
    def sincro_manual(self):
        try:
            db.sincronizar_bancos()
            for tab in self.abas.values():
                tab.load_data(silent=True)
            self.tab_relatorios.load_data()
            try: self.tab_auditoria.load_data() 
            except: pass
            messagebox.showinfo("Sucesso", "Banco sincronizado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na sincronização: {e}")

    def auto_sync_loop(self):
        try:
            db.sincronizar_bancos()
            for tab in self.abas.values():
                tab.load_data(silent=True)
            try: self.tab_auditoria.load_data()
            except: pass
        except Exception:
            pass
        self.root.after(3000, self.auto_sync_loop)
        
    def auto_backup_loop(self):
        try:
            db.fazer_backup()
        except:
            pass
        # 900000 ms = 15 minutos
        self.root.after(900000, self.auto_backup_loop)
        
    def action_trocar_usuario(self):
        if messagebox.askyesno("Confirmar", "Deseja realizar o Logo-ff do sistema?"):
            self.root.destroy()
            try:
                # Procura pelo batch de execução original dois niveis acima
                base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                bat_path = os.path.join(base_dir, "run.bat")
                if os.path.exists(bat_path):
                    import subprocess
                    subprocess.Popen([bat_path], shell=True)
                else:
                    # Fallback para script local
                    import subprocess
                    subprocess.Popen([sys.executable, sys.argv[0]])
            except:
                pass
            sys.exit()

