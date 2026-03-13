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
        self.title("Acesso ao Sistema")
        self.geometry("380x340")
        self.resizable(False, False)
        
        self.configure(bg="#ffffff")
        
        self.usuario_selecionado = None
        
        # Frame Principal com Padding
        container = tk.Frame(self, bg="#ffffff", padx=35, pady=35,
                             highlightbackground="#d0d4da", highlightthickness=1)
        container.pack(expand=True, padx=15, pady=15, fill=tk.BOTH)

        tk.Label(
            container, 
            text="Sistema Numerador", 
            font=("Segoe UI", 14, "bold"), 
            bg="#ffffff", 
            fg="#1a2638"
        ).pack(pady=(0, 4))
        
        tk.Label(
            container, 
            text="Identifique-se para continuar", 
            font=("Segoe UI", 9), 
            bg="#ffffff", 
            fg="#8492a6"
        ).pack(pady=(0, 20))

        self.usuarios = db.get_all_usuarios()
        
        self.var_usuario = tk.StringVar()
        style = ttk.Style()
        style.configure("Custom.TCombobox", padding=5)
        
        self.combo_usuarios = ttk.Combobox(
            container, 
            textvariable=self.var_usuario, 
            values=self.usuarios,
            font=("Segoe UI", 11),
            style="Custom.TCombobox",
            state="readonly"
        )
        self.combo_usuarios.pack(fill=tk.X, pady=(0, 20), ipady=5)
        
        if self.usuarios:
            self.combo_usuarios.current(0)

        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=9)

        btn_entrar = ttk.Button(container, text="Entrar", style="Primary.TButton", command=self.entrar)
        btn_entrar.pack(fill=tk.X, pady=(0, 12))
        
        btn_frame = tk.Frame(container, bg="#ffffff")
        btn_frame.pack(fill=tk.X)
        
        style.configure("Secondary.TButton", font=("Segoe UI", 9), padding=5)
        
        btn_novo = ttk.Button(btn_frame, text="Novo Usuario", style="Secondary.TButton", command=self.novo_usuario)
        btn_novo.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        btn_excluir = ttk.Button(btn_frame, text="Excluir Usuario", style="Secondary.TButton", command=self.excluir_usuario)
        btn_excluir.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0))

        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

        self.focus_force()
        self.grab_set()
        self.bind('<Return>', lambda e: self.entrar())

    def entrar(self):
        selecionado = self.var_usuario.get()
        if not selecionado:
            messagebox.showwarning("Acesso Negado", "Selecione um usuário para entrar.", parent=self)
            return
            
        from network_lock import acquire_lock, start_lock_heartbeat
        sucesso, msg = acquire_lock(selecionado)
        
        if not sucesso:
            messagebox.showerror("Trancado", msg, parent=self)
            return
            
        start_lock_heartbeat(selecionado)
        self.usuario_selecionado = selecionado
        self.destroy()

    def novo_usuario(self):
        novo = simpledialog.askstring("Novo Usuário Administrativo", "Digite o nome exato (Matrícula ou Nome):", parent=self)
        if novo and novo.strip():
            novo = novo.strip().upper()
            try:
                db.add_usuario(novo)
                self.usuarios = db.get_all_usuarios()
                self.combo_usuarios.config(values=self.usuarios)
                self.combo_usuarios.set(novo)
                messagebox.showinfo("Sucesso", f"Usuário '{novo}' gravado na rede!", parent=self)
            except Exception as e:
                messagebox.showerror("Erro de I/O", f"Falha ao gravar arquivo JSON:\n{e}", parent=self)

    def excluir_usuario(self):
        selecionado = self.var_usuario.get()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione algo para excluir.", parent=self)
            return
        if selecionado in ['VIA DCT', 'DIRETORIA']:
            messagebox.showwarning("Proibido", "A base estrutural não permite excluir os usuários chefes.", parent=self)
            return
        if messagebox.askyesno("Confirmar Exclusão", f"Deseja excluir o usuário '{selecionado}'? Esta ação encerrará a sessão dele.", parent=self):
            db.delete_usuario(selecionado)
            self.usuarios = db.get_all_usuarios()
            self.combo_usuarios.config(values=self.usuarios)
            if self.usuarios:
                self.combo_usuarios.current(0)
            else:
                self.var_usuario.set("")
            messagebox.showinfo("Usuário Excluído", f"Usuário '{selecionado}' foi removido com sucesso.", parent=self)


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
        top_container = ttk.Frame(self, padding="20 20 20 10")
        top_container.grid(row=0, column=0, sticky="nsew")
        top_container.columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Custom.TLabelframe", background="#ffffff", bordercolor="#d1d5db")
        style.configure("Custom.TLabelframe.Label", font=("Segoe UI", 11, "bold"), foreground="#2c3e50")

        form_frame = ttk.LabelFrame(top_container, text=f"  Novo Registro: {self.titulo_aba}  ", padding="15", style="Custom.TLabelframe")
        form_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 15))
        
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
        self.entry_assunto = ttk.Combobox(form_frame, textvariable=self.var_assunto, font=("Segoe UI", 10), style="Custom.TCombobox")
        self.entry_assunto.grid(row=1, column=1, columnspan=7, sticky="we", pady=5)
        
        # Linha 3
        ttk.Label(form_frame, text="Destino:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.var_destino = tk.StringVar()
        self.entry_destino = ttk.Combobox(form_frame, textvariable=self.var_destino, font=("Segoe UI", 10), style="Custom.TCombobox")
        self.entry_destino.grid(row=2, column=1, columnspan=7, sticky="we", pady=5)
        
        # Linha 4
        ttk.Label(form_frame, text="Observações:").grid(row=3, column=0, sticky=tk.W, pady=(5, 0))
        self.var_obs = tk.StringVar()
        self.entry_obs = ttk.Entry(form_frame, textvariable=self.var_obs, font=("Segoe UI", 10))
        self.entry_obs.grid(row=3, column=1, columnspan=7, sticky="we", pady=(5, 0))
        
        # Botoes
        btn_container = ttk.Frame(top_container)
        btn_container.grid(row=1, column=0, sticky="we", pady=(5, 5))
        
        style.configure("Action.TButton", font=("Segoe UI", 10), padding=6)
        
        grupo_registro = ttk.LabelFrame(btn_container, text=" Ações Primárias ", padding="10", style="Custom.TLabelframe")
        grupo_registro.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        ttk.Button(grupo_registro, text="Limpar / Novo", style="Action.TButton", command=self.action_novo).pack(fill=tk.X, pady=3)
        ttk.Button(grupo_registro, text="Salvar Alterações", style="Action.TButton", command=self.action_salvar).pack(fill=tk.X, pady=3)
        
        grupo_docs = ttk.LabelFrame(btn_container, text=" Exportação DOCX ", padding="10", style="Custom.TLabelframe")
        grupo_docs.pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(grupo_docs, text="Gerar Relatório DOCX", style="Action.TButton", command=self.action_abrir_numerador).pack(fill=tk.X, pady=3)
        ttk.Button(grupo_docs, text="Abrir Pasta de Saída", style="Action.TButton", command=self.action_abrir_pasta).pack(fill=tk.X, pady=3)

        # --- Modelos Favoritos ---
        grupo_modelos = ttk.LabelFrame(btn_container, text=" Modelos Frequentes ⭐ ", padding="10", style="Custom.TLabelframe")
        grupo_modelos.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        self.var_modelo = tk.StringVar()
        self.combo_modelos = ttk.Combobox(grupo_modelos, textvariable=self.var_modelo, font=("Segoe UI", 10), state="readonly", width=18)
        self.combo_modelos.pack(fill=tk.X, padx=3, pady=3)
        self.combo_modelos.bind("<<ComboboxSelected>>", self.action_carregar_modelo)
        
        frm_botoes_mod = ttk.Frame(grupo_modelos)
        frm_botoes_mod.pack(fill=tk.X, pady=3)
        ttk.Button(frm_botoes_mod, text="Salvar Atual", command=self.action_salvar_modelo, width=12).pack(side=tk.LEFT, padx=(0,2))
        ttk.Button(frm_botoes_mod, text="X", command=self.action_excluir_modelo, width=3).pack(side=tk.LEFT)

        # --- Gerenciamento (Lado Direito) ---
        grupo_sys = ttk.LabelFrame(btn_container, text=" Gerenciamento ", padding="10", style="Custom.TLabelframe")
        grupo_sys.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        ttk.Button(grupo_sys, text="Sincronizar Agora", style="Action.TButton", command=self.app.sincro_manual).pack(side=tk.LEFT, fill=tk.X, padx=3, pady=3)
        ttk.Button(grupo_sys, text="Trocar Usuário", style="Action.TButton", command=self.app.action_trocar_usuario).pack(side=tk.LEFT, fill=tk.X, padx=3, pady=3)
        if self.app.role == "admin":
            ttk.Button(grupo_sys, text="Excluir Seleção", style="Action.TButton", command=self.action_excluir).pack(side=tk.LEFT, fill=tk.X, padx=3, pady=3)
        
        # --- TABELA Inferior ---
        bottom_container = ttk.Frame(self, padding="20 0 20 20")
        bottom_container.grid(row=1, column=0, sticky="nsew")
        bottom_container.rowconfigure(1, weight=1)
        bottom_container.columnconfigure(0, weight=1)

        search_frame = ttk.Frame(bottom_container)
        search_frame.grid(row=0, column=0, sticky="we", pady=(10, 5))
        
        ttk.Label(search_frame, text="Pesquisa Textual:").pack(side=tk.LEFT, padx=(0, 5))
        self.var_busca = tk.StringVar()
        self.entry_busca = ttk.Entry(search_frame, textvariable=self.var_busca, width=28, font=("Segoe UI", 10))
        self.entry_busca.pack(side=tk.LEFT)
        self.entry_busca.bind('<Return>', lambda e: self.load_data())
        
        ttk.Label(search_frame, text="De:").pack(side=tk.LEFT, padx=(10, 5))
        self.var_data_inicio = tk.StringVar()
        self.entry_data_inicio = ttk.Entry(search_frame, textvariable=self.var_data_inicio, width=12, font=("Segoe UI", 10))
        self.entry_data_inicio.pack(side=tk.LEFT)
        self.entry_data_inicio.bind('<Return>', lambda e: self.load_data())
        
        ttk.Label(search_frame, text="Até:").pack(side=tk.LEFT, padx=(10, 5))
        self.var_data_fim = tk.StringVar()
        self.entry_data_fim = ttk.Entry(search_frame, textvariable=self.var_data_fim, width=12, font=("Segoe UI", 10))
        self.entry_data_fim.pack(side=tk.LEFT)
        self.entry_data_fim.bind('<Return>', lambda e: self.load_data())

        ttk.Button(search_frame, text="Buscar", command=self.load_data, width=10).pack(side=tk.LEFT, padx=10)
        ttk.Button(search_frame, text="X", command=self.limpar_busca, width=3).pack(side=tk.LEFT)

        banco_mode = "REDE (G:)" if db.settings.get_rede_path() is not None else "LOCAL (Fallback)"
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
        self.tree.tag_configure('newrow', background='#d4edda', foreground='#155724')
        
        self.update_autocompletes()
        self.update_modelos()

    def update_modelos(self):
        try:
            self.modelos_dict = db.get_modelos(self.tipo_db)
            nomes = list(self.modelos_dict.keys())
            self.combo_modelos['values'] = nomes
            if nomes:
                self.var_modelo.set("Selecione um Modelo")
            else:
                self.var_modelo.set("Sem modelos")
        except: pass

    def action_carregar_modelo(self, event=None):
        nome = self.var_modelo.get()
        if hasattr(self, 'modelos_dict') and nome in self.modelos_dict:
            mod = self.modelos_dict[nome]
            self.var_assunto.set(mod.get("assunto", ""))
            self.var_destino.set(mod.get("destino", ""))
            self.var_obs.set(mod.get("obs", ""))

    def action_salvar_modelo(self):
        ass = self.var_assunto.get().strip()
        dest = self.var_destino.get().strip()
        obs = self.var_obs.get().strip()
        if not ass and not dest:
            messagebox.showwarning("Aviso", "Preencha ao menos o Assunto ou Destino para criar um modelo.")
            return

        nome_modelo = simpledialog.askstring("Novo Modelo", "Digite um nome curto para este modelo (ex: 'Ofício de Compras'):", parent=self)
        if nome_modelo:
            nome_modelo = nome_modelo.strip()
            if nome_modelo:
                db.save_modelo(self.tipo_db, nome_modelo, ass, dest, obs)
                self.update_modelos()
                self.var_modelo.set(nome_modelo)
                messagebox.showinfo("Sucesso", f"Modelo '{nome_modelo}' salvo com sucesso!")

    def action_excluir_modelo(self):
        nome = self.var_modelo.get()
        if hasattr(self, 'modelos_dict') and nome in self.modelos_dict:
            if messagebox.askyesno("Excluir Modelo", f"Tem certeza que deseja apagar o modelo '{nome}'?"):
                db.delete_modelo(self.tipo_db, nome)
                self.update_modelos()
                messagebox.showinfo("Sucesso", "Modelo removido.")

    def update_autocompletes(self):
        try:
            self.entry_assunto['values'] = db.get_historico_assuntos(30)
            self.entry_destino['values'] = db.get_historico_destinos(30)
        except:
            pass

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
            self.update_autocompletes()
            try:
                export_docx.exportar_para_docx(self.tipo_db)
            except:
                pass
            messagebox.showinfo("Sucesso", "Registrado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao registrar:\n{e}")
            
    def action_excluir(self):
        if self.app.role != "admin":
            messagebox.showwarning("Acesso Negado", "Você não tem permissão para excluir registros.")
            return

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

    def action_abrir_pasta(self):
        try:
            pasta = export_docx.get_active_output_dir(self.tipo_db)
            if not os.path.exists(pasta):
                os.makedirs(pasta)
            os.startfile(pasta)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possivel abrir a pasta:\n{e}")
            
    def action_abrir_numerador(self):
        try:
            caminho = export_docx.exportar_para_docx(self.tipo_db)
            os.startfile(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar e abrir o DOCX:\n{e}")
            
    def load_data(self, silent=False):
        busca = self.var_busca.get().strip() if hasattr(self, 'var_busca') else ""
        data_i = self.var_data_inicio.get().strip() if hasattr(self, 'var_data_inicio') else ""
        data_f = self.var_data_fim.get().strip() if hasattr(self, 'var_data_fim') else ""
        registros = db.get_all_registros(self.tipo_db, busca, data_i, data_f)
        
        novos = set()
        if silent:
            current_ids = [self.tree.item(item)['values'][0] for item in self.tree.get_children()]
            new_ids = [r[0] for r in registros]
            if current_ids == new_ids:
                return
            novos = set(new_ids) - set(current_ids)
            if novos:
                try: 
                    import winsound
                    winsound.MessageBeep(winsound.MB_ICONINFORMATION)
                except: pass

        for row in self.tree.get_children():
            self.tree.delete(row)
            
        for i, r in enumerate(registros):
            p_id, numero, placa, data, assunto, destino, obs, usuario = r
            numero_fmt = f"{numero:03d}"
            
            if silent and p_id in novos:
                tags = ('newrow',)
            else:
                tags = ('evenrow',) if i % 2 == 0 else ('oddrow',)
            
            if self.tipo_db == "CERTIDAO":
                self.tree.insert("", tk.END, values=(p_id, numero_fmt, placa, data, assunto, destino, obs, usuario), tags=tags)
            else:
                self.tree.insert("", tk.END, values=(p_id, numero_fmt, data, assunto, destino, obs, usuario), tags=tags)

        if novos:
            self.after(4000, self.remover_destaques)

    def remover_destaques(self):
        try:
            for i, item in enumerate(self.tree.get_children()):
                tags = self.tree.item(item, "tags")
                if 'newrow' in tags:
                    self.tree.item(item, tags=('evenrow',) if i % 2 == 0 else ('oddrow',))
        except: pass
            
    def limpar_busca(self):
        self.var_busca.set("")
        self.var_data_inicio.set("")
        self.var_data_fim.set("")
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

class TabDashboard(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        container = ttk.Frame(self, padding="30 30 30 30")
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        
        # Titulo
        ttk.Label(container, text="Visão Geral do Mês Corrente", font=("Segoe UI", 18, "bold"), foreground="#2c3e50").grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky="w")
        
        # Cards (Esquerda)
        frm_cards = ttk.Frame(container)
        frm_cards.grid(row=1, column=0, sticky="nw", padx=(0, 20))
        
        self.lbl_total = ttk.Label(frm_cards, text="Carregando...", font=("Segoe UI", 14, "bold"), foreground="#0056b3")
        self.lbl_total.pack(anchor="w", pady=5)
        
        self.lbl_usuarios = ttk.Label(frm_cards, text="...", font=("Segoe UI", 12), foreground="gray")
        self.lbl_usuarios.pack(anchor="w", pady=5)
        
        self.lbl_detalhes = ttk.Label(frm_cards, text="", font=("Segoe UI", 11), foreground="#444")
        self.lbl_detalhes.pack(anchor="w", pady=10)
        
        # Top Destinos (Direita)
        frm_destinos = ttk.LabelFrame(container, text=" Destinos Frequentes (Geral) ", padding="15", style="Custom.TLabelframe")
        frm_destinos.grid(row=1, column=1, sticky="ne")
        
        self.list_destinos = tk.Listbox(frm_destinos, font=("Segoe UI", 10), height=8, width=45,
                                        selectbackground="#dee2e6", selectforeground="#1a2638", activestyle="none", bg="#f9f9f9")
        self.list_destinos.pack(fill=tk.BOTH, expand=True)
        
        ttk.Button(container, text="Atualizar Painel", command=self.load_data).grid(row=2, column=0, pady=20, sticky="w")
        
        self.load_data()
        
    def load_data(self):
        stats = db.get_estatisticas_dashboard()
        
        total = stats.get("total_mes", 0)
        users = stats.get("usuarios_ativos", 0)
        
        self.lbl_total.config(text=f"Total Emitidos (Mês Atual): {total:03d}")
        self.lbl_usuarios.config(text=f"Usuários ativos neste mês: {users:02d}")
        
        detalhes = "Distribuição por Categoria:\n"
        for t, q in stats.get("por_tipo", {}).items():
            detalhes += f" • {t}: {q:03d}\n"
        self.lbl_detalhes.config(text=detalhes)
        
        self.list_destinos.delete(0, tk.END)
        for dest, qtd in db.get_top_destinos(8):
            self.list_destinos.insert(tk.END, f"{qtd:03d} envios - {dest}")

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

class TabOpcoes(ttk.Frame):
    """Aba de opcoes e administracao do Sistema Numerador, espelhando o Credencial."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app

        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)

        self._build_ui()
        self.refresh()

    def _build_ui(self):
        style = ttk.Style()
        style.configure("Opts.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#1a2638")

        # ---- LINHA 1 – STATUS DA REDE ----------------------------------------
        frm_rede = ttk.LabelFrame(self, text=" Status de Rede ", padding="15", style="Opts.TLabelframe")
        frm_rede.grid(row=0, column=0, sticky="nsew", padx=(20, 10), pady=(20, 10))

        self.lbl_rede = ttk.Label(frm_rede, text="...", font=("Segoe UI", 10))
        self.lbl_rede.pack(anchor="w")
        self.lbl_db = ttk.Label(frm_rede, text="...", font=("Segoe UI", 9), foreground="gray")
        self.lbl_db.pack(anchor="w", pady=(4, 0))

        # ---- LINHA 1 – BACKUP MANUAL -----------------------------------------
        frm_backup = ttk.LabelFrame(self, text=" Backup Manual ", padding="15", style="Opts.TLabelframe")
        frm_backup.grid(row=0, column=1, sticky="nsew", padx=(10, 20), pady=(20, 10))

        ttk.Label(frm_backup, text="O backup automático roda a cada 15 min.\nClique abaixo para acionar imediatamente.",
                  font=("Segoe UI", 9), foreground="gray").pack(anchor="w", pady=(0, 10))
        ttk.Button(frm_backup, text="Executar Backup Agora",
                   command=self._fazer_backup_manual).pack(fill=tk.X, pady=2)
        ttk.Button(frm_backup, text="Abrir Pasta de Backups",
                   command=self._abrir_pasta_backup).pack(fill=tk.X, pady=2)

        self.lbl_ultimo_backup = ttk.Label(frm_backup, text="", font=("Segoe UI", 9), foreground="gray")
        self.lbl_ultimo_backup.pack(anchor="w", pady=(8, 0))

        # ---- LINHA 2 – USUARIOS REGISTRADOS ------------------------------------
        frm_users = ttk.LabelFrame(self, text=" Usuários Cadastrados ", padding="15", style="Opts.TLabelframe")
        frm_users.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=10)

        self.listbox_users = tk.Listbox(frm_users, font=("Segoe UI", 10), height=5,
                                        selectbackground="#dee2e6", selectforeground="#1a2638",
                                        activestyle="none")
        self.listbox_users.pack(fill=tk.BOTH, expand=True)

        btn_frame_u = ttk.Frame(frm_users)
        btn_frame_u.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_frame_u, text="Adicionar", command=self._add_user).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame_u, text="Remover Selecionado", command=self._del_user).pack(side=tk.LEFT)

        # ---- LINHA 2 – SESSOES ATIVAS -----------------------------------------
        frm_sess = ttk.LabelFrame(self, text=" Sessões Ativas na Rede ", padding="15", style="Opts.TLabelframe")
        frm_sess.grid(row=1, column=1, sticky="nsew", padx=(10, 20), pady=10)

        cols = ("usuario", "status", "ultimo_ping", "maquina")
        self.tree_sess = ttk.Treeview(frm_sess, columns=cols, show="headings", height=5)
        self.tree_sess.heading("usuario",    text="Usuário")
        self.tree_sess.heading("status",     text="Status")
        self.tree_sess.heading("ultimo_ping",text="Último Ping")
        self.tree_sess.heading("maquina",    text="Máquina")
        self.tree_sess.column("usuario",    width=120)
        self.tree_sess.column("status",     width=70,  anchor=tk.CENTER)
        self.tree_sess.column("ultimo_ping",width=150, anchor=tk.CENTER)
        self.tree_sess.column("maquina",    width=120)
        self.tree_sess.pack(fill=tk.BOTH, expand=True)

        ttk.Button(frm_sess, text="Atualizar Lista",
                   command=self.refresh).pack(anchor="w", pady=(8, 0))

        # ---- LINHA 3 – LOG DE BACKUPS ----------------------------------------
        frm_log = ttk.LabelFrame(self, text=" Histórico de Backups ", padding="15", style="Opts.TLabelframe")
        frm_log.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=20, pady=(0, 20))

        self.listbox_bk = tk.Listbox(frm_log, font=("Segoe UI", 9), height=6,
                                     selectbackground="#dee2e6", selectforeground="#1a2638",
                                     activestyle="none")

        sb = ttk.Scrollbar(frm_log, orient=tk.VERTICAL, command=self.listbox_bk.yview)
        self.listbox_bk.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox_bk.pack(fill=tk.BOTH, expand=True)

        ttk.Button(frm_log, text="Atualizar", command=self._load_backups).pack(anchor="w", pady=(8, 0))

    # ------------------------------------------------------------------
    def refresh(self):
        self._load_rede_status()
        self._load_usuarios()
        self._load_sessoes()
        self._load_backups()

    def _load_rede_status(self):
        import settings as sett
        rede = sett.get_rede_path()
        if rede:
            self.lbl_rede.config(text=f"Rede: CONECTADO  ({rede})", foreground="#2e7d32")
            self.lbl_db.config(text=f"Banco ativo: {sett.get_db_path()}")
        else:
            self.lbl_rede.config(text="Rede: DESCONECTADO  (usando banco local)", foreground="#c62828")
            self.lbl_db.config(text=f"Banco local: {sett.LOCAL_DB_FILE}")

    def _load_usuarios(self):
        self.listbox_users.delete(0, tk.END)
        for u in db.get_all_usuarios():
            self.listbox_users.insert(tk.END, u)

    def _load_sessoes(self):
        for row in self.tree_sess.get_children():
            self.tree_sess.delete(row)
        try:
            import sqlite3, settings as sett
            db_path = sett.get_db_path()
            with sqlite3.connect(db_path, timeout=5) as con:
                con.row_factory = sqlite3.Row
                rows = con.execute(
                    "SELECT user_name, status, last_ping, host_name FROM active_sessions ORDER BY last_ping DESC"
                ).fetchall()
            for r in rows:
                tag = "ocupado" if (r["status"] or "").lower() == "ocupado" else ""
                self.tree_sess.insert("", tk.END, values=(r["user_name"], r["status"], r["last_ping"], r["host_name"] or ""), tags=(tag,))
            self.tree_sess.tag_configure("ocupado", foreground="#c62828")
        except Exception:
            pass

    def _load_backups(self):
        self.listbox_bk.delete(0, tk.END)
        import settings as sett
        from pathlib import Path

        dirs = [sett.BACKUP_DIR]
        if sett.get_rede_path():
            dirs.insert(0, sett.get_network_backup_dir())

        arquivos = []
        for d in dirs:
            try:
                arquivos += list(Path(d).glob("numerador_backup_*.sqlite"))
            except Exception:
                pass

        arquivos.sort(key=lambda x: x.stat().st_mtime, reverse=True)

        if arquivos:
            self.lbl_ultimo_backup.config(
                text=f"Ultimo backup: {arquivos[0].name}"
            )
            for f in arquivos[:30]:
                from datetime import datetime
                mtime = datetime.fromtimestamp(f.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
                self.listbox_bk.insert(tk.END, f"  {mtime}   {f.name}   [{f.parent}]")
        else:
            self.lbl_ultimo_backup.config(text="Nenhum backup encontrado.")

    def _fazer_backup_manual(self):
        try:
            from backup import perform_backup
            ok, msg = perform_backup(is_manual=True)
            if ok:
                messagebox.showinfo("Backup", msg)
            else:
                messagebox.showwarning("Backup", msg)
        except Exception as e:
            messagebox.showerror("Erro", str(e))
        self._load_backups()

    def _abrir_pasta_backup(self):
        import settings as sett, subprocess
        d = sett.get_network_backup_dir() if sett.get_rede_path() else sett.BACKUP_DIR
        try:
            subprocess.Popen(f'explorer "{d}"')
        except Exception:
            messagebox.showinfo("Caminho", str(d))

    def _add_user(self):
        novo = simpledialog.askstring("Novo Usuario", "Nome / Matricula:", parent=self)
        if novo and novo.strip():
            db.add_usuario(novo.strip().upper())
            self._load_usuarios()

    def _del_user(self):
        sel = self.listbox_users.curselection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um usuario.", parent=self)
            return
        nome = self.listbox_users.get(sel[0])
        if nome.upper() in ("VIA DCT", "DIRETORIA"):
            messagebox.showwarning("Proibido", "Nao e possivel remover usuarios nativos.", parent=self)
            return
        if messagebox.askyesno("Confirmar", f"Excluir '{nome}'?", parent=self):
            db.delete_usuario(nome)
            self._load_usuarios()



class TabLixeira(ttk.Frame):
    """Aba de Lixeira: registros excluídos com soft-delete, com restauração e exclusão definitiva."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self._build_ui()
        self.load_data()

    def _build_ui(self):
        style = ttk.Style()
        style.configure("Lixeira.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#1a2638")

        # --- Barra de ferramentas ---
        toolbar = ttk.Frame(self, padding="10 10 10 5")
        toolbar.grid(row=0, column=0, sticky="we")

        ttk.Label(toolbar, text="🗑 Lixeira", font=("Segoe UI", 12, "bold"), foreground="#c62828").pack(side=tk.LEFT)
        ttk.Label(toolbar, text="  Registros excluídos por soft-delete. Restaure ou exclua definitivamente.",
                  font=("Segoe UI", 9), foreground="gray").pack(side=tk.LEFT)

        ttk.Button(toolbar, text="Esvaziar Lixeira", command=self._esvaziar).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(toolbar, text="Excluir Definitivamente", command=self._excluir_definitivo).pack(side=tk.RIGHT, padx=5)
        ttk.Button(toolbar, text="Restaurar Registro", command=self._restaurar).pack(side=tk.RIGHT, padx=5)
        ttk.Button(toolbar, text="Atualizar", command=self.load_data).pack(side=tk.RIGHT, padx=5)

        # --- Filtro por tipo ---
        filtro_frame = ttk.Frame(self, padding="10 0 10 5")
        filtro_frame.grid(row=1, column=0, sticky="we")
        self.rowconfigure(1, weight=0)
        self.rowconfigure(2, weight=1)

        ttk.Label(filtro_frame, text="Filtrar por tipo:").pack(side=tk.LEFT, padx=(0, 5))
        self.var_tipo_filtro = tk.StringVar(value="(Todos)")
        tipos_opcoes = ["(Todos)"] + [titulo for _, titulo in TIPOS_NUMERADOR]
        self.combo_tipo = ttk.Combobox(filtro_frame, textvariable=self.var_tipo_filtro,
                                       values=tipos_opcoes, state="readonly", width=22)
        self.combo_tipo.pack(side=tk.LEFT)
        self.combo_tipo.bind("<<ComboboxSelected>>", lambda e: self.load_data())

        self.lbl_contagem = ttk.Label(filtro_frame, text="", foreground="gray", font=("Segoe UI", 9))
        self.lbl_contagem.pack(side=tk.LEFT, padx=(15, 0))

        # --- Tabela ---
        tree_frame = ttk.Frame(self, padding="10 0 10 10")
        tree_frame.grid(row=2, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        cols = ("id", "tipo", "numero", "assunto", "destino", "usuario", "deleted_at")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="browse")

        self.tree.heading("id",         text="ID")
        self.tree.heading("tipo",       text="Tipo")
        self.tree.heading("numero",     text="Nº")
        self.tree.heading("assunto",    text="Assunto")
        self.tree.heading("destino",    text="Destino")
        self.tree.heading("usuario",    text="Responsável")
        self.tree.heading("deleted_at", text="Excluído em")

        self.tree.column("id",         width=50,  stretch=tk.NO, anchor=tk.CENTER)
        self.tree.column("tipo",       width=130, stretch=tk.NO, anchor=tk.CENTER)
        self.tree.column("numero",     width=60,  stretch=tk.NO, anchor=tk.CENTER)
        self.tree.column("assunto",    width=280, stretch=tk.YES)
        self.tree.column("destino",    width=180, stretch=tk.YES)
        self.tree.column("usuario",    width=130, stretch=tk.NO, anchor=tk.W)
        self.tree.column("deleted_at", width=150, stretch=tk.NO, anchor=tk.CENTER)

        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scroll_y.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")

        self.tree.tag_configure("oddrow",  background="#FFF3F3")
        self.tree.tag_configure("evenrow", background="#FFF8F8")

    # --- Mapa de tipo_db → titulo ---
    _TIPO_PARA_DB = {titulo: tipo_db for tipo_db, titulo in TIPOS_NUMERADOR}

    def _tipo_db_filtrado(self) -> str | None:
        sel = self.var_tipo_filtro.get()
        if sel == "(Todos)":
            return None
        return self._TIPO_PARA_DB.get(sel)

    def _get_selected_id(self) -> int | None:
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return int(self.tree.item(sel[0], "values")[0])
        except Exception:
            return None

    def load_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        tipo_filtro = self._tipo_db_filtrado()
        registros = db.list_lixeira(tipo_filtro)

        # Mapa tipo_db → titulo legível
        tipo_titulo = {t: n for t, n in TIPOS_NUMERADOR}

        for i, r in enumerate(registros):
            tags = ("evenrow",) if i % 2 == 0 else ("oddrow",)
            tipo_nome = tipo_titulo.get(r["tipo"], r["tipo"])
            self.tree.insert("", tk.END, values=(
                r["id"],
                tipo_nome,
                f"{r['numero']:03d}" if r["numero"] else "-",
                r["assunto"] or "-",
                r["destino"] or "-",
                r["usuario"] or "-",
                r["deleted_at"] or "-",
            ), tags=tags)

        self.lbl_contagem.config(text=f"{len(registros)} registro(s) na lixeira")

    def _restaurar(self):
        rec_id = self._get_selected_id()
        if rec_id is None:
            messagebox.showwarning("Seleção", "Selecione um registro para restaurar.")
            return
        if messagebox.askyesno("Restaurar", f"Restaurar o registro ID {rec_id} para a lista ativa?"):
            try:
                db.restore_registro(rec_id)
                db.log_auditoria(self.app.usuario_atual, "RESTAURAR", f"Registro ID {rec_id}", target_id=rec_id)
                self.load_data()
                # Atualiza as abas de numeradores
                for tab in self.app.abas.values():
                    tab.load_data()
                messagebox.showinfo("Sucesso", f"Registro ID {rec_id} restaurado com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível restaurar:\n{e}")

    def _excluir_definitivo(self):
        rec_id = self._get_selected_id()
        if rec_id is None:
            messagebox.showwarning("Seleção", "Selecione um registro para excluir definitivamente.")
            return
        if messagebox.askyesno(
            "Exclusão Definitiva",
            f"ATENÇÃO!\nExcluir PERMANENTEMENTE o registro ID {rec_id}?\nEsta ação NÃO pode ser desfeita!",
        ):
            try:
                db.delete_permanente(rec_id)
                db.log_auditoria(self.app.usuario_atual, "EXCLUIR DEFINITIVO", f"Registro ID {rec_id}", target_id=rec_id)
                self.load_data()
                messagebox.showinfo("Excluído", f"Registro ID {rec_id} excluído permanentemente.")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível excluir:\n{e}")

    def _esvaziar(self):
        count = len(self.tree.get_children())
        if count == 0:
            messagebox.showinfo("Lixeira vazia", "Não há registros na lixeira para excluir.")
            return
        if messagebox.askyesno(
            "Esvaziar Lixeira",
            f"Excluir PERMANENTEMENTE {count} registro(s) da lixeira?\nEsta ação NÃO pode ser desfeita!",
        ):
            try:
                deleted = db.esvaziar_lixeira()
                db.log_auditoria(self.app.usuario_atual, "ESVAZIAR LIXEIRA", f"{deleted} registros excluídos permanentemente")
                self.load_data()
                messagebox.showinfo("Lixeira esvaziada", f"{deleted} registro(s) excluído(s) permanentemente.")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível esvaziar a lixeira:\n{e}")


class NumeradorApp:
    def __init__(self, root, usuario_atual):
        self.root = root
        self.usuario_atual = usuario_atual
        self.role = db.get_usuario_role(self.usuario_atual)
        self.root.title(f"Sistema Único de Numeradores 2026  |  Usuário: {self.usuario_atual}  |  Acesso: {self.role.upper()}")
        self.root.geometry("1050x750")
        self.root.minsize(1050, 750)
        
        # INTERCEPTAR O FECHAMENTO DO X DA JANELA (POPUP)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        style = ttk.Style()
        # Tema claro/branco com fonte consistente
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure(".", background="#ffffff", fieldbackground="#ffffff")
        style.configure("TFrame", background="#ffffff")
        style.configure("TLabel", font=("Segoe UI", 10), background="#ffffff")
        style.configure("TButton", font=("Segoe UI", 10), padding=5)
        style.configure("TNotebook", background="#f0f2f5")
        style.configure("TNotebook.Tab", font=("Segoe UI", 9), padding=[10, 4])
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e8eaed", foreground="#1a2638")
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28, background="#ffffff", fieldbackground="#ffffff")
        style.configure("TLabelframe", background="#ffffff")
        style.configure("TLabelframe.Label", background="#ffffff", font=("Segoe UI", 10, "bold"), foreground="#2c3e50")
        
        # O Notebook Multiplexador de Abas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        self.abas = {}
        
        # Dashboard Principal
        self.tab_dashboard = TabDashboard(self.notebook, self)
        self.notebook.add(self.tab_dashboard, text=" Gestão e Analytics ")
        
        # Abas de Edição
        for (tipo_db, titulo_aba) in TIPOS_NUMERADOR:
            tab = TabNumerador(self.notebook, self, tipo_db, titulo_aba)
            self.notebook.add(tab, text=f" {titulo_aba} ")
            self.abas[tipo_db] = tab
            
        # Aba de Relatorios Isolada
        self.tab_relatorios = TabRelatorios(self.notebook, self)
        self.notebook.add(self.tab_relatorios, text=" Relatórios ")
        
        # Aba de Auditoria / Log de ações
        self.tab_auditoria = TabAuditoria(self.notebook, self)
        self.notebook.add(self.tab_auditoria, text=" Registros ")

        # Aba de opcoes e administracao (só admin)
        self.tab_lixeira = None
        if self.role == "admin":
            self.tab_lixeira = TabLixeira(self.notebook, self)
            self.notebook.add(self.tab_lixeira, text=" 🗑 Lixeira ")

            self.tab_opcoes = TabOpcoes(self.notebook, self)
            self.notebook.add(self.tab_opcoes, text=" Opções ")
            
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
            try:
                if self.tab_lixeira:
                    self.tab_lixeira.load_data()
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
            try: self.tab_dashboard.load_data()
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

