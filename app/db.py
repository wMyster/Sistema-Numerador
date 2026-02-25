import sqlite3
import os
import shutil
from datetime import datetime
import docx

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data')
BACKUP_DIR = os.path.join(BASE_DIR, 'backup')
NUMERADORES_DIR = os.path.join(BASE_DIR, 'Numeradores')

# Unidade de Rede (G:)
REDE_DIR = r"G:\NUMERADORES DADOS\DATA"
REDE_DB_PATH = os.path.join(REDE_DIR, 'numerador.sqlite')
LOCAL_DB_PATH = os.path.join(DATA_DIR, 'numerador.sqlite')

def get_active_db_path():
    if os.path.exists(r"G:\\"):
        if not os.path.exists(REDE_DIR):
            try: os.makedirs(REDE_DIR)
            except: return LOCAL_DB_PATH
        return REDE_DB_PATH
    return LOCAL_DB_PATH

def init_schema(db_path):
    db_dir = os.path.dirname(db_path)
    if not os.path.exists(db_dir):
        os.makedirs(db_dir)
    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT UNIQUE NOT NULL)''')
        # Tabela central unificada
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS registros (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo TEXT NOT NULL,
                numero INTEGER NOT NULL,
                placa TEXT,
                data TEXT,
                assunto TEXT,
                destino TEXT,
                obs TEXT,
                usuario TEXT,
                created_at TEXT,
                updated_at TEXT,
                UNIQUE(tipo, numero)
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS auditoria (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_hora TEXT NOT NULL,
                usuario TEXT NOT NULL,
                acao TEXT NOT NULL,
                detalhes TEXT NOT NULL
            )
        ''')
        # Limpar registros fantasmas que estragam a numeração MAX() e vêm sujos do Word
        # Eles podem conter só " " ou "-"
        cursor.execute('''
            DELETE FROM registros 
            WHERE (assunto IS NULL OR TRIM(assunto) = '' OR TRIM(assunto) = '-') 
              AND (destino IS NULL OR TRIM(destino) = '' OR TRIM(destino) = '-') 
              AND (obs IS NULL OR TRIM(obs) = '' OR TRIM(obs) = '-') 
              AND (usuario IS NULL OR TRIM(usuario) = '' OR TRIM(usuario) = '-')
        ''')
        # Re-arrumar IDs do auto-increment pra quando todos do topo forem apagados (opcional mas bom)
        conn.commit()

def merge_dbs(db_main, db_attached):
    try:
        init_schema(db_main)
        init_schema(db_attached)
        with sqlite3.connect(db_main, timeout=10) as conn:
            conn.execute(f"ATTACH DATABASE '{db_attached}' AS aux;")
            conn.execute("INSERT OR IGNORE INTO usuarios (nome) SELECT nome FROM aux.usuarios;")
            conn.execute('''
                INSERT OR IGNORE INTO registros (tipo, numero, placa, data, assunto, destino, obs, usuario, created_at, updated_at)
                SELECT tipo, numero, placa, data, assunto, destino, obs, usuario, created_at, updated_at FROM aux.registros
                WHERE (tipo, numero) NOT IN (SELECT tipo, numero FROM registros);
            ''')
            conn.execute('''
                UPDATE registros
                SET 
                    placa = (SELECT placa FROM aux.registros WHERE tipo = registros.tipo AND numero = registros.numero),
                    data = (SELECT data FROM aux.registros WHERE tipo = registros.tipo AND numero = registros.numero),
                    assunto = (SELECT assunto FROM aux.registros WHERE tipo = registros.tipo AND numero = registros.numero),
                    destino = (SELECT destino FROM aux.registros WHERE tipo = registros.tipo AND numero = registros.numero),
                    obs = (SELECT obs FROM aux.registros WHERE tipo = registros.tipo AND numero = registros.numero),
                    usuario = (SELECT usuario FROM aux.registros WHERE tipo = registros.tipo AND numero = registros.numero),
                    updated_at = (SELECT updated_at FROM aux.registros WHERE tipo = registros.tipo AND numero = registros.numero)
                WHERE (tipo, numero) IN (
                    SELECT m.tipo, m.numero FROM registros m
                    JOIN aux.registros a ON m.tipo = a.tipo AND m.numero = a.numero
                    WHERE a.updated_at > m.updated_at
                );
            ''')
            conn.execute('''
                INSERT INTO auditoria (data_hora, usuario, acao, detalhes)
                SELECT data_hora, usuario, acao, detalhes FROM aux.auditoria
                WHERE (data_hora || acao || detalhes) NOT IN (SELECT data_hora || acao || detalhes FROM auditoria)
            ''')
            conn.commit()
    except Exception as e:
        pass 

def log_auditoria(usuario, acao, detalhes):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO auditoria (data_hora, usuario, acao, detalhes)
            VALUES (?, ?, ?, ?)
        ''', (agora, usuario, acao, detalhes))
        conn.commit()
        
def get_todas_auditorias(limit=1000):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT data_hora, usuario, acao, detalhes 
            FROM auditoria ORDER BY id DESC LIMIT ?
        ''', (limit,))
        return cursor.fetchall()

def log_debug(msg):
    # Pode usar pra debugar importacao do word se quiser
    pass

def bootstrap_from_docx_if_empty():
    """ 
    Procura os arquivos na pasta Numeradores e alimenta a nova tabela 'registros'. 
    Somente importa linhas que possuam um Nº válido (int).
    """
    mapa = {
        "NUMERADOR DE OFÍCIO 2026.docx": "OFICIO",
        "NUMERADOR DE MEMORANDO 2026.docx": "MEMORANDO",
        "NUMERADOR DE CIRCULAR INTERNA2026.docx": "CIRCULAR_INTERNA",
        "NUMERADOR DE NOTIFICAÇAO 2026.docx": "NOTIFICACAO",
        "NUMERADOR DE PORTARIA 2026.docx": "PORTARIA",
        "NUMERADOR DE AUTORIZAÇÃO PARA CONDUÇÃO DE VEÍCULO OFÍCIAL 2026.docx": "AUTORIZACAO_VEICULO",
        "NUMERADOR DE CERTIDÃO  2026.docx": "CERTIDAO"
    }
    
    with get_connection() as conn:
        c = conn.cursor()
        c.execute('SELECT COUNT(*) FROM registros')
        if c.fetchone()[0] > 0:
            return # Ja tem arquivos, nao faz o import inicial 

    if not os.path.exists(NUMERADORES_DIR):
        return
        
    for nome_arquivo, tipo_db in mapa.items():
        filepath = os.path.join(NUMERADORES_DIR, nome_arquivo)
        if os.path.exists(filepath):
            try:
                doc = docx.Document(filepath)
                if doc.tables:
                    table = doc.tables[0]
                    for row in table.rows[1:]: # pula header
                        cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                        
                        # Extrair o numero (garantir q n eh linha vazia inteira)
                        try:
                            n_str = cells[0].strip()
                            if not n_str: continue
                            numero = int(n_str)
                        except:
                            continue
                            
                        placa = ""
                        data = ""
                        assunto = ""
                        destino = ""
                        obs = ""
                        
                        if tipo_db == "CERTIDAO":
                            if len(cells) >= 6:
                                placa = cells[1]
                                data = cells[2]
                                assunto = cells[3]
                                destino = cells[4]
                                obs = cells[5]
                        else:
                            if len(cells) >= 5:
                                data = cells[1]
                                assunto = cells[2]
                                destino = cells[3]
                                obs = cells[4]
                                
                        usuario = "" # usuario costuma estar em obs, mas vamos salvar o obs puro
                        
                        try:
                            insert_registro(tipo_db, numero, placa, data, assunto, destino, obs, usuario, skip_sync=True)
                        except Exception as e:
                            log_debug(f"Erro inserindo {tipo_db} - {numero}: {e}")
            except Exception as e:
                log_debug(f"Erro ao ler DOCX {nome_arquivo}: {e}")
                
    # Ao final do bootstrap, forcar um sync pra rechear a rede/local
    sincronizar_bancos()


def sincronizar_bancos():
    rede_disponivel = os.path.exists(r"G:\\")
    if rede_disponivel and not os.path.exists(REDE_DIR):
        try: os.makedirs(REDE_DIR)
        except: rede_disponivel = False

    if rede_disponivel:
        if os.path.exists(LOCAL_DB_PATH) and os.path.exists(REDE_DB_PATH):
            merge_dbs(REDE_DB_PATH, LOCAL_DB_PATH)
            merge_dbs(LOCAL_DB_PATH, REDE_DB_PATH)
        elif os.path.exists(REDE_DB_PATH) and not os.path.exists(LOCAL_DB_PATH):
            if not os.path.exists(DATA_DIR): os.makedirs(DATA_DIR)
            shutil.copy2(REDE_DB_PATH, LOCAL_DB_PATH)
        elif os.path.exists(LOCAL_DB_PATH) and not os.path.exists(REDE_DB_PATH):
            shutil.copy2(LOCAL_DB_PATH, REDE_DB_PATH)

def get_connection():
    db_path = get_active_db_path()
    db_dir = os.path.dirname(db_path)
    if not os.path.exists(db_dir):
        os.makedirs(db_dir)
    return sqlite3.connect(db_path, timeout=10)

def init_db():
    init_schema(LOCAL_DB_PATH)
    with get_connection() as conn:
        cursor = conn.cursor()
        
        # Busca usuários atuais para não perder ninguém que foi criado manualmente
        cursor.execute("SELECT nome FROM usuarios")
        nomes_existentes = [r[0].strip().upper() for r in cursor.fetchall()]
        
        # Raspagem de nomes antigos salvos em 'obs'
        try:
            cursor.execute('''
                SELECT DISTINCT UPPER(TRIM(obs)) FROM registros 
                WHERE obs IS NOT NULL AND obs != '' AND LENGTH(obs) < 50
            ''')
            possiveis_nomes = [r[0].strip().upper() for r in cursor.fetchall()]
        except:
            possiveis_nomes = []
            
        # Limpa tabela para resetar a formatação UPPER sem conflito UNIQUE
        cursor.execute("DELETE FROM usuarios")
        
        # Junta todas as fontes e remove os indesejados (set pra tirar duplicados)
        usuarios_padrao = ['DIRETORIA', 'VIA DCT']
        todos_nomes = set(nomes_existentes + possiveis_nomes + usuarios_padrao)
        
        remover = {'ADMINISTRADOR', 'SECRETARIA', 'FERNANDO VIA DCT', 'RENATO'}
        todos_nomes = todos_nomes - remover
        
        # Re-insere já sanitizado
        for u in sorted(todos_nomes):
            if u:  # Previne strings vazias
                cursor.execute('INSERT OR IGNORE INTO usuarios (nome) VALUES (?)', (u,))
                
        # Migra as siglas/nomes erroneamente colocadas na OBS para a coluna verdadeira USUARIO 
        cursor.execute("""
            UPDATE registros 
            SET usuario = UPPER(TRIM(obs)), obs = '' 
            WHERE (usuario IS NULL OR TRIM(usuario) = '') 
            AND obs IS NOT NULL AND TRIM(obs) != '' AND LENGTH(obs) < 50
        """)
                
        conn.commit()
    sincronizar_bancos()
    bootstrap_from_docx_if_empty()

# --- USUARIOS ---
def get_all_usuarios():
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT nome FROM usuarios ORDER BY nome ASC')
        return [row[0] for row in cursor.fetchall()]

def add_usuario(nome):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('INSERT INTO usuarios (nome) VALUES (?)', (nome.upper(),))
        conn.commit()
    sincronizar_bancos()

def delete_usuario(nome):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('DELETE FROM usuarios WHERE nome = ?', (nome.upper(),))
        conn.commit()
    sincronizar_bancos()

# --- REGISTROS GERAIS ---
def get_proximo_numero(tipo):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT MAX(numero) FROM registros WHERE tipo=?', (tipo,))
        row = cursor.fetchone()
        return (row[0] + 1) if row[0] is not None else 1

def fazer_backup():
    rede_ok = os.path.exists(r"G:\\")
    if rede_ok:
        backup_folder = os.path.join(r"G:\NUMERADORES DADOS", "BACKUPS")
    else:
        backup_folder = os.path.join(DATA_DIR, "BACKUPS")
        
    if not os.path.exists(backup_folder):
        try: os.makedirs(backup_folder)
        except: return None
        
    agora = datetime.now()
    nome_backup = f'numerador_backup_{agora.strftime("%Y_%m_%d_%H%M%S")}.sqlite'
    backup_file = os.path.join(backup_folder, nome_backup)
    
    db_ativo = get_active_db_path()
    try:
        if os.path.exists(db_ativo):
            shutil.copy2(db_ativo, backup_file)
            return backup_file
    except:
        pass
    return None

def insert_registro(tipo, numero, placa, data, assunto, destino, obs, usuario, skip_sync=False):
    with get_connection() as conn:
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        cursor.execute('''
            INSERT OR IGNORE INTO registros (tipo, numero, placa, data, assunto, destino, obs, usuario, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (tipo, numero, placa, data, assunto, destino, obs, usuario, now, now))
        conn.commit()
    if not skip_sync:
        sincronizar_bancos()
        fazer_backup()

def update_registro(id_registro, placa, data, assunto, destino, obs, usuario):
    with get_connection() as conn:
        cursor = conn.cursor()
        now = datetime.now().isoformat()
        cursor.execute('''
            UPDATE registros
            SET placa=?, data=?, assunto=?, destino=?, obs=?, usuario=?, updated_at=?
            WHERE id=?
        ''', (placa, data, assunto, destino, obs, usuario, now, id_registro))
        conn.commit()
    sincronizar_bancos()
    fazer_backup()

def delete_registro(id_registro):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('DELETE FROM registros WHERE id=?', (id_registro,))
        conn.commit()
    sincronizar_bancos()
    fazer_backup()

def get_all_registros(tipo, busca=""):
    with get_connection() as conn:
        cursor = conn.cursor()
        if busca:
            termo = f'%{busca}%'
            cursor.execute('''
                SELECT id, numero, placa, data, assunto, destino, obs, usuario
                FROM registros
                WHERE tipo=? AND (assunto LIKE ? OR destino LIKE ? OR obs LIKE ? OR usuario LIKE ? OR placa LIKE ?)
                ORDER BY numero ASC
            ''', (tipo, termo, termo, termo, termo, termo))
        else:
            cursor.execute('SELECT id, numero, placa, data, assunto, destino, obs, usuario FROM registros WHERE tipo=? ORDER BY numero ASC', (tipo,))
        return cursor.fetchall()

def get_estatisticas():
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT tipo, COUNT(*) 
            FROM registros 
            WHERE (assunto IS NOT NULL AND TRIM(assunto) != '') 
               OR (destino IS NOT NULL AND TRIM(destino) != '') 
               OR (obs IS NOT NULL AND TRIM(obs) != '') 
               OR (usuario IS NOT NULL AND TRIM(usuario) != '')
            GROUP BY tipo 
            ORDER BY tipo ASC
        ''')
        return cursor.fetchall()

