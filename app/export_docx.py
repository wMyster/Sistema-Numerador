import os
import shutil
import db
from docx import Document
from docx.shared import Pt, Inches

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
REDE_OUTPUT_DIR = r"G:\NUMERADORES DADOS\OUTPUT"

TITULOS = {
    "OFICIO": "NUMERADOR DE OFÍCIO 2026",
    "MEMORANDO": "NUMERADOR DE MEMORANDO 2026",
    "CIRCULAR_INTERNA": "NUMERADOR DE CIRCULAR INTERNA 2026",
    "NOTIFICACAO": "NUMERADOR DE NOTIFICAÇÃO 2026",
    "PORTARIA": "NUMERADOR DE PORTARIA 2026",
    "AUTORIZACAO_VEICULO": "AUTORIZAÇÃO PARA CONDUÇÃO DE VEÍCULO OFICIAL 2026",
    "CERTIDAO": "NUMERADOR DE CERTIDÃO 2026"
}

NOME_UNICO = {
    "OFICIO": "Ofício",
    "MEMORANDO": "Memorando",
    "CIRCULAR_INTERNA": "Circular Interna",
    "NOTIFICACAO": "Notificação",
    "PORTARIA": "Portaria",
    "AUTORIZACAO_VEICULO": "Autorização de Veículo",
    "CERTIDAO": "Certidão"
}

def exportar_para_docx(tipo):
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        
    file_path = os.path.join(OUTPUT_DIR, f'Numerador_{tipo}_2026.docx')
    doc = Document()
    
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
    
    titulo_texto = TITULOS.get(tipo, f"NUMERADOR DE {tipo} 2026")
    titulo = doc.add_heading(str(), 0)
    run_titulo = titulo.add_run(titulo_texto)
    run_titulo.font.size = Pt(22)
    titulo.alignment = 1 
    
    registros = db.get_all_registros(tipo)
    registros = sorted(registros, key=lambda x: x[1])
    
    colunas = ['Nº', 'DATA', 'ASSUNTO', 'DESTINO', 'OBS', 'USUÁRIO']
    if tipo == "CERTIDAO":
        colunas.insert(1, 'PLACA')
        
    table = doc.add_table(rows=1, cols=len(colunas))
    table.style = 'Table Grid'
    table.allow_autofit = False
    
    hdr_cells = table.rows[0].cells
    for i, nome in enumerate(colunas):
        hdr_cells[i].text = nome
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(11)
                
    if tipo == "CERTIDAO":
        widths = [Inches(0.6), Inches(1.0), Inches(1.1), Inches(2.2), Inches(1.5), Inches(0.8), Inches(0.8)]
    else:
        widths = [Inches(0.6), Inches(1.1), Inches(2.8), Inches(1.5), Inches(1.0), Inches(1.0)]
        
    for idx, width in enumerate(widths):
        hdr_cells[idx].width = width
    
    for r in registros:
        p_id, numero, placa, data, assunto, destino, obs, usuario = r
        row_cells = table.add_row().cells
        
        row_cells[0].text = f"{numero:03d}"
        if tipo == "CERTIDAO":
            row_cells[1].text = placa or ""
            row_cells[2].text = data or ""
            row_cells[3].text = assunto or ""
            row_cells[4].text = destino or ""
            row_cells[5].text = obs or ""
            row_cells[6].text = usuario or ""
        else:
            row_cells[1].text = data or ""
            row_cells[2].text = assunto or ""
            row_cells[3].text = destino or ""
            row_cells[4].text = obs or ""
            row_cells[5].text = usuario or ""
            
        for i, width in enumerate(widths):
            row_cells[i].width = width
            
    doc.save(file_path)
    
    rede_disponivel = os.path.exists(r"G:\\")
    if rede_disponivel:
        if not os.path.exists(REDE_OUTPUT_DIR):
            try: os.makedirs(REDE_OUTPUT_DIR)
            except: pass
            
        if os.path.exists(REDE_OUTPUT_DIR):
            try:
                dest_path = os.path.join(REDE_OUTPUT_DIR, f'Numerador_{tipo}_2026.docx')
                shutil.copy2(file_path, dest_path)
            except:
                pass
                
    return file_path

