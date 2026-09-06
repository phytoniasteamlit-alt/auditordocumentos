import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Triagem NAQH", page_icon="🛡️", layout="wide")
st.title("Triagem & Engenharia de Layout Avançada - NAQH")
st.markdown("""
### 🧠 Formatação Estrita e Ajuste Físico de Elementos
O robô agora reconstrói as propriedades invisíveis do Word, removendo parágrafos fantasmas, 
alinhando numerações de listas, unificando tabelas e blindando o cabeçalho oficial.
""")

# --- 2. FUNÇÕES AUXILIARES DE ENGENHARIA DE XML (WORD) ---
def fix_table_layout(table):
    """Aplica propriedades estritas no XML da tabela para evitar quebras e desalinhamentos."""
    trPrs = table._element.xpath('//w:trPr')
    for trPr in trPrs:
        # Previne que uma linha da tabela seja cortada ao meio entre duas páginas
        cantSplit = OxmlElement('w:cantSplit')
        trPr.append(cantSplit)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Ajusta o recuo interno (padding) das células para compactar o texto."""
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# --- 3. FLUXO DE CARREGAMENTO ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para correção e alinhamento milimétrico", type=["docx"])

if arquivo_word:
    doc = docx.Document(arquivo_word)
    
    # --- PASSO 1: LIMPEZA ABSOLUTA DE ELEMENTOS FANTASMAS ---
    # Elimina linhas em branco e clarões gerados por 'Enters' desnecessários no texto original
    p_indices_remover = []
    for i, p in enumerate(doc.paragraphs):
        if not p.text.strip():
            p_indices_remover.append(i)
    for index in sorted(p_indices_remover, reverse=True):
        p_element = doc.paragraphs[index]._element
        p_element.getparent().remove(p_element)

    # Consolidação de Strings para a Varredura Inteligente
    texto_total_validacao = "".join([p.text.strip().upper() for p in doc.paragraphs])
    texto_total_validacao += "".join([cell.text.strip().upper() for t in doc.tables for r in t.rows for cell in r.cells])

    # --- PASSO 2: IDENTIFICAÇÃO DO TIPO DE DOCUMENTO ---
    tipo_detectado = "NORMA"
    if "PROTOCOLO" in texto_total_validacao or "PROT_" in texto_total_validacao:
        tipo_detectado = "PROTOCOLO"
    elif "PROCEDIMENTO OPERACIONAL PADRÃO" in texto_total_validacao or "POP" in texto_total_validacao:
        tipo_detectado = "POP"
    elif "ROTINA" in texto_total_validacao or "ROT_" in texto_total_validacao:
        tipo_detectado = "ROTINA"

    st.write(f"📊 **Tipo de Documento Identificado:** {tipo_detectado}")

    # Requisitos do Quadro 1 (Validação Dinâmica)
    secoes_obrigatorias = {}
    if tipo_detectado == "PROTOCOLO":
        secoes_obrigatorias = {
            "OBJETIVO": "OBJETIVO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "REFERENCIAL TEÓRICO": "REFERENCIAL" in texto_total_validacao or "TEÓRICO" in texto_total_validacao,
            "ESTRATÉGIAS DE MONITORAMENTO": "MONITORAMENTO" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao or "REFERENCIA" in texto_total_validacao
        }
    elif tipo_detectado == "NORMA":
        secoes_obrigatorias = {
            "INTRODUÇÃO": "INTRODUÇÃO" in texto_total_validacao,
            "OBJETIVO": "OBJETIVO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "DESCRIÇÃO DA NORMA": "DESCRIÇÃO" in texto_total_validacao,
            "RESPONSÁVEL": "RESPONSÁVEL" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao
        }

    possui_codigo = "CÓDIGO" in texto_total_validacao or "CODIGO" in texto_total_validacao or "PROT_" in texto_total_validacao
    possui_versao = "VERSÃO" in texto_total_validacao or "VERSAO" in texto_total_validacao

    erros_gravissimos = []
    if not possui_codigo: erros_gravissimos.append("❌ **CABEÇALHO INCOMPLETO:** Falta o campo 'CÓDIGO'.")
    if not possui_versao: erros_gravissimos.append("❌ **CABEÇALHO INCOMPLETO:** Falta o campo 'VERSÃO'.")
    for secao, encontrada in secoes_obrigatorias.items():
        if not encontrada: erros_gravissimos.append(f"❌ **OMISSÃO DE SEÇÃO CRÍTICA:** Seção '{secao}' não localizada.")

    # --- PASSO 3: ENGENHARIA DE RECONSTRUÇÃO ESTÉTICA ---
    if not erros_gravissimos:
        st.success("✅ **TRIAGEM APROVADA!** Reestruturando o layout...")

        # 1. Ajuste Geométrico das Margens (3,0 cm x 3,0 cm x 2,0 cm x 2,0 cm)
        for section in doc.sections:
            section.top_margin = Cm(3.0)     
            section.left_margin = Cm(3.0)    
            section.bottom_margin = Cm(2.0)  
            section.right_margin = Cm(2.0)   

        # 2. Formatação Cirúrgica do Corpo do Texto e Alinhamento de Listas
        for paragraph in doc.paragraphs:
            texto_limpo = paragraph.text.strip()
            texto_upper = texto_limpo.upper()
            
            if "REFERÊNCIAS" in texto_upper or "REFERENCIA" in texto_upper:
                continue
                
            # Limpa qualquer espaçamento ou Tab manual herdado que cause desalinhamentos
            paragraph.paragraph_format.space_before = Pt(4)
            paragraph.paragraph_format.space_after = Pt(4)
            paragraph.paragraph_format.line_spacing = 1.5
            
            # ENGENHARIA DE LISTAS: Se o parágrafo começar com números (Ex: 5.4.1, 5.1 ou a), b), c))
            # Aplica tabulação eletrônica estrita para afastar os números do texto de forma organizada
            primeiros_caracteres = texto_limpo[:8]
            if primeiros_caracteres and (primeiros_caracteres[0].isdigit() or (len(primeiros_caracteres) > 1 and primeiros_caracteres[1] == ')')):
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                paragraph.paragraph_format.left_indent = Cm(1.25)
                paragraph.paragraph_format.first_line_indent = Cm(-1.25) # Recuo deslocado (Hanging Indent)
            else:
                # Corpo do texto padrão (Alinhamento Justificado e Recuo clássico de 1,25 cm)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                paragraph.paragraph_format.left_indent = Cm(0)
                paragraph.paragraph_format.first_line_indent = Cm(1.25)
            
            for run in paragraph.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        # 3. Engenharia de Tabelas (Prevenção de Invasões e Alinhamento com a Tabela de Cima)
        largura_maxima_util = Cm(16.0) # Alinha perfeitamente todas as tabelas na mesma largura útil
        
        for idx_tabela, table in enumerate(doc.tables):
            table.autofit = False
            table.allow_autofit = False
            fix_table_layout(table) # Trava as linhas para não quebrarem feio
            
            # Se for a primeira tabela do arquivo (CABEÇALHO), aplica proteção de isolamento
            if idx_tabela == 0:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            p.paragraph_format.space_before = Pt(0)
                            p.paragraph_format.space_after = Pt(0)
                continue # Pula a formatação interna de tamanho para não mexer nos títulos oficiais do topo
            
            # Tabelas do Corpo (Quadro 1, Quadro 2 e Apêndices)
            for i, row in enumerate(table.rows):
                for cell in row.cells:
                    cell.width = largura_maxima_util
                    set_cell_margins(cell, top=80, bottom=80, left=120, right=120) # Reduz clarões internos
                    
                    for paragraph in cell.paragraphs:
                        # Zera recuos do corpo para o texto não ficar espremido na célula
                        paragraph.paragraph_format.left_indent = Cm(0)
                        paragraph.paragraph_format.first_line_indent = Cm(0)
                        paragraph.paragraph_format.line_spacing = 1.15
                        paragraph.paragraph_format.space_before = Pt(2)
                        paragraph.paragraph_format.space_after = Pt(2)
                        
                        if i == 0: # Cabeçalho interno da tabela (Calibri 10 Negrito)
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            for run in paragraph.runs:
                                run.font.name = 'Calibri'
                                run.font.size = Pt(10)
                                run.font.bold = True
                        else: # Dados (Calibri 9 Justificado)
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            for run in paragraph.runs:
                                run.font.name = 'Calibri'
                                run.font.size = Pt(9)
                                run.font.bold = False

        # 4. Ajuste Estrito de Referências (Calibri 10, Esquerda)
        capturar_referencias = False
        for paragraph in doc.paragraphs:
            if "REFERÊNCIAS" in paragraph.text.upper().strip():
                capturar_referencias = True
                continue
            if capturar_referencias and paragraph.text.strip():
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                paragraph.paragraph_format.left_indent = Cm(0)
                paragraph.paragraph_format.first_line_indent = Cm(0)
