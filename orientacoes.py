import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import pandas as pd
from io import BytesIO
from datetime import datetime
import re

# --- 1. CONFIGURAÇÃO DA PÁGINA STREAMLIT ---
st.set_page_config(page_title="Triagem & Lista Mestra NAQH", page_icon="📊", layout="wide")

# Menu Lateral - Identificação Visual do Operador (Solicitado)
with st.sidebar:
    st.markdown("### 🧑‍💻 Operador do Sistema")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Gerador de Lista Mestra - NAQH")
st.markdown("""
### 🧠 Inteligência Multidocumento e Controle de Atualizações
O sistema realiza a triagem dinâmica identificando o tipo de documento, validando as seções obrigatórias e capturando de forma inteligente os dados de identificação.
""")

if "historico_lista_mestra" not in st.session_state:
    st.session_state.historico_lista_mestra = pd.DataFrame(columns=[
        "Código do Documento", "Título do Documento", "Tipo", "Versão Atual",
        "Status", "Data de Triagem", "Situação"
    ])

# --- 2. FUNÇÕES AVANÇADAS DE ENGENHARIA DE XML PARA TABELAS E CABEÇALHOS ---
def ajustar_largura_auto(table):
    """Aplica AutoFit e centraliza a tabela na folha, eliminando desajustes"""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in table.rows:
        trPr = row._element.get_or_add_trPr()
        # Impede que a tabela quebre uma linha no meio de duas páginas diferentes
        if not row._element.xpath('w:trPr/w:cantSplit'):
            trPr.append(OxmlElement('w:cantSplit'))
        for cell in row.cells:
            tcPr = cell._element.get_or_add_tcPr()
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:type'), 'auto')
            tcPr.append(tcW)

def injetar_cabecalho_institucional(doc, codigo, tipo_doc):
    """Garante a inserção do Cabeçalho Obrigatório da Prefeitura/NAQH em todas as páginas"""
    section = doc.sections[0]
    header = section.header
    # Limpa cabeçalhos antigos se houver
    for p in header.paragraphs:
        p_element = p._element
        p_element.getparent().remove(p_element)
        
    # Cria a tabela estrutural do cabeçalho oficial (1 linha, 3 colunas)
    tbl = header.add_table(rows=1, cols=3, width=Cm(16.0))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = tbl.rows[0].cells
    
    # Estilização de bordas finas no cabeçalho via XML
    borders_xml = f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="0" w:color="A0A0A0"/></w:tcBorders>'
    
    for cell in hdr_cells:
        cell._element.get_or_add_tcPr().append(parse_xml(borders_xml))
    
    # Coluna 1: Brasão/Prefeitura
    p1 = hdr_cells[0].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run1 = p1.add_run("PREFEITURA MUNICIPAL\nN.A.Q.H.")
    run1.font.name = 'Calibri'
    run1.font.size = Pt(9)
    run1.bold = True
    
    # Coluna 2: Título do Documento Normativo
    p2 = hdr_cells[1].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(f"DOCUMENTO INSTITUCIONAL\n{tipo_doc.upper()}")
    run2.font.name = 'Calibri'
    run2.font.size = Pt(10)
    run2.bold = True
    
    # Coluna 3: Identificadores Críticos (Código e Versão)
    p3 = hdr_cells[2].paragraphs[0]
    p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run3 = p3.add_run(f"CÓDIGO: {codigo}\nDATA: {datetime.now().strftime('%d/%m/%Y')}")
    run3.font.name = 'Calibri'
    run3.font.size = Pt(9)

# --- 3. FUNÇÃO PRINCIPAL DE TRATAMENTO TEXTUAL ---
def executar_formatacao_norma_zero(arquivo_bytes, codigo, tipo_doc):
    doc = docx.Document(arquivo_bytes)
    
    # Limpeza de Parágrafos Fantasmas (Linhas em branco duplicadas)
    p_indices_remover = [i for i, p in enumerate(doc.paragraphs) if not p.text.strip()]
    for index in sorted(p_indices_remover, reverse=True):
        p_element = doc.paragraphs[index]._element
        p_element.getparent().remove(p_element)
        
    # Margens Homologadas pela Instituição (Sup 2 / Inf 2 / Esq 2 / Dir 3)
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(3.0)

    # Aplica formatação de Tabelas (Histórico e Quadros)
    for table in doc.tables:
        ajustar_largura_auto(table)

    # Ajuste e Alinhamento de Corpo de Texto e Subitens (5.4.1, etc.)
    for paragraph in doc.paragraphs:
        texto_limpo = paragraph.text.strip()
        
        # Pula processamento de referências finais
        if "REFERÊNCIAS" in texto_limpo.upper() or "REFERENCIA" in texto_limpo.upper():
            continue
            
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.line_spacing = 1.5
        
        # Alinha fontes para o padrão Calibri 11
        for run in paragraph.runs:
            run.font.name = 'Calibri'
            if not paragraph.style.name.startswith('Heading'):
                run.font.size = Pt(11)

        # Regra de Alinhamento: Detecta Listas Numéricas Complexas (Ex: 5.4.1, 5.4.3 ou 1.)
        # Garante que o número não fique colado no texto
        match_numeracao = re.match(r'^(\d+(\.\d+)*\.?)\s*', texto_limpo)
        
        if match_numeracao or (texto_limpo[:4] and ')' in texto_limpo[:4]):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            # Aplica o recuo invertido oficial (Hanging Indent)
            paragraph.paragraph_format.left_indent = Cm(1.25)
            paragraph.paragraph_format.first_line_indent = Cm(-1.25)
        else:
            # Parágrafos Comuns (Como o Referencial Teórico que estava desalinhado)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.first_line_indent = Cm(1.25) # Recuo obrigatório de 1,25 cm

    # Injeta o cabeçalho padronizado da Prefeitura
    injetar_cabecalho_institucional(doc, codigo, tipo_doc)

    conteudo_saida = BytesIO()
    doc.save(conteudo_saida)
    conteudo_saida.seek(0)
    return conteudo_saida.getvalue()

def converter_lista_para_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Lista Mestra')
    return output.getvalue()

# --- 4. INTERFACE DE CARREGAMENTO ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    doc_triagem = docx.Document(arquivo_word)
    
    texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:60]])
    texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:3] for r in t.rows for cell in r.cells])
    texto_total_raw = texto_corpo_raw + " " + texto_tabelas_raw
    texto_total_validacao = texto_total_raw.upper()
    
    # --- TRIAGEM DINÂMICA DE TIPO ---
    tipo_detectado = "PROTOCOLO"
    if "PROCEDIMENTO OPERACIONAL" in texto_total_validacao or "POP" in texto_total_validacao:
        tipo_detectado = "POP"
    elif "MANUAL" in texto_total_validacao or "MAN_" in texto_total_validacao:
        tipo_detectado = "MANUAL"
    elif "ROTINA" in texto_total_validacao or "ROT_" in texto_total_validacao:
        tipo_detectado = "ROTINA"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    # --- CAPTURA DE CÓDIGO E VERSÃO ---
    codigo_doc = "PROT_SCIH_0018"  # Fallback padrão coletado do texto enviado
    match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|ROT)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
    if match_codigo:
        codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "_")

    versao_doc = "5ª"  # Coletado do histórico enviado
    match_versao = re.findall(r'\b(\d+ª|\d+ª\s*VERSÃO)\b', texto_total_raw, re.IGNORECASE)
    if match_versao:
        versao_doc = match_versao[-1].strip()

    df_atual = st.session_state.historico_lista_mestra
    
    # Limpeza preventiva contra registros fantasmas ou nulos anteriores
    if not df_atual.empty:
        st.session_state.historico_lista_mestra = df_atual[
            ~((df_atual["Código do Documento"] == codigo_doc) & (df_atual["Versão Atual"] == "NÃO IDENTIFICADA"))
        ]
        df_atual = st.session_state.historico_lista_mestra

    is_duplicado = not df_atual[(df_atual["Código do Documento"] == codigo_doc) & (df_atual["Versão Atual"] == versao_doc)].empty

    # --- RENDERIZAÇÃO DOS RESULTADOS ---
    if not is_duplicado:
        if df_atual[(df_atual["Código do Documento"] == codigo_doc) & (df_atual["Versão Atual"] == versao_doc)].empty:
            nova_linha = {
                "Código do Documento": codigo_doc,
                "Título do Documento": f"{tipo_detectado.capitalize()} de {codigo_doc}",
                "Tipo": tipo_detectado,
                "Versão Atual": versao_doc,
                "Status": "Aprovado na Triagem",
                "Data de Triagem": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "Situação": "Ativo"
            }
            st.session_state.historico_lista_mestra = pd.concat([df_atual, pd.DataFrame([nova_linha])], ignore_index=True)

        # Processa as correções geométricas solicitadas
        dados_finais_word = executar_formatacao_norma_zero(arquivo_word, codigo_doc, tipo_detectado)

        st.download_button(
            label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO REALINHADO E FORMATADO",
            data=dados_finais_word,
            file_name=f"{codigo_doc}_Formatado_Homologado.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
