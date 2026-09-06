import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pandas as pd
from io import BytesIO
from datetime import datetime

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Triagem & Lista Mestra NAQH", page_icon="📊", layout="wide")

# Menu Lateral - Identificação do Operador
with st.sidebar:
    st.markdown("### 🧑‍💻 Operador do Sistema")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Gerador de Lista Mestra - NAQH")
st.markdown("""
### 🧠 Inteligência Multidocumento e Controle de Atualizações
O sistema agora realiza a triagem dinâmica identificando o tipo de documento e validando se todas as **seções obrigatórias exigidas pela Norma Zero** estão presentes.
""")

# Inicializa o banco de dados do Histórico de Validações (Lista Mestra) na memória do servidor
if "historico_lista_mestra" not in st.session_state:
    st.session_state.historico_lista_mestra = pd.DataFrame(columns=[
        "Código do Documento", "Título do Documento", "Tipo", "Versão Atual",
        "Status", "Data de Triagem", "Situação"
    ])

# --- 2. FUNÇÕES AUXILIARES DE ENGENHARIA DE XML ---
def fix_table_layout(table):
    trPrs = table._element.xpath('//w:trPr')
    for trPr in trPrs:
        cantSplit = OxmlElement('w:cantSplit')
        trPr.append(cantSplit)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# --- 3. FLUXO DE CARREGAMENTO DO ARQUIVO ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    doc = docx.Document(arquivo_word)
    
    # Limpeza de Parágrafos Fantasmas (Espaços em Branco)
    p_indices_remover = []
    for i, p in enumerate(doc.paragraphs):
        if not p.text.strip():
            p_indices_remover.append(i)
            
    for index in sorted(p_indices_remover, reverse=True):
        p_element = doc.paragraphs[index]._element
        p_element.getparent().remove(p_element)
        
    # Consolidação de Texto para Varredura Completa (Corpo e Tabelas)
    texto_corpo_completo = "".join([p.text.strip().upper() for p in doc.paragraphs])
    texto_tabelas_completo = "".join([cell.text.strip().upper() for t in doc.tables for r in t.rows for cell in r.cells])
    texto_total_validacao = texto_corpo_completo + texto_tabelas_completo
    
    # --- ETAPA 1: IDENTIFICAÇÃO DINÂMICA DO TIPO (9 MODELOS DO MANUAL) ---
    tipo_detectado = "NORMA"
    if "PROTOCOLO" in texto_total_validacao or "PROT_" in texto_total_validacao:
        tipo_detectado = "PROTOCOLO"
    elif "PROCEDIMENTO OPERACIONAL PADRÃO" in texto_total_validacao or "POP" in texto_total_validacao:
        tipo_detectado = "POP"
    elif "MANUAL" in texto_total_validacao or "MAN_" in texto_total_validacao:
        tipo_detectado = "MANUAL"
    elif "ROTINA" in texto_total_validacao or "ROT_" in texto_total_validacao:
        tipo_detectado = "ROTINA"
    elif "REGIMENTO" in texto_total_validacao or "REG_" in texto_total_validacao:
        tipo_detectado = "REGIMENTO"
    elif "POLÍTICA INSTITUCIONAL" in texto_total_validacao or "POL_" in texto_total_validacao:
        tipo_detectado = "POLÍTICA INSTITUCIONAL"
    elif "PLANO DE CONTINGÊNCIA" in texto_total_validacao or "PLANC_" in texto_total_validacao:
        tipo_detectado = "PLANO DE CONTINGÊNCIA"
    elif "PROGRAMA" in texto_total_validacao or "PROG_" in texto_total_validacao:
        tipo_detectado = "PROGRAMA"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    # --- ETAPA 2: MAPEAMENTO DE REQUISITOS DA NORMA ZERO (CÉLULAS LIVRES DE ERRO) ---
    # Dicionário mestre contendo as palavras-chave mapeadas para cada tipo documental
    requisitos_mapa = {
        "PROTOCOLO": ["OBJETIVO", "APLICABILIDADE", "REFERENCIAL", "MONITORAMENTO", "REFERÊNCIAS"],
        "POP": ["DEFINIÇÃO", "APLICABILIDADE", "RESPONSÁVEL", "MATERIAIS", "TAREFA", "CRÍTICAS", "PROIBIDOS", "REFERÊNCIAS"],
        "MANUAL": ["CAPA", "ELABORADORES", "COLABORADORES", "SUMÁRIO", "APRESENTAÇÃO", "DESCRIÇÃO", "REFERÊNCIAS"],
        "NORMA": ["INTRODUÇÃO", "OBJETIVO", "APLICABILIDADE", "DESCRIÇÃO DA NORMA", "RESPONSÁVEL", "CUMPRIMENTO", "REFERÊNCIAS"],
        "ROTINA": ["DEFINIÇÃO", "OBJETIVO", "APLICABILIDADE", "DESCRIÇÃO DA ROTINA"],
        "REGIMENTO": ["FINALIDADE", "COMPOSIÇÃO", "MANDATO", "FUNCIONAMENTO", "COMPETÊNCIAS", "ATRIBUIÇÕES", "FINAIS"],
        "POLÍTICA INSTITUCIONAL": ["INTRODUÇÃO", "OBJETIVO", "PRINCÍPIOS", "DIRETRIZES", "RESPONSABILIDADES", "MONITORAMENTO", "REFERÊNCIAS"],
        "PLANO DE CONTINGÊNCIA": ["OBJETIVO", "APLICABILIDADE", "DEFINIÇÃO DE TERMOS", "SITUAÇÃO ATUAL", "CONTINGÊNCIA", "REFERÊNCIAS"],
        "PROGRAMA": ["REFERENCIAL", "PADRONIZAÇÃO", "MONITORAMENTO", "DESCRIÇÃO DO PROGRAMA", "EDUCACIONAIS", "REFERÊNCIAS"]
    }

    # Obter os termos obrigatórios correspondentes ao tipo identificado
    termos_obrigatorios = requisitos_mapa.get(tipo_detectado, [])
    erros_gravissimos = []

    # Validar se cada termo obrigatório está contido no texto consolidado
    for termo in termos_obrigatorios:
        if termo not in texto_total_validacao:
            erros_gravissimos.append(f"⚠️ **OMISSÃO DE SEÇÃO CRÍTICA (Modelo {tipo_detectado}):** A seção contendo o termo obrigatório **'{termo}'** não foi localizada no arquivo.")

    # --- ETAPA 3: EXTRAÇÃO DE CÓDIGO E VERSÃO DOS CABEÇALHOS ---
    codigo_doc = "NÃO IDENTIFICADO"
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if "CÓDIGO:" in txt.upper() or "CODIGO:" in txt.upper():
                    codigo_doc = txt.split(":")[-1].strip()

    versao_doc = "1ª"
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if "VERSÃO:" in txt.upper() or "VERSAO:" in txt.upper():
                    versao_doc = txt.split(":")[-1].strip()

    possui_codigo = codigo_doc != "NÃO IDENTIFICADO" and len(codigo_doc) > 2
    possui_versao = any(char.isdigit() for char in versao_doc)

    if not possui_codigo: 
        erros_gravissimos.append("❌ **CABEÇALHO INCOMPLETO:** O campo 'CÓDIGO' está ausente nas tabelas estruturais.")
    if not possui_versao: 
        erros_gravissimos.append("❌ **CABEÇALHO INCOMPLETO:** O campo 'VERSÃO' está ausente ou mal formatado.")

    # --- ETAPA 4: GERENCIADOR DE DUPLICIDADE ---
    df_atual = st.session_state.historico_lista_mestra
    is_duplicado = not df_atual[(df_atual["Código do Documento"] == codigo_doc) & (df_atual["Versão Atual"] == versao_doc)].empty

    if is_duplicado:
        erros_gravissimos.append(f"🚫 **BLOQUEIO DE DUPLICIDADE:** O documento **{codigo_doc}** já foi validado na versão **{versao_doc}**.")

    # --- ETAPA 5: FORMATAÇÃO ESTÉTICA E GERAÇÃO DE SAÍDA ---
    if not erros_gravissimos:
        st.success("✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Layout validado e liberado para arquivamento.")
        
        # 1. Registro Automático no Histórico da Lista Mestra
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
        
        # 2. Configuração Geométrica Oficial (Quadro 4: Sup 2 / Inf 2 / Esquer 2 / Dir 3)
        for section in doc.sections:
            section.top_margin = Cm(2.0)
            section.bottom_margin = Cm(2.0)
            section.left_margin = Cm(2.0)
            section.right_margin = Cm(3.0)

        # 3. Formatação do Corpo de Texto (Calibri 11, Espaçamento 1.5, Justificado)
        for paragraph in doc.paragraphs:
            texto_limpo = paragraph.text.strip()
            if "REFERÊNCIAS" in texto_limpo.upper() or "REFERENCIA" in texto_limpo.upper():
                continue
                
            paragraph.paragraph_format.space_before = Pt(4)
            paragraph.paragraph_format.space_after = Pt(4)
            paragraph.paragraph_format.line_spacing = 1.5
            
            primeiros_caracteres = texto_limpo[:8]
            if primeiros_caracteres and (primeiros_caracteres.replace('.', '').isdigit() or ')' in primeiros_caracteres):
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                paragraph.paragraph_format.left_indent = Cm(1.25)
                paragraph.paragraph_format.first_line_indent = Cm(-1.25)
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                paragraph.paragraph_format.left_indent = Cm(0)
                paragraph.paragraph_format.first_line_indent = Cm(1.25)

        # Preparar arquivo de download em memória
        conteudo_saida = BytesIO()
        doc.save(conteudo_saida)
        conteudo_saida.seek(0)
        
        st.download_button(
            label="📥 Baixar Documento Formatado e Homologado",
            data=conteudo_saida,
            file_name=f"{codigo_doc}_Formatado.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
