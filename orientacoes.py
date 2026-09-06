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
st.title("Triagem Avançada & Gerador de Lista Mestra - NAQH")
st.markdown("""
### 🧠 Inteligência Multidocumento e Controle de Atualizações
O sistema agora valida **Protocolos, POPs e Manuais**, corrigindo o layout automaticamente. 
Abaixo, o sistema gerencia a **Lista Mestra em Excel**, impedindo duplicidades de documentos, 
exceto quando houver atualização de versão (1ª, 2ª, 3ª, etc.).
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

    # Consolidação de Texto para Varredura
    texto_corpo_completo = "".join([p.text.strip().upper() for p in doc.paragraphs])
    texto_tabelas_completo = "".join([cell.text.strip().upper() for t in doc.tables for r in t.rows for cell in r.cells])
    texto_total_validacao = texto_corpo_completo + texto_tabelas_completo

    # --- ETAPA 1: IDENTIFICAÇÃO DINÂMICA DO TIPO (EXPANDIDA) ---
    tipo_detectado = "NORMA"
    if "PROTOCOLO" in texto_total_validacao or "PROT_" in texto_total_validacao:
        tipo_detectado = "PROTOCOLO"
    elif "PROCEDIMENTO OPERACIONAL PADRÃO" in texto_total_validacao or "POP" in texto_total_validacao:
        tipo_detectado = "POP"
    elif "MANUAL" in texto_total_validacao or "MAN_" in texto_total_validacao:
        tipo_detectado = "MANUAL"
    elif "ROTINA" in texto_total_validacao or "ROT_" in texto_total_validacao:
        tipo_detectado = "ROTINA"

    st.write(f"📊 **Tipo de Documento Identificado:** {tipo_detectado}")

    # --- ETAPA 2: DEFINIÇÃO DE REQUISITOS SEGUNDO O QUADRO 1 DA NORMA ZERO ---
    secoes_obrigatorias = {}
    
    if tipo_detectado == "PROTOCOLO":
        secoes_obrigatorias = {
            "OBJETIVO": "OBJETIVO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "REFERENCIAL TEÓRICO": "REFERENCIAL" in texto_total_validacao or "TEÓRICO" in texto_total_validacao,
            "ESTRATÉGIAS DE MONITORAMENTO": "MONITORAMENTO" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao or "REFERENCIA" in texto_total_validacao
        }
    elif tipo_detectado == "POP":
        # Conforme Quadro 1 - POP exige Definição, Aplicabilidade, Responsável, Materiais, Atividades Críticas, etc.
        secoes_obrigatorias = {
            "DEFINIÇÃO": "DEFINIÇÃO" in texto_total_validacao or "DEFINICAO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "RESPONSÁVEL PELA EXECUÇÃO": "RESPONSÁVEL" in texto_total_validacao or "RESPONSAVEL" in texto_total_validacao,
            "MATERIAIS UTILIZADOS": "MATERIAIS" in texto_total_validacao,
            "ATIVIDADES CRÍTICAS": "CRÍTICAS" in texto_total_validacao or "CRITICAS" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao or "REFERENCIA" in texto_total_validacao
        }
    elif tipo_detectado == "MANUAL":
        # Conforme Quadro 1 - Manual exige Capa, Sumário, Apresentação, Descrição, etc.
        secoes_obrigatorias = {
            "SUMÁRIO": "SUMÁRIO" in texto_total_validacao or "SUMARIO" in texto_total_validacao,
            "APRESENTAÇÃO": "APRESENTAÇÃO" in texto_total_validacao or "APRESENTACAO" in texto_total_validacao,
            "DESCRIÇÃO DO MANUAL": "DESCRIÇÃO" in texto_total_validacao or "DESCRICAO" in texto_total_validacao,
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

    # --- ETAPA 3: EXTRAÇÃO E EXTRA-ABRANGÊNCIA DE CÓDIGO E VERSÃO ---
    # Extrai o código sugerido do arquivo de forma inteligente para checar duplicidade
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

    erros_gravissimos = []
    if not possui_codigo: erros_gravissimos.append("❌ **CABEÇALHO INCOMPLETO:** O campo 'CÓDIGO' está ausente.")
    if not possui_versao: erros_gravissimos.append("❌ **CABEÇALHO INCOMPLETO:** O campo 'VERSÃO' está ausente.")
    
    for secao, encontrada in secoes_obrigatorias.items():
        if not encontrada: 
            erros_gravissimos.append(f"❌ **OMISSÃO DE SEÇÃO CRÍTICA (Modelo {tipo_detectado}):** A seção obrigatória **'{secao}'** não foi localizada.")

    # --- ETAPA 4: GERENCIADOR DE DUPLICIDADE E VERSÕES DA LISTA MESTRA ---
    df_atual = st.session_state.historico_lista_mestra
    
    # Regra estrita solicitada: Evita duplicidade (Bloqueia se já existir o MESMO Código na MESMA Versão)
    is_duplicado = not df_atual[(df_atual["Código do Documento"] == codigo_doc) & (df_atual["Versão Atual"] == versao_doc)].empty
    
    if is_duplicado:
        erros_gravissimos.append(f"⛔ **BLOQUEIO DE DUPLICIDADE:** O documento **{codigo_doc}** já foi validado na versão **{versao_doc}**. Mude a versão no arquivo se isto for uma atualização!")

    # --- ETAPA 5: FORMATAÇÃO ESTÉTICA (SE APROVADO) ---
    if not erros_gravissimos:
        st.success("✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Layout liberado.")

        # 1. Registro Automático no Histórico da Lista Mestra
        nova_linha = {
            "Código do Documento": codigo_doc,
            "Título do Documento": f"Protocolo/Diretriz de {tipo_detectado}",
            "Tipo": tipo_detectado,
            "Versão Atual": versao_doc,
            "Status": "Aprovado na Triagem",
            "Data de Triagem": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Situação": "Ativo"
        }
        st.session_state.historico_lista_mestra = pd.concat([df_atual, pd.DataFrame([nova_linha])], ignore_index=True)

        # 2. Configuração Geométrica das Margens (3x3x2x2)
        for section in doc.sections:
            section.top_margin = Cm(3.0)     
            section.left_margin = Cm(3.0)    
            section.bottom_margin = Cm(2.0)  
            section.right_margin = Cm(2.0)   

        # 3. Formatação do Corpo do Texto e Listas Numéricas (Calibri 11)
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
            
