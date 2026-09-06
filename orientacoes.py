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
    
    # --- ETAPA 2: DEFINIÇÃO DE SEÇÕES SEGUNDO O QUADRO 1 DA NORMA ZERO ---
    secoes_obrigatorias = {}
    
    if tipo_detectado == "PROTOCOLO":
        secoes_obrigatorias = {
            "OBJETIVO": "OBJETIVO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "REFERENCIAL TEÓRICO": "REFERENCIAL" in texto_total_validacao or "TEÓRICO" in texto_total_validacao,
            "DESCRIÇÃO DO PROTOCOLO": "DESCRIÇÃO DO PROTOCOLO" in texto_total_validacao or "DESCRICAO DO PROTOCOLO" in texto_total_validacao,
            "ESTRATÉGIAS DE MONITORAMENTO": "MONITORAMENTO" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao or "REFERENCIA" in texto_total_validacao
        }
    elif tipo_detectado == "POP":
        secoes_obrigatorias = {
            "DEFINIÇÃO": "DEFINIÇÃO" in texto_total_validacao or "DEFINICAO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "RESPONSÁVEL PELA EXECUÇÃO": "RESPONSÁVEL PELA EXECUÇÃO" in texto_total_validacao or "RESPONSAVEL PELA EXECUCAO" in texto_total_validacao,
            "MATERIAIS UTILIZADOS": "MATERIAIS UTILIZADOS" in texto_total_validacao or "MATERIAIS UTILIZADOS NA REALIZAÇÃO" in texto_total_validacao,
            "DESCRIÇÃO DA TAREFA/ATIVIDADE": "DESCRIÇÃO DA TAREFA" in texto_total_validacao or "DESCRICAO DA TAREFA" in texto_total_validacao,
            "ATIVIDADES CRÍTICAS": "CRÍTICAS" in texto_total_validacao or "CRITICAS" in texto_total_validacao,
            "PONTOS PROIBIDOS NA EXECUÇÃO": "PONTOS PROIBIDOS" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao or "REFERENCIA" in texto_total_validacao
        }
    elif tipo_detectado == "MANUAL":
        secoes_obrigatorias = {
            "CAPA": "CAPA" in texto_total_validacao,
            "ELABORADORES": "ELABORADORES" in texto_total_validacao,
            "COLABORADORES": "COLABORADORES" in texto_total_validacao,
            "SUMÁRIO": "SUMÁRIO" in texto_total_validacao or "SUMARIO" in texto_total_validacao,
            "APRESENTAÇÃO": "APRESENTAÇÃO" in texto_total_validacao or "APRESENTACAO" in texto_total_validacao,
            "DESCRIÇÃO": "DESCRIÇÃO" in texto_total_validacao or "DESCRICAO" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao or "REFERENCIA" in texto_total_validacao
        }
    elif tipo_detectado == "NORMA":
        secoes_obrigatorias = {
            "INTRODUÇÃO": "INTRODUÇÃO" in texto_total_validacao or "INTRODUCAO" in texto_total_validacao,
            "OBJETIVO": "OBJETIVO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "DESCRIÇÃO DA NORMA": "DESCRIÇÃO" in texto_total_validacao or "DESCRICAO" in texto_total_validacao,
            "RESPONSÁVEL": "RESPONSÁVEL" in texto_total_validacao or "RESPONSAVEL" in texto_total_validacao,
            "EFEITOS DO NÃO CUMPRIMENTO": "EFEITOS DO NÃO CUMPRIMENTO" in texto_total_validacao or "CUMPRIMENTO DA NORMA" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao or "REFERENCIA" in texto_total_validacao
        }
    elif tipo_detectado == "ROTINA":
        secoes_obrigatorias = {
            "DEFINIÇÃO": "DEFINIÇÃO" in texto_total_validacao or "DEFINICAO" in texto_total_validacao,
            "OBJETIVO": "OBJETIVO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "DESCRIÇÃO DA ROTINA": "DESCRIÇÃO" in texto_total_validacao or "DESCRICAO" in texto_total_validacao
        }
    elif tipo_detectado == "REGIMENTO":
        secoes_obrigatorias = {
            "DA FINALIDADE": "FINALIDADE" in texto_total_validacao,
            "DA COMPOSIÇÃO - MEMBROS": "COMPOSIÇÃO" in texto_total_validacao or "COMPOSICAO" in texto_total_validacao,
            "DO MANDATO": "MANDATO" in texto_total_validacao,
            "DO FUNCIONAMENTO E ORGANIZAÇÃO": "FUNCIONAMENTO" in texto_total_validacao,
            "DAS COMPETÊNCIAS": "COMPETÊNCIAS" in texto_total_validacao or "COMPETENCIAS" in texto_total_validacao,
            "DAS ATRIBUIÇÕES": "ATRIBUIÇÕES" in texto_total_validacao or "ATRIBUICOES" in texto_total_validacao,
            "DISPOSIÇÕES FINAIS": "FINAIS" in texto_total_validacao
        }
    elif tipo_detectado == "POLÍTICA INSTITUCIONAL":
        secoes_obrigatorias = {
            "INTRODUÇÃO": "INTRODUÇÃO" in texto_total_validacao or "INTRODUCAO" in texto_total_validacao,
            "OBJETIVO": "OBJETIVO" in texto_total_validacao,
            "PRINCÍPIOS": "PRINCÍPIOS" in texto_total_validacao or "PRINCIPIOS" in texto_total_validacao,
            "DIRETRIZES": "DIRETRIZES" in texto_total_validacao,
            "RESPONSABILIDADES": "RESPONSABILIDADES" in texto_total_validacao,
            "ESTRATÉGIA DE MONITORAMENTO": "MONITORAMENTO" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao or "REFERENCIA" in texto_total_validacao
        }
    elif tipo_detectado == "PLANO DE CONTINGÊNCIA":
        secoes_obrigatorias = {
            "OBJETIVO": "OBJETIVO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "DEFINIÇÃO DE TERMOS": "DEFINIÇÃO DE TERMOS" in texto_total_validacao or "DEFINICAO DE TERMOS" in texto_total_validacao,
            "IDENTIFICAÇÃO DA SITUAÇÃO ATUAL": "SITUAÇÃO ATUAL" in texto_total_validacao or "SITUACAO ATUAL" in texto_total_validacao,
            "MEDIDAS DE CONTINGÊNCIA": "MEDIDAS DE CONTINGÊNCIA" in texto_total_validacao or "MEDIDAS DE CONTINGENCIA" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao or "REFERENCIA" in texto_total_validacao
        }
    elif tipo_detectado == "PROGRAMA":
        secoes_obrigatorias = {
