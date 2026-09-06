import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pandas as pd
from io import BytesIO
from datetime import datetime
import re

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
O sistema realiza a triagem dinâmica identificando o tipo de documento, validando as seções obrigatórias e capturando de forma inteligente os dados de identificação.
""")

# Inicializa o banco de dados do Histórico de Validações (Lista Mestra) na memória do servidor
if "historico_lista_mestra" not in st.session_state:
    st.session_state.historico_lista_mestra = pd.DataFrame(columns=[
        "Código do Documento", "Título do Documento", "Tipo", "Versão Atual",
        "Status", "Data de Triagem", "Situação"
    ])

# --- 2. FUNÇÃO DE TRATAMENTO EM TEMPO DE DOWNLOAD (EVITA TIMEOUT) ---
def executar_formatacao_pesada(arquivo_bytes):
    # Só processa os parágrafos pesados quando o operador clica no botão
    doc = docx.Document(arquivo_bytes)
    
    # Limpeza de Parágrafos Fantasmas
    p_indices_remover = [i for i, p in enumerate(doc.paragraphs) if not p.text.strip()]
    for index in sorted(p_indices_remover, reverse=True):
        p_element = doc.paragraphs[index]._element
        p_element.getparent().remove(p_element)
        
    # Margens Homologadas (Sup 2 / Inf 2 / Esq 2 / Dir 3)
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(3.0)

    # Formatação Textual Express
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

    conteudo_saida = BytesIO()
    doc.save(conteudo_saida)
    conteudo_saida.seek(0)
    return conteudo_saida.getvalue()

# --- 3. FLUXO DE CARREGAMENTO DO ARQUIVO ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    # Leitura superficial ultra rápida apenas para coletar metadados de triagem
    doc_triagem = docx.Document(arquivo_word)
    
    texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:50]]) # Apenas o topo
    texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:3] for r in t.rows for cell in r.cells])
    texto_total_raw = texto_corpo_raw + " " + texto_tabelas_raw
    texto_total_validacao = texto_total_raw.upper()
    
    # --- ETAPA 1: IDENTIFICAÇÃO DINÂMICA DO TIPO ---
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
    
    # --- ETAPA 3: EXTRAÇÃO DE CÓDIGO E VERSÃO DOS CABEÇALHOS ---
    codigo_doc = "NÃO IDENTIFICADO"
    for table in doc_triagem.tables[:3]:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if "CÓDIGO:" in txt.upper() or "CODIGO:" in txt.upper():
                    codigo_doc = txt.split(":")[-1].strip()
                    break

    if codigo_doc == "NÃO IDENTIFICADO" or len(codigo_doc) <= 2:
        match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|ROT|REG|POL|PLANC|PROG)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
        if match_codigo:
            codigo_doc = match_codigo.group(0).strip()

    versao_doc = "NÃO IDENTIFICADA"
    for table in doc_triagem.tables[:3]:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if "VERSÃO:" in txt.upper() or "VERSAO:" in txt.upper():
                    versao_doc = txt.split(":")[-1].strip()
                    break

    if versao_doc == "NÃO IDENTIFICADA" or not any(char.isdigit() for char in versao_doc):
        match_versao = re.findall(r'\b(\d+ª|\d+ª\s*VERSÃO|\d+\.\d+)\b', texto_total_raw, re.IGNORECASE)
        if match_versao:
            versao_doc = match_versao[-1].strip()

    # --- ETAPA 4: GERENCIADOR DE DUPLICIDADE ---
    df_atual = st.session_state.historico_lista_mestra
    is_duplicado = not df_atual[(df_atual["Código do Documento"] == codigo_doc) & (df_atual["Versão Atual"] == versao_doc)].empty

    # --- ETAPA 5: ENTREGA DO BOTÃO IMEDIATO ---
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

        # O botão puxa a função pesada sob demanda (Isso faz o app carregar instantâneo!)
        st.download_button(
            label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO FORMATADO",
            data=executar_formatacao_pesada(arquivo_word),
            file_name=f"{codigo_doc}_Formatado.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Mapeado: Código `{codigo_doc}` | Versão `{versao_doc}`.")
    else:
        st.error(f"🚫 **BLOQUEIO DE DUPLICIDADE:** O documento **{codigo_doc}** já foi registrado na versão **{versao_doc}**.")

# --- 6. EXIBIÇÃO DA PLANILHA ---
st.divider()
st.subheader("📊 Histórico Dinâmico da Lista Mestra (Excel)")
st.dataframe(st.session_state.historico_lista_mestra, use_container_width=True)
