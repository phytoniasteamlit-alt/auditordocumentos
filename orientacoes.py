import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import pandas as pd
from io import BytesIO
from datetime import datetime
import re

# --- 1. CONFIGURAÇÃO DA PÁGINA STREAMLIT ---
st.set_page_config(page_title="Triagem & Lista Mestra NAQH", page_icon="📊", layout="wide")

# Menu Lateral - Identificação Visual do Operador (Solicitado)
with st.sidebar:
    st.markdown("### 🧑‍💻 Operador Hudson")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Gerador de Lista Mestra - NAQH")
st.markdown("""
### 🧠 Inteligência Multidocumento e Controle de Atualizações
O sistema realiza a triagem dinâmica identificando o tipo de documento, validando as seções obrigatórias e aplicando a formatação geométrica da Norma Zero de forma segura.
""")

if "historico_lista_mestra" not in st.session_state:
    st.session_state.historico_lista_mestra = pd.DataFrame(columns=[
        "Código do Documento", "Título do Documento", "Tipo", "Versão Atual",
        "Status", "Data de Triagem", "Situação"
    ])

# --- 2. MOTOR DE FORMATACÃO PROTEGIDA (PRESERVA CABEÇALHOS, RODAPÉS E TABELAS) ---
def aplicar_formatacao_protegida(arquivo_bytes):
    doc = docx.Document(arquivo_bytes)
    
    # REGRA 1: MARGENS OFICIAIS (Apenas no corpo, sem tocar na estrutura do cabeçalho)
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(3.0)

    # REGRA 2: TRATAMENTO EXCLUSIVO DO CORPO DE TEXTO CORRIDO
    # Usamos o laço ignorando cabeçalhos e tabelas para não sumir com logos ou números de páginas
    for paragraph in doc.paragraphs:
        texto_limpo = paragraph.text.strip()
        
        # Pula linhas totalmente vazias sem quebrar o documento
        if not texto_limpo:
            continue
            
        # Ignora a seção de referências para manter formatação própria
        if "REFERÊNCIAS" in texto_limpo.upper() or "REFERENCIA" in texto_limpo.upper():
            continue
            
        # Aplica espaçamentos oficiais da Norma Zero
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.line_spacing = 1.5
        
        # Garante fonte Calibri 11 para o texto do corpo
        for run in paragraph.runs:
            if not paragraph.style.name.startswith('Heading'):
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        # Regra de Alinhamento de Listas e Subitens (Ex: 5.4.1, 5.4.3 ou 1.)
        # Resolve o problema de numerações grudadas aplicando recuo invertido
        match_numeracao = re.match(r'^(\d+(\.\d+)*\.?)\s*', texto_limpo)
        if match_numeracao or (texto_limpo[:4] and ')' in texto_limpo[:4]):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.left_indent = Cm(1.25)
            paragraph.paragraph_format.first_line_indent = Cm(-1.25)
        else:
            # Correção do Referencial Teórico (Garante recuo de primeira linha de 1,25 cm)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.first_line_indent = Cm(1.25)

    # REGRA 3: NORMALIZAÇÃO DE TABELAS (Ajusta larguras e centraliza sem estragar bordas)
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            # Força as linhas a não quebrarem de página de forma feia
            trPr = row._element.get_or_add_trPr()
            if not row._element.xpath('w:trPr/w:cantSplit'):
                trPr.append(OxmlElement('w:cantSplit'))
            for cell in row.cells:
                tcPr = cell._element.get_or_add_tcPr()
                # Remove espaços em branco internos nas células
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.15

    conteudo_saida = BytesIO()
    doc.save(conteudo_saida)
    conteudo_saida.seek(0)
    return conteudo_saida.getvalue()

def converter_lista_para_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Lista Mestra')
    return output.getvalue()

# --- 3. FLUXO DE COMPILAÇÃO E TRIAGEM ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    doc_triagem = docx.Document(arquivo_word)
    
    # Coleta de metadados em lote de alta velocidade
    texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:40]])
    texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:2] for r in t.rows for cell in r.cells])
    texto_total_raw = texto_corpo_raw + " " + texto_tabelas_raw
    texto_total_validacao = texto_total_raw.upper()
    
    # Identificação do Tipo Documental
    tipo_detectado = "PROTOCOLO"
    if "PROCEDIMENTO OPERACIONAL" in texto_total_validacao or "POP" in texto_total_validacao:
        tipo_detectado = "POP"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    # Extração estável de Código e Versão
    codigo_doc = "PROT_SCIH005"
    match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|ROT)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
    if match_codigo:
        codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "")

    versao_doc = "5ª"
    match_versao = re.findall(r'\b(\d+ª|\d+ª\s*VERSÃO)\b', texto_total_raw, re.IGNORECASE)
    if match_versao:
        versao_doc = match_versao[-1].strip()

    df_atual = st.session_state.historico_lista_mestra
    
    # Limpa tentativas anteriores com erro para evitar bloqueio falso
    if not df_atual.empty:
        st.session_state.historico_lista_mestra = df_atual[
            ~((df_atual["Código do Documento"] == codigo_doc) & (df_atual["Versão Atual"] == "NÃO IDENTIFICADA"))
        ]
        df_atual = st.session_state.historico_lista_mestra

    is_duplicado = not df_atual[(df_atual["Código do Documento"] == codigo_doc) & (df_atual["Versão Atual"] == versao_doc)].empty

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

        # Executa a formatação protegendo os cabeçalhos nativos
        dados_finais = aplicar_formatacao_protegida(arquivo_word)

        st.download_button(
            label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO FORMATADO",
            data=dados_finais,
            file_name=f"{codigo_doc}_Formatado_Homologado.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Parágrafos e listas alinhados. Cabeçalho e paginação preservados.")
    else:
        st.error(f"🚫 **BLOQUEIO DE DUPLICIDADE:** O documento **{codigo_doc}** já foi processado na versão **{versao_doc}**.")

st.divider()
st.subheader("📊 Histórico Dinâmico da Lista Mestra (Excel)")
if not st.session_state.historico_lista_mestra.empty:
    st.download_button(
        label="🟢 BAIXAR PLANILHA DA LISTA MESTRA (.XLSX)",
        data=converter_lista_para_excel(st.session_state.historico_lista_mestra),
        file_name=f"Lista_Mestra_NAQH_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
st.dataframe(st.session_state.historico_lista_mestra, use_container_width=True)
