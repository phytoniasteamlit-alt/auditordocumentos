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

# Menu Lateral - Identificação do Operador
with st.sidebar:
    st.markdown("### 🧑‍💻 Operador do Sistema")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Gerador de Lista Mestra - NAQH")
st.markdown("""
### 🧠 Inteligência Multidocumento e Controle de Atualizações
O sistema realiza a triagem dinâmica identificando o tipo de documento, validando as seções obrigatórias e aplicando a formatação geométrica da Norma Zero.
""")

if "historico_lista_mestra" not in st.session_state:
    st.session_state.historico_lista_mestra = pd.DataFrame(columns=[
        "Código do Documento", "Título do Documento", "Tipo", "Versão Atual",
        "Status", "Data de Triagem", "Situação"
    ])

# --- 2. ENGENHARIA DE XML AVANÇADA PARA CORREÇÃO DE TABELAS ---
def normalizar_tabela_institucional(table):
    """Força o ajuste automático e impede a quebra de linhas entre páginas"""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in table.rows:
        trPr = row._element.get_or_add_trPr()
        if not row._element.xpath('w:trPr/w:cantSplit'):
            trPr.append(OxmlElement('w:cantSplit'))
        for cell in row.cells:
            tcPr = cell._element.get_or_add_tcPr()
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:type'), 'auto')
            tcPr.append(tcW)

def construir_cabecalho_naqh(doc, codigo, tipo_doc):
    """Injeta a tabela estrutural de cabeçalho oficial da Prefeitura em todas as folhas"""
    section = doc.sections[0]
    header = section.header
    
    # Limpeza de resíduos textuais do cabeçalho anterior
    for p in header.paragraphs:
        p_element = p._element
        p_element.getparent().remove(p_element)
        
    # Tabela estrutural do cabeçalho (1 linha x 3 colunas)
    tbl = header.add_table(rows=1, cols=3, width=Cm(16.5))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = tbl.rows[0].cells
    
    # Aplicação de borda fina inferior de divisão via XML
    borda_xml = f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="0" w:color="B0B0B0"/></w:tcBorders>'
    for cell in hdr_cells:
        cell._element.get_or_add_tcPr().append(parse_xml(borda_xml))
        
    # Coluna 1: Brasão Textual
    p1 = hdr_cells[0].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = p1.add_run("PREFEITURA MUNICIPAL\nSÃO LUÍS - SEMUS")
    r1.font.name = 'Calibri'
    r1.font.size = Pt(9)
    r1.bold = True
    
    # Coluna 2: Nome Técnico da Diretriz
    p2 = hdr_cells[1].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"PREVENÇÃO DE INFECÇÕES\nSÍTIO CIRÚRGICO (ISC)")
    r2.font.name = 'Calibri'
    r2.font.size = Pt(10)
    r2.bold = True
    
    # Coluna 3: Identificadores (Código, Versão e Data)
    p3 = hdr_cells[2].paragraphs[0]
    p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r3 = p3.add_run(f"CÓDIGO: {codigo}\nVERSÃO: 5ª | PÁG: 1-21")
    r3.font.name = 'Calibri'
    r3.font.size = Pt(9)

# --- 3. FLUXO DE FORMATAÇÃO E REALINHAMENTO CRÍTICO ---
def executar_correcoes_esteticas(arquivo_bytes, codigo, tipo_doc):
    doc = docx.Document(arquivo_bytes)
    
    # Limpeza em lote de parágrafos fantasmas (linhas vazias)
    p_indices_remover = [i for i, p in enumerate(doc.paragraphs) if not p.text.strip()]
    for index in sorted(p_indices_remover, reverse=True):
        p_element = doc.paragraphs[index]._element
        p_element.getparent().remove(p_element)
        
    # Margens Homologadas Estritas (Sup 2 / Inf 2 / Esq 2 / Dir 3)
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(3.0)

    # Corrige e realinha todas as tabelas (Quadros, Apêndices e Histórico)
    for table in doc.tables:
        normalizar_tabela_institucional(table)

    # Realinhamento de parágrafos convencionais e listas numeradas compostas
    for paragraph in doc.paragraphs:
        texto_limpo = paragraph.text.strip()
        if "REFERÊNCIAS" in texto_limpo.upper() or "REFERENCIA" in texto_limpo.upper():
            continue
            
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.line_spacing = 1.5
        
        for run in paragraph.runs:
            run.font.name = 'Calibri'
            if not paragraph.style.name.startswith('Heading'):
                run.font.size = Pt(11)

        # Regra de Recuo Invertido: Trata subitens grudados (Ex: 5.4.1, 5.4.3, 1.)
        match_lista = re.match(r'^(\d+(\.\d+)*\.?)\s*', texto_limpo)
        if match_lista or (texto_limpo[:4] and ')' in texto_limpo[:4]):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.left_indent = Cm(1.25)
            paragraph.paragraph_format.first_line_indent = Cm(-1.25)
        else:
            # Parágrafos convencionais (Ajusta o Referencial Teórico desalinhado)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.first_line_indent = Cm(1.25)

    # Injeta a estrutura de cabeçalho padronizada
    construir_cabecalho_naqh(doc, codigo, tipo_doc)

    conteudo_saida = BytesIO()
    doc.save(conteudo_saida)
    conteudo_saida.seek(0)
    return conteudo_saida.getvalue()

def converter_lista_para_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Lista Mestra')
    return output.getvalue()

# --- 4. INTERFACE DE COMPILAÇÃO ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    doc_triagem = docx.Document(arquivo_word)
    texto_total_validacao = " ".join([p.text.strip().upper() for p in doc_triagem.paragraphs[:40]])
    
    tipo_detectado = "PROTOCOLO"
    if "PROCEDIMENTO OPERACIONAL" in texto_total_validacao or "POP" in texto_total_validacao:
        tipo_detectado = "POP"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    codigo_doc = "PROT_SCIH005"
    versao_doc = "5ª"

    df_atual = st.session_state.historico_lista_mestra
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

        # Executa a limpeza pesada e o realinhamento geométrico
        dados_formatados = executar_correcoes_esteticas(arquivo_word, codigo_doc, tipo_detectado)

        st.download_button(
            label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO REALINHADO E FORMATADO",
            data=dados_formatados,
            file_name=f"{codigo_doc}_Formatado_Homologado.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Layout corrigido conforme a Norma Zero.")
    else:
        st.error(f"🚫 **BLOQUEIO DE DUPLICIDADE:** O documento **{codigo_doc}** já está registrado na versão **{versao_doc}**.")

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
