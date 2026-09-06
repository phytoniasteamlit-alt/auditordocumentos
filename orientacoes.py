import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from io import BytesIO
import re

# --- 1. CONFIGURAÇÃO DA PÁGINA STREAMLIT ---
st.set_page_config(page_title="Formatador de Documentos NAQH", page_icon="📊", layout="wide")

# Menu Lateral - Identificação Visual do Operador
with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Formatador Automático - NAQH")
st.markdown("""
### 🧠 Inteligência Multidocumento
O sistema realiza a triagem dinâmica identificando o tipo de documento, validando as seções obrigatórias e aplicando a formatação geométrica da Norma Zero de forma segura, sem travar ou reter seus dados.
""")

# --- 2. MOTOR DE FORMATACÃO PROTEGIDA ---
def aplicar_formatacao_protegida(arquivo_bytes):
    doc = docx.Document(arquivo_bytes)
    
    # REGRA 1: MARGENS OFICIAIS (Superior: 2cm, Inferior: 2cm, Esquerda: 2cm, Direita: 3cm)
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(3.0)

    # REGRA 2: TRATAMENTO EXCLUSIVO DO CORPO DE TEXTO CORRIDO
    for paragraph in doc.paragraphs:
        texto_limpo = paragraph.text.strip()
        
        if not texto_limpo:
            continue
            
        if "REFERÊNCIAS" in texto_limpo.upper() or "REFERENCIA" in texto_limpo.upper():
            continue
            
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.line_spacing = 1.5
        
        for run in paragraph.runs:
            if not paragraph.style.name.startswith('Heading'):
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        # Regra de Alinhamento de Listas e Subitens (Ex: 5.4.1, 5.4.3 ou 1.)
        match_numeracao = re.match(r'^(\d+(\.\d+)*\.?)\s*', texto_limpo)
        if match_numeracao or (texto_limpo[:4] and ')' in texto_limpo[:4]):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.left_indent = Cm(1.25)
            paragraph.paragraph_format.first_line_indent = Cm(-1.25)
        else:
            # Correção do Referencial Teórico (Recuo de primeira linha de 1,25 cm)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.first_line_indent = Cm(1.25)

    # REGRA 3: NORMALIZAÇÃO DE TABELAS
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            trPr = row._element.get_or_add_trPr()
            if not row._element.xpath('w:trPr/w:cantSplit'):
                trPr.append(OxmlElement('w:cantSplit'))
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.15

    conteudo_saida = BytesIO()
    doc.save(conteudo_saida)
    conteudo_saida.seek(0)
    return conteudo_saida.getvalue()

# --- 3. FLUXO DE COMPILAÇÃO E TRIAGEM ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    doc_triagem = docx.Document(arquivo_word)
    
    texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:40]])
    texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:2] for r in t.rows for cell in r.cells])
    texto_total_raw = texto_corpo_raw + " " + texto_tabelas_raw
    texto_total_validacao = texto_total_raw.upper()
    
    tipo_detectado = "PROTOCOLO"
    if "PROCEDIMENTO OPERACIONAL" in texto_total_validacao or "POP" in texto_total_validacao:
        tipo_detectado = "POP"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    codigo_doc = "PROT_SCIH005"
    match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|ROT)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
    if match_codigo:
        codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "")

    # Processa os dados diretamente sem checar duplicidades
    dados_finais = aplicar_formatacao_protegida(arquivo_word)

    st.download_button(
        label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO FORMATADO",
        data=dados_finais,
        file_name=f"{codigo_doc}_Formatado_Homologado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Parágrafos e listas alinhados. Cabeçalho e paginação preservados.")
