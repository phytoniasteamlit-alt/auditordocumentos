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
### 🧠 Inteligência Multidocumento e Controle de Atualizações
O sistema realiza a triagem dinâmica identificando o tipo de documento e aplicando a formatação geométrica da Norma Zero de forma segura, **preservando o alinhamento nativo de listas e subseções**.
""")

# --- 2. MOTOR DE FORMATACÃO GEOMÉTRICA CALIBRADO (ANTI-QUEBRAS) ---
def aplicar_formatacao_protegida(arquivo_bytes):
    doc = docx.Document(arquivo_bytes)
    
    # REGRA 1: CONFIGURAÇÃO GEOMÉTRICA DE MARGENS (Norma Zero: Sup 2 / Inf 2 / Esq 2 / Dir 3)
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(3.0)

    # REGRA 2: TRATAMENTO DO CORPO DE TEXTO CORRIDO
    for paragraph in doc.paragraphs:
        texto_limpo = paragraph.text.strip()
        
        # Pula parágrafos vazios
        if not texto_limpo:
            continue
            
        # Preserva a integridade visual da seção de referências bibliográficas
        if "REFERÊNCIAS" in texto_limpo.upper() or "REFERENCIA" in texto_limpo.upper():
            continue
            
        # Aplica os padrões obrigatórios de espaçamento vertical e entrelinhas (1.5)
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.line_spacing = 1.5
        
        # Padroniza a família tipográfica para Calibri 11
        for run in paragraph.runs:
            if not paragraph.style.name.startswith('Heading'):
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        # --- FILTRO PROTETOR DE ALINHAMENTO ---
        # Detecta se o parágrafo já é um título numerado, subitem com parêntese ou marcador de ponto
        # Se for, o script APENAS formata o texto e não mexe nas tabulações originais para não quebrar
        is_subitem_letra = re.match(r'^[a-z]\s*\)', texto_limpo, re.IGNORECASE)
        is_numeracao_composta = re.match(r'^(\d+(\.\d+)*\.?)\s*', texto_limpo)
        is_marcador_ponto = texto_limpo.startswith('•') or paragraph.style.name.startswith('List')
        
        if is_subitem_letra or is_numeracao_composta or is_marcador_ponto:
            # Mantém o alinhamento à esquerda nativo do arquivo, impedindo que as letras fiquem soltas
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            # Parágrafos de texto convencional longo (Como o Objetivo e Referencial Teórico)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            # Aplica o recuo clássico de parágrafo de 1,25 cm apenas onde não for lista
            if paragraph.paragraph_format.first_line_indent is None or paragraph.paragraph_format.first_line_indent == Cm(0):
                paragraph.paragraph_format.first_line_indent = Cm(1.25)

    # REGRA 3: MANUTENÇÃO ESTÁVEL DE TABELAS (Garante que fiquem perfeitas como o Quadro 1)
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

# --- 3. FLUXO DE PROCESSAMENTO ---
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

    # Gera a saída de alta fidelidade visual
    dados_finais = aplicar_formatacao_protegida(arquivo_word)

    st.download_button(
        label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO FORMATADO",
        data=dados_finais,
        file_name=f"{codigo_doc}_Formatado_Homologado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Documento reajustado para as margens institucionais com proteção de listas e tabelas.")
