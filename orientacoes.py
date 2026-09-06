import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
import zipfile
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO DA PÁGINA STREAMLIT ---
st.set_page_config(page_title="Formatador de Documentos NAQH", page_icon="📊", layout="wide")

# Menu Lateral - Identificação Visual do Operador
with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Formatador Automático - NAQH")
st.markdown("""
### 🧠 Inteligência XML com Deduplicação Automática
O sistema identifica e elimina parágrafos duplicados por erro de digitação, limpa quebras e ajusta as margens da Norma Zero **preservando logomarcas e cabeçalhos**.
""")

# --- 2. MOTOR DE LIMPEZA E DEDUPLICAÇÃO DE TEXTOS ---
def corrigir_e_deduplicar_texto(doc):
    """Varre o documento eliminando repetições de texto e aplicando recuos da Norma Zero"""
    textos_processados = set()
    p_indices_remover = []
    
    for i, paragraph in enumerate(doc.paragraphs):
        texto_limpo = paragraph.text.strip()
        
        if not texto_limpo:
            continue
            
        # 1. FILTRO DE DEDUPLICAÇÃO DE PARÁGRAFOS INTEIROS REPETIDOS (Caso da Paramentação)
        # Se o parágrafo for idêntico a um anterior, ele é marcado para remoção imediata
        if texto_limpo in textos_processados:
            p_indices_remover.append(i)
            continue
        textos_processados.add(texto_limpo)
        
        # 2. FILTRO DE DEDUPLICAÇÃO DE TEXTO NA MESMA LINHA (Caso do Quadro 1)
        # Remove repetições consecutivas na mesma linha (ex: Texto. Texto.)
        metade = len(texto_limpo) // 2
        if metade > 10 and texto_limpo[:metade].strip() == texto_limpo[metade:].strip():
            paragraph.text = texto_limpo[:metade].strip()
            texto_limpo = paragraph.text.strip()

        if "REFERÊNCIAS" in texto_limpo.upper() or "REFERENCIA" in texto_limpo.upper():
            continue
            
        # Configuração padrão de espaçamento vertical e entrelinhas (1.5)
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.line_spacing = 1.5
        
        for run in paragraph.runs:
            if not paragraph.style.name.startswith('Heading'):
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        # Regra de Alinhamento e Recuo de Listas, Letras e Marcadores
        is_subitem_letra = re.match(r'^[a-z]\s*\)', texto_limpo, re.IGNORECASE)
        is_marcador_ponto = texto_limpo.startswith('•') or paragraph.style.name.startswith('List')
        is_numeracao_composta = re.match(r'^(\d+(\.\d+)*\.?)\s*', texto_limpo)
        
        if is_subitem_letra or is_marcador_ponto or is_numeracao_composta:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.left_indent = Cm(0.5)
            paragraph.paragraph_format.first_line_indent = Cm(0)
        else:
            # Parágrafos longos normais (Objetivo, Referencial Teórico) recebem recuo de 1,25 cm
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.first_line_indent = Cm(1.25)

    # Executa a remoção dos parágrafos duplicados de trás para frente para não quebrar os índices
    for index in sorted(p_indices_remover, reverse=True):
        p_element = doc.paragraphs[index]._element
        p_element.getparent().remove(p_element)

    # Tratamento de Tabelas: Unifica e impede quebra de linhas entre páginas
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            trPr = row._element.get_or_add_trPr()
            if not row._element.xpath('w:trPr/w:cantSplit'):
                trPr.append(OxmlElement('w:cantSplit'))
            if not row._element.xpath('w:trPr/w:keepNext'):
                trPr.append(OxmlElement('w:keepNext'))

# --- 3. MOTOR XML DIRETO (FORÇA GEOMETRIA DE MARGENS E LIMPA ESPAÇOS) ---
def processar_pacote_completo_xml(arquivo_bytes):
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1134", "1134", "1134", "1701"
    
    # Aplica a limpeza e deduplicação no motor de estrutura primeiro
    stream_temp = BytesIO(arquivo_bytes)
    doc_limpo = docx.Document(stream_temp)
    corrigir_e_deduplicar_texto(doc_limpo)
    
    buffer_intermediario = BytesIO()
    doc_limpo.save(buffer_intermediario)
    bytes_limpos = buffer_intermediario.getvalue()
    
    # Injeta a geometria de margens direto no XML sem tocar nos mídias/cabeçalhos nativos
    zip_original = zipfile.ZipFile(BytesIO(bytes_limpos))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            if item.filename == "word/document.xml":
                xml_texto = conteudo.decode("utf-8")
                
                # Injeção das Margens da Norma Zero (2x2x2x3 cm)
                xml_texto = re.sub(r'w:top="[^"]*"', f'w:top="{top_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:bottom="[^"]*"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="[^"]*"', f'w:left="{left_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:right="[^"]*"', f'w:right="{right_dxa}"', xml_texto)
                
                # Remove quebras de página fantasmas geradas por excesso de Enters manuais
                xml_texto = re.sub(r'<w:p[^>]*>\s*<w:pPr>\s*<w:pageBreakBefore[^>]*/>\s*</w:pPr>\s*</w:p>', '', xml_texto)
                
                conteudo = xml_texto.encode("utf-8")
                
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# --- 4. FLUXO DE COMPILAÇÃO STREAMLIT ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    dados_brutos = arquivo_word.read()
    
    doc_triagem = docx.Document(BytesIO(dados_brutos))
    texto_total_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:20]]).upper()
    
    tipo_detectado = "PROTOCOLO"
    if "PROCEDIMENTO OPERACIONAL" in texto_total_raw or "POP" in texto_total_raw:
        tipo_detectado = "POP"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    codigo_doc = "PROT_SCIH005"
    match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|ROT)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
    if match_codigo:
        codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "")

    # Roda o processamento inteligente com filtro de deduplicação
    dados_finais = processar_pacote_completo_xml(dados_brutos)

    st.download_button(
        label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO FORMATADO",
        data=dados_finais,
        file_name=f"{codigo_doc}_Formatado_Homologado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Textos duplicados removidos, parágrafos alinhados e cabeçalhos preservados.")
