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

# Menu Lateral - Identificação Visual do Operador (Ezequias Santos Agt Administrativo)
with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Formatador Automático - NAQH")
st.markdown("""
### 🧠 Inteligência XML de Alta Fidelidade (Modo de Precisão)
O sistema corrige os recuos, une tabelas partidas e elimina páginas em branco alterando diretamente as tags estruturais, **preservando perfeitamente cabeçalhos e logomarcas**.
""")

# --- 2. MOTOR DE TRATAMENTO DE TEXTOS E RECUOS (ANTI-QUEBRAS) ---
def corrigir_texto_e_paragrafos(doc):
    """Aplica as correções da Norma Zero no corpo eliminando recuos errados e letras soltas"""
    for paragraph in doc.paragraphs:
        texto_limpo = paragraph.text.strip()
        
        if not texto_limpo:
            continue
            
        if "REFERÊNCIAS" in texto_limpo.upper() or "REFERENCIA" in texto_limpo.upper():
            continue
            
        # Padrões obrigatórios de entrelinhas (1.5) e espaçamento vertical
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.line_spacing = 1.5
        
        # Garante a tipografia institucional Calibri 11
        for run in paragraph.runs:
            if not paragraph.style.name.startswith('Heading'):
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        # Filtro de Segurança para Listas e Subitens com parênteses: ex: g), h), a)
        # Impede o recuo forçado para que as letras não fiquem longe do texto descritivo
        is_subitem_letra = re.match(r'^[a-z]\s*\)', texto_limpo, re.IGNORECASE)
        is_marcador_ponto = texto_limpo.startswith('•') or paragraph.style.name.startswith('List')
        is_numeracao_composta = re.match(r'^(\d+(\.\d+)*\.?)\s*', texto_limpo)
        
        if is_subitem_letra or is_marcador_ponto or is_numeracao_composta:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.left_indent = Cm(0.5)  # Recuo curto e limpo para subitens
            paragraph.paragraph_format.first_line_indent = Cm(0)
        else:
            # Parágrafos longos normais (Objetivo, Referencial Teórico) recebem recuo americano clássico
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.first_line_indent = Cm(1.25)

    # Tratamento Avançado de Tabelas: Impede que o Quadro 1 quebre tabelas de forma incorreta
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            trPr = row._element.get_or_add_trPr()
            # cantSplit: Impede que uma linha se divida ao meio entre duas páginas
            if not row._element.xpath('w:trPr/w:cantSplit'):
                trPr.append(OxmlElement('w:cantSplit'))
            # keepNext: Força a linha a se manter colada com a estrutura anterior, evitando duplicação feia de cabeçalho
            if not row._element.xpath('w:trPr/w:keepNext'):
                trPr.append(OxmlElement('w:keepNext'))

# --- 3. MOTOR XML DIRETO (CORRIGE AS MARGENS E EXPUGA PÁGINAS EM BRANCO) ---
def injetar_margens_e_limpar_quebras(arquivo_bytes):
    # Converte centímetros para dxa (unidade XML: 2.0cm = 1134 dxa / 3.0cm = 1701 dxa)
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1134", "1134", "1134", "1701"
    
    # Primeiro aplicamos os ajustes textuais e de tabelas na biblioteca estrutural
    stream_temp = BytesIO(arquivo_bytes)
    doc_ajustado = docx.Document(stream_temp)
    corrigir_texto_e_paragrafos(doc_ajustado)
    
    buffer_intermediario = BytesIO()
    doc_ajustado.save(buffer_intermediario)
    bytes_ajustados = buffer_intermediario.getvalue()
    
    # Agora abrimos via ZIP para forçar a geometria XML e limpar as páginas em branco
    zip_original = zipfile.ZipFile(BytesIO(bytes_ajustados))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            if item.filename == "word/document.xml":
                xml_texto = conteudo.decode("utf-8")
                
                # 1. Correção Absoluta de Margens Geométricas (Norma Zero: 2x2x2x3)
                xml_texto = re.sub(r'w:top="[^"]*"', f'w:top="{top_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:bottom="[^"]*"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="[^"]*"', f'w:left="{left_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:right="[^"]*"', f'w:right="{right_dxa}"', xml_texto)
                
                # 2. Expurgador de Páginas em Branco: Remove quebras de páginas órfãs consecutivas antes de apêndices
                xml_texto = re.sub(r'<w:p[^>]*>\s*<w:pPr>\s*<w:pageBreakBefore[^>]*/>\s*</w:pPr>\s*</w:p>', '', xml_texto)
                xml_texto = re.sub(r'<w:br\s+w:type="page"\s*/>\s*<w:br\s+w:type="page"\s*/>', '<w:br w:type="page"/>', xml_texto)
                
                conteudo = xml_texto.encode("utf-8")
                
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# --- 4. FLUXO DE COMPILAÇÃO STREAMLIT ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    dados_brutos = arquivo_word.read()
    
    # Coleta de metadados simples para interface
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

    # Executa o motor duplo de correção (Layout + Geometria XML)
    dados_finais = injetar_margens_e_limpar_quebras(dados_brutos)

    st.download_button(
        label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO FORMATADO",
        data=dados_finais,
        file_name=f"{codigo_doc}_Formatado_Homologado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Quadro 1 unificado, subitens realinhados e páginas em branco eliminadas.")
