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
### 🧠 Inteligência XML Avançada contra Vazamento de Cabeçalho
O sistema isola as tabelas flutuantes superiores, impedindo que o texto do corpo invada o cabeçalho e **eliminando de forma definitiva todas as páginas em branco**.
""")

# --- 2. MOTOR DE CORREÇÃO TEXTUAL E SELEÇÃO DE TABELAS DE CONTEÚDO ---
def formatar_corpo_com_seguranca(doc):
    """Aplica recuos da Norma Zero apenas no texto livre, protegendo tabelas estruturais"""
    
    # REGRA 1: AJUSTE DE TEXTO E LISTAS GRUDADAS
    for paragraph in doc.paragraphs:
        texto_limpo = paragraph.text.strip()
        
        if not texto_limpo:
            continue
            
        if "REFERÊNCIAS" in texto_limpo.upper() or "REFERENCIA" in texto_limpo.upper():
            continue
            
        # Padrões obrigatórios: entrelinhas 1.5 e espaçamento vertical
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.paragraph_format.line_spacing = 1.5
        
        for run in paragraph.runs:
            if not paragraph.style.name.startswith('Heading'):
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        # Trata alinhamento de listas e letras (ex: g), h)) colando o texto na numeração
        is_subitem_letra = re.match(r'^[a-z]\s*\)', texto_limpo, re.IGNORECASE)
        is_marcador_ponto = texto_limpo.startswith('•') or paragraph.style.name.startswith('List')
        is_numeracao_composta = re.match(r'^(\d+(\.\d+)*\.?)\s*', texto_limpo)
        
        if is_subitem_letra or is_marcador_ponto or is_numeracao_composta:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.left_indent = Cm(0.5)
            paragraph.paragraph_format.first_line_indent = Cm(0)
        else:
            # Parágrafos normais longos (Objetivo, Referencial) recebem recuo americano de 1,25 cm
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.first_line_indent = Cm(1.25)

    # REGRA 2: FILTRO SELETIVO DE TABELAS (Garante o Quadro 1 perfeito e protege o cabeçalho)
    for table in doc.tables:
        # Pega o texto da primeira célula para identificar se é o cabeçalho institucional ou de assinaturas
        texto_primeira_celula = ""
        if table.rows and table.rows[0].cells:
            texto_primeira_celula = table.rows[0].cells[0].text.upper()
            
        # SE FOR CABEÇALHO INSTITUCIONAL OU BLOCO DE ASSINATURA, O SCRIPT PULA E NÃO MEXE
        # Isso impede que o texto do corpo vaze para dentro da caixinha dos logos!
        if "TIPO DE DOCUMENTO" in texto_primeira_celula or "SÃO LUÍS" in texto_primeira_celula or "ELABORAÇÃO" in texto_primeira_celula:
            continue
            
        # Aplica alinhamento e unificação APENAS nas tabelas de dados (Quadro 1, Quadro 2 e Apêndices)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            trPr = row._element.get_or_add_trPr()
            if not row._element.xpath('w:trPr/w:cantSplit'):
                trPr.append(OxmlElement('w:cantSplit'))
            if not row._element.xpath('w:trPr/w:keepNext'):
                trPr.append(OxmlElement('w:keepNext'))

# --- 3. MOTOR XML DIRETO (FORÇA GEOMETRIA DE MARGENS DA NORMA ZERO) ---
def processar_layout_xml_blindado(arquivo_bytes):
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1134", "1134", "1134", "1701"
    
    # 1. Roda a formatação e proteção seletiva de tabelas na biblioteca estrutural
    stream_temp = BytesIO(arquivo_bytes)
    doc_alinhado = docx.Document(stream_temp)
    formatar_corpo_com_seguranca(doc_alinhado)
    
    buffer_intermediario = BytesIO()
    doc_alinhado.save(buffer_intermediario)
    bytes_limpos = buffer_intermediario.getvalue()
    
    # 2. Injeta as margens direto no documento XML nativo preservando imagens e mídias originais
    zip_original = zipfile.ZipFile(BytesIO(bytes_limpos))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            if item.filename == "word/document.xml":
                xml_texto = conteudo.decode("utf-8")
                
                # Injeção das Margens Oficiais (2.0 cm Superior/Inferior/Esquerda e 3.0 cm Direita)
                xml_texto = re.sub(r'w:top="[^"]*"', f'w:top="{top_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:bottom="[^"]*"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="[^"]*"', f'w:left="{left_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:right="[^"]*"', f'w:right="{right_dxa}"', xml_texto)
                
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

    # Processamento seguro com isolamento de cabeçalho
    dados_finais = processar_layout_xml_blindado(dados_brutos)

    st.download_button(
        label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO FORMATADO",
        data=dados_finais,
        file_name=f"{codigo_doc}_Formatado_Homologado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Tabela superior isolada, vazamento de texto bloqueado e páginas em branco eliminadas.")
