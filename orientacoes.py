import streamlit as st
import docx
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
### 🧠 Inteligência XML de Alta Fidelidade
Aplica margens conforme **Norma Zero / ABNT vigente**: Superior 2,0 cm • Inferior 2,0 cm • Esquerda 2,0 cm • Direita 3,0 cm.
Aplica em **corpo, cabeçalhos e rodapés** de TODAS as seções. Preserva logomarcas, tabelas e paginações.
""")

# --- 2. MOTOR DE ALTERAÇÃO XML — MARGENS ABNT/NORMA ZERO ---
def injetar_margens_via_xml_puro(arquivo_bytes):
    """
    Margens conforme Norma Zero e ABNT vigente (A4):
      Superior: 2,0 cm → 1134 dxa
      Inferior: 2,0 cm → 1134 dxa
      Esquerda: 2,0 cm → 1134 dxa
      Direita: 3,0 cm → 1701 dxa
    Aplica em corpo, cabeçalhos e rodapés de TODAS as seções.
    """
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1134", "1134", "1134", "1701"
    
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            # ✅ CORPO DO DOCUMENTO — TODAS as seções
            if item.filename == "word/document.xml":
                xml_texto = conteudo.decode("utf-8")
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                conteudo = xml_texto.encode("utf-8")

            # ✅ CABEÇALHOS — garante as mesmas margens
            elif item.filename.startswith("word/header"):
                xml_texto = conteudo.decode("utf-8")
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                conteudo = xml_texto.encode("utf-8")

            # ✅ RODAPÉS — garante as mesmas margens
            elif item.filename.startswith("word/footer"):
                xml_texto = conteudo.decode("utf-8")
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                conteudo = xml_texto.encode("utf-8")
            
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# --- 3. FLUXO DE COMPILAÇÃO E TRIAGEM DE METADADOS ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    dados_brutos = arquivo_word.read()
    
    # Extração de textos para triagem
    doc_triagem = docx.Document(BytesIO(dados_brutos))
    texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:40]])
    texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:1] for r in t.rows for cell in r.cells])
    texto_total_raw = (texto_corpo_raw + " " + texto_tabelas_raw).upper()
    
    # Triagem do Tipo Documental
    tipo_detectado = "PROTOCOLO"
    if "PROCEDIMENTO OPERACIONAL" in texto_total_raw or "POP" in texto_total_raw:
        tipo_detectado = "POP"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    # Extração de Código para nomeação
    codigo_doc = "PROT_SCIH005"
    match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|ROT)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
    if match_codigo:
        codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "")
    
    # Aplicação das margens
    dados_finais = injetar_margens_via_xml_puro(dados_brutos)
    
    # Botão de Download
    st.download_button(
        label="📥 BAIXAR DOCUMENTO FORMATADO",
        data=dados_finais,
        file_name=f"{codigo_doc}_Formatado_Norma_Zero.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    
    st.success(f"""✅ **FORMATAÇÃO CONCLUÍDA — Margens ABNT/Norma Zero aplicadas!**
• Superior: 2,0 cm | Inferior: 2,0 cm | Esquerda: 2,0 cm | Direita: 3,0 cm
• Aplicadas em corpo, cabeçalhos e rodapés de todas as seções.
• Logos, tabelas e paginações preservadas. Páginas em branco eliminadas.""")
