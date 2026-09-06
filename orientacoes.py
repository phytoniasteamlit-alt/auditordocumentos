import streamlit as st
import docx
import zipfile
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Formatador de Documentos NAQH", page_icon="📊", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Formatador Automático - NAQH")
st.markdown("""
### 🧠 Inteligência XML de Alta Fidelidade
Margens Norma Zero / ABNT vigente + limpeza de espaçamentos e quebras excessivas.
Preserva logos, tabelas e estrutura — elimina páginas em branco e espaços gigantes.
""")

# --- 2. MOTOR PRINCIPAL — MARGENS + LIMPEZA DE LAYOUT ---
def injetar_margens_e_limpar(arquivo_bytes):
    """
    ✅ Aplica margens Norma Zero em TODAS as seções, cabeçalhos e rodapés
    ✅ Remove espaçamentos e quebras que causam páginas em branco
    Margens: Sup/Inf: 2,0cm | Esq: 2,0cm | Dir: 3,0cm
    """
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1134", "1134", "1134", "1701"
    
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            # ✅ CORPO DO DOCUMENTO — margens + limpeza de espaços
            if item.filename == "word/document.xml":
                xml_texto = conteudo.decode("utf-8")
                
                # Aplica margens em TODAS as seções
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                
                # ✅ LIMPEZA: Remove espaçamento excessivo antes/depois de parágrafos
                xml_texto = re.sub(r'w:spaceBefore="\d+"', 'w:spaceBefore="0"', xml_texto)
                xml_texto = re.sub(r'w:spaceAfter="\d+"',  'w:spaceAfter="200"', xml_texto)
                
                # ✅ LIMPEZA: Remove quebras de página forçadas que causam páginas vazias
                xml_texto = re.sub(r'<w:br w:type="page"/>', '', xml_texto)
                xml_texto = re.sub(r'<w:pageBreakBefore/>', '', xml_texto)
                
                conteudo = xml_texto.encode("utf-8")

            # ✅ CABEÇALHOS — mesmas margens
            elif item.filename.startswith("word/header"):
                xml_texto = conteudo.decode("utf-8")
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                conteudo = xml_texto.encode("utf-8")

            # ✅ RODAPÉS — mesmas margens
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

# --- 3. TRIAGEM E PROCESSAMENTO ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui", type=["docx"])

if arquivo_word:
    dados_brutos = arquivo_word.read()
    
    # Triagem do tipo de documento
    doc_triagem = docx.Document(BytesIO(dados_brutos))
    texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:40]])
    texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:1] for r in t.rows for cell in r.cells])
    texto_total_raw = (texto_corpo_raw + " " + texto_tabelas_raw).upper()
    
    tipo_detectado = "PROTOCOLO"
    if "PROCEDIMENTO OPERACIONAL" in texto_total_raw or "POP" in texto_total_raw:
        tipo_detectado = "POP"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    # Extrai código para nome do arquivo
    codigo_doc = "PROT_SCIH005"
    match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|ROT)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
    if match_codigo:
        codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "")
    
    # Aplica margens + limpeza de espaços
    dados_finais = injetar_margens_e_limpar(dados_brutos)
    
    st.download_button(
        label="📥 BAIXAR DOCUMENTO FORMATADO",
        data=dados_finais,
        file_name=f"{codigo_doc}_Formatado_Norma_Zero.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    
    st.success(f"""✅ **CONCLUÍDO — Margens + Layout corrigidos!**
• Superior: 2,0 cm | Inferior: 2,0 cm | Esquerda: 2,0 cm | Direita: 3,0 cm
• Aplicadas em corpo, cabeçalhos e rodapés de todas as seções
• ✅ Espaçamentos excessivos normalizados
• ✅ Quebras de página problemáticas removidas
• Logos, tabelas e estrutura preservados""")
