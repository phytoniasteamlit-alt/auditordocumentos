import streamlit as st
import docx
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
### 🧠 Inteligência XML de Alta Fidelidade
O sistema aplica as margens oficiais da Norma Zero alterando diretamente o código estrutural do arquivo, **garantindo a permanência absoluta de logomarcas, cabeçalhos e paginações**.
""")

# --- 2. MOTOR DE ALTERAÇÃO XML DIRETA (PRESERVAÇÃO ABSOLUTA DE IMAGENS E CABEÇALHOS) ---
def forcar_margens_via_xml(arquivo_bytes):
    # Converte os centímetros oficiais em dxa (unidade padrão do XML do Word: 1 cm = 567 dxa)
    # Norma Zero: Sup 2.0cm (1134), Inf 2.0cm (1134), Esq 2.0cm (1134), Dir 3.0cm (1701)
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1134", "1134", "1134", "1701"
    
    # Abre o arquivo original como um pacote ZIP em memória
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    zip_novo = zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED)
    
    for item in zip_original.infolist():
        conteudo = zip_original.read(item.filename)
        
        # Altera apenas o arquivo principal que controla a geometria e layout da folha
        if item.filename == "word/document.xml":
            xml_texto = conteudo.decode("utf-8")
            
            # Substitui as margens via expressões regulares no XML estrutural do arquivo
            xml_texto = re.sub(r'w:top="[^"]*"', f'w:top="{top_dxa}"', xml_texto)
            xml_texto = re.sub(r'w:bottom="[^"]*"', f'w:bottom="{bottom_dxa}"', xml_texto)
            xml_texto = re.sub(r'w:left="[^"]*"', f'w:left="{left_dxa}"', xml_texto)
            xml_texto = re.sub(r'w:right="[^"]*"', f'w:right="{right_dxa}"', xml_texto)
            
            conteudo = xml_texto.encode("utf-8")
            
        zip_novo.writestr(item, conteudo)
        
    zip_original.close()
    zip_novo.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# --- 3. FLUXO DE PROCESSAMENTO ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    # Lemos os bytes brutos uma única vez
    dados_brutos = arquivo_word.read()
    
    # Varredura superficial e rápida apenas para coletar metadados de identificação
    doc_triagem = docx.Document(BytesIO(dados_brutos))
    texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:30]])
    texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:1] for r in t.rows for cell in r.cells])
    texto_total_raw = (texto_corpo_raw + " " + texto_tabelas_raw).upper()
    
    tipo_detectado = "PROTOCOLO"
    if "PROCEDIMENTO OPERACIONAL" in texto_total_raw or "POP" in texto_total_raw:
        tipo_detectado = "POP"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    codigo_doc = "PROT_SCIH005"
    match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|ROT)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
    if match_codigo:
        codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "")

    # Executa a alteração blindada de margens via XML
    dados_finais = forcar_margens_via_xml(dados_brutos)

    st.download_button(
        label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO FORMATADO",
        data=dados_finais,
        file_name=f"{codigo_doc}_Formatado_Homologado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Margens injetadas diretamente no XML da página. Cabeçalhos, logomarcas e paginação originais preservados com 100% de integridade.")
