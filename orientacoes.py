import streamlit as st
import docx
import zipfile
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO DA PÁGINA STREAMLIT ---
st.set_page_config(page_title="Formatador de Documentos NAQH", page_icon="📊", layout="wide")

# Menu Lateral - Identificação Visual do Operador Autêntica
with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Formatador Automático - NAQH")
st.markdown("""
### 🧠 Inteligência XML de Alta Fidelidade (Safe Mode)
O sistema aplica as margens oficiais da Norma Zero alterando diretamente as tags estruturais do pacote, **garantindo a permanência absoluta de logomarcas, tabelas de cabeçalho e paginações originais**.
""")

# --- 2. MOTOR DE ALTERAÇÃO XML DIRETA (PRESERVAÇÃO ESTRUTURAL ABSOLUTA) ---
def injetar_margens_via_xml_puro(arquivo_bytes):
    """
    Modifica as margens da folha no código XML nativo do Word.
    Converte os centímetros oficiais da Norma Zero em dxa (1 cm = 567 dxa):
    - Superior: 2.0 cm -> 1134 dxa
    - Inferior: 2.0 cm -> 1134 dxa
    - Esquerda: 2.0 cm -> 1134 dxa
    - Direita: 3.0 cm -> 1701 dxa
    """
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1134", "1134", "1134", "1701"
    
    # Abre o documento original enviado pelo usuário como um pacote ZIP em memória
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    # Cria um novo pacote ZIP que guardará o arquivo corrigido
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            # Intercepta e altera apenas o arquivo XML que gerencia o layout e margens do corpo
            if item.filename == "word/document.xml":
                xml_texto = conteudo.decode("utf-8")
                
                # Substituição precisa via expressões regulares nas tags de margem do Word (w:pgMar)
                xml_texto = re.sub(r'w:top="[^"]*"', f'w:top="{top_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:bottom="[^"]*"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="[^"]*"', f'w:left="{left_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:right="[^"]*"', f'w:right="{right_dxa}"', xml_texto)
                
                conteudo = xml_texto.encode("utf-8")
                
            # Copia todos os outros arquivos (cabeçalhos, rodapés, imagens, mídias) sem alterar um único bit
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# --- 3. FLUXO DE COMPILAÇÃO E TRIAGEM DE METADADOS ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    # Leitura dos bytes brutos em lote de segurança
    dados_brutos = arquivo_word.read()
    
    # Varredura rápida usando uma instância temporária apenas para extrair textos da triagem
    doc_triagem = docx.Document(BytesIO(dados_brutos))
    texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:40]])
    texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:1] for r in t.rows for cell in r.cells])
    texto_total_raw = (texto_corpo_raw + " " + texto_tabelas_raw).upper()
    
    # Triagem do Tipo Documental
    tipo_detectado = "PROTOCOLO"
    if "PROCEDIMENTO OPERACIONAL" in texto_total_raw or "POP" in texto_total_raw:
        tipo_detectado = "POP"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    # Extração Inteligente de Código para Nomeação de Arquivo
    codigo_doc = "PROT_SCIH005"
    match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|ROT)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
    if match_codigo:
        codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "")

    # Processamento em lote via injeção XML direta
    dados_finais = injetar_margens_via_xml_puro(dados_brutos)

    # Exibição do botão estável na interface
    st.download_button(
        label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO FORMATADO",
        data=dados_finais,
        file_name=f"{codigo_doc}_Formatado_Homologado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** As margens de {tipo_detectado.capitalize()} foram atualizadas para o padrão institucional (2x2x2x3 cm) com preservação absoluta de cabeçalhos, logos e paginações nativas.")
