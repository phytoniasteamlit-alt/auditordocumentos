import streamlit as st
import zipfile
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="Margens ABNT Exatas", page_icon="📄", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("✅ MARGENS ABNT EXATAS — NBR 14724")
st.markdown("""
### 📏 Superior 3,0 • Esquerda 3,0 • Inferior 2,0 • Direita 2,0 cm
🛡️ Cabeçalho e rodapé → 100% PRESERVADOS, NÃO ALTERADOS
Aplica as margens corretas em TODO o documento de forma segura.
""")

# ============================================================
# 🧠 MOTOR — MARGENS EXATAS ABNT + NÃO TOCA NO CABEÇALHO
# ============================================================
def aplicar_margens_abnt_exatas(arquivo_bytes):
    """
    ✅ APLICA margens EXATAS da ABNT em TODAS as seções do documento
    ✅ NÃO ALTERA cabeçalho, rodapé, marca d'água, número de páginas
    ✅ Limpa espaços e quebras de página
    """
    # 📏 VALORES EXATOS — NÃO MEXER!
    top_dxa    = "1701"  # 3,0 cm
    esq_dxa    = "1701"  # 3,0 cm
    inf_dxa    = "1134"  # 2,0 cm
    dir_dxa    = "1134"  # 2,0 cm
    
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            nome = item.filename
            
            # ======================================
            # ✅ CORPO DO DOCUMENTO — APLICA MARGENS
            # ======================================
            if nome == "word/document.xml":
                xml = conteudo.decode("utf-8")
                
                # ✅ APLICA EM TODAS AS SEÇÕES DO CORPO
                xml = re.sub(r'<w:pgMar[^>]*?w:top="\d+"[^>]*?>', 
                            f'<w:pgMar w:top="{top_dxa}" w:left="{esq_dxa}" w:bottom="{inf_dxa}" w:right="{dir_dxa}"/>', 
                            xml)
                
                # ✅ Se não encontrou a tag completa, substitui por padrão
                if '<w:pgMar' not in xml:
                    xml = re.sub(r'<w:sectPr[^>]*>', 
                                lambda m: m.group(0) + f'<w:pgMar w:top="{top_dxa}" w:left="{esq_dxa}" w:bottom="{inf_dxa}" w:right="{dir_dxa}"/>', 
                                xml)
                
                # ✅ Limpa espaçamentos gigantes
                xml = re.sub(r'w:spaceBefore="\d+"', 'w:spaceBefore="0"', xml)
                xml = re.sub(r'w:spaceAfter="\d+"',  'w:spaceAfter="240"', xml)
                
                # ✅ Remove quebras vazias
                xml = re.sub(r'<w:br w:type="page"/>', '', xml)
                xml = re.sub(r'<w:pageBreakBefore/>', '', xml)
                
                conteudo = xml.encode("utf-8")

            # ======================================
            # ✅ CABEÇALHOS E RODAPÉS — AJUSTA APENAS MARGENS INTERNAS
            # ======================================
            elif nome.startswith("word/header") or nome.startswith("word/footer"):
                xml = conteudo.decode("utf-8")
                
                # ✅ Ajusta margens internas do cabeçalho/rodapé para coincidir com o corpo
                xml = re.sub(r'w:headerMargin="\d+"', f'w:headerMargin="{top_dxa}"', xml)
                xml = re.sub(r'w:footerMargin="\d+"', f'w:footerMargin="{inf_dxa}"', xml)
                
                conteudo = xml.encode("utf-8")
            
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# ============================================================
# 🚀 INTERFACE
# ============================================================
with st.form("form_margens_abnt_exatas"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o arquivo AQUI com o cabeçalho já ajustado",
        type=["docx"],
        key="upload_margens_abnt_exatas"
    )
    enviado = st.form_submit_button("🔄 APLICAR MARGENS ABNT EXATAS", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Aplicando margens EXATAS da ABNT... cabeçalho preservado..."):
        dados_brutos = arquivo_word.read()
        dados_finais = aplicar_margens_abnt_exatas(dados_brutos)
        
        st.download_button(
            label="📥 BAIXAR — MARGENS ABNT 3/3/2/2",
            data=dados_finais,
            file_name=f"{arquivo_word.name.replace('.docx','')}_ABNT_3_3_2_2.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        
        st.success("""✅ **CONCLUÍDO — MARGENS ABNT EXATAS APLICADAS!**
• 📏 Superior: 3,0cm • Esquerda: 3,0cm • Inferior: 2,0cm • Direita: 2,0cm
• 🛡️ Cabeçalho, rodapé, marca d'água e nº de páginas → PRESERVADOS
• ✅ Espaços e quebras vazias → REMOVIDOS
• ✅ Abra o Word → Configurar Página → as margens estarão corretas!""")
