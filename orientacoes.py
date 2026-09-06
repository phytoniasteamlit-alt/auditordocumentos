import streamlit as st
import zipfile
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="Remover Páginas Vazias e Espaços", page_icon="📄", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("✅ REMOVER BURACOS GIGANTES E PÁGINAS VAZIAS")
st.markdown("""
### 🧹 ELIMINA ESPAÇOS GIGANTES QUE EMPURRAM O CONTEÚDO
✅ Remove quebras de página forçadas e espaçamentos enormes
✅ Mantém margens ABNT e cabeçalho 100% intactos
✅ Puxa o conteúdo para logo abaixo do cabeçalho — sem buraco!
""")

# ============================================================
# 🧠 MOTOR — ELIMINA ESPAÇOS GIGANTES + PRESERVA TUDO
# ============================================================
def remover_espacos_e_quebras(arquivo_bytes):
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            if item.filename == "word/document.xml":
                xml = conteudo.decode("utf-8")
                
                # 🧹 ZERA TODO espaçamento ANTES — ELIMINA BURACOS GIGANTES
                xml = re.sub(r'w:spaceBefore="\d+"', 'w:spaceBefore="0"', xml)
                
                # 🧹 Padroniza espaçamento DEPOIS → espaços normais
                xml = re.sub(r'w:spaceAfter="\d+"', 'w:spaceAfter="240"', xml)
                
                # 🧹 REMOVE TODAS as quebras de página FORÇADAS
                xml = re.sub(r'<w:br w:type="page"/>', '', xml)
                xml = re.sub(r'<w:pageBreakBefore/>', '', xml)
                xml = re.sub(r'<w:lastRenderedPageBreak/>', '', xml)
                xml = re.sub(r'<w:bookmarkStart[^>]*>', '', xml)
                xml = re.sub(r'<w:bookmarkEnd[^>]*>', '', xml)
                
                # 🧹 REMOVE parágrafos VAZIOS que ocupam espaço
                xml = re.sub(r'<w:p[^>]*>\s*<w:pPr[^>]*>\s*</w:pPr>\s*</w:p>', '', xml)
                
                # 🧹 REMOVE valores de espaçamento ABSURDOS (acima de 1000 = buraco gigante)
                xml = re.sub(r'w:spaceBefore="(1\d{3}|[2-9]\d{3})"', 'w:spaceBefore="0"', xml)
                xml = re.sub(r'w:spaceAfter="(1\d{3}|[2-9]\d{3})"', 'w:spaceAfter="240"', xml)
                
                conteudo = xml.encode("utf-8")

            # ✅ CABEÇALHO E RODAPÉ → COPIA IGUAL, SEM ALTERAÇÃO!
            
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# ============================================================
# 🚀 INTERFACE
# ============================================================
with st.form("form_final_buracos"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o arquivo AQUI com margens corretas",
        type=["docx"],
        key="upload_buracos"
    )
    enviado = st.form_submit_button("🔄 ELIMINAR BURACOS E PÁGINAS VAZIAS", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Removendo espaços gigantes que empurram o conteúdo..."):
        dados_brutos = arquivo_word.read()
        dados_finais = remover_espacos_e_quebras(dados_brutos)
        
        st.download_button(
            label="📥 BAIXAR — SEM BURACOS E SEM PÁGINAS VAZIAS",
            data=dados_finais,
            file_name=f"{arquivo_word.name.replace('.docx','')}_SEM_BURACOS.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        
        st.success("""✅ **CONCLUÍDO! BURACOS GIGANTES ELIMINADOS!**
• ✅ Conteúdo puxado para logo abaixo do cabeçalho
• ✅ Espaçamentos gigantes → ZERADOS
• ✅ Quebras de página forçadas → REMOVIDAS
• 🛡️ Cabeçalho, margens e rodapé → 100% PRESERVADOS""")
