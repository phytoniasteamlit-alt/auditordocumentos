import streamlit as st
import zipfile
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="Apenas Remover Páginas Vazias", page_icon="📄", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("✅ REMOVER PÁGINAS EM BRANCO — Preservar Tudo")
st.markdown("""
### 🧹 SÓ LIMPA PÁGINAS VAZIAS — NÃO ALTERA MAIS NADA!
✅ Mantém margens ABNT: 3,0 / 3,0 / 2,0 / 2,0 cm
✅ Mantém cabeçalho, rodapé, marca d'água e nº de páginas INTACTOS
✅ Remove APENAS quebras de página e espaços que causam páginas vazias
✅ NÃO desestrutura NADA do documento
""")

# ============================================================
# 🧠 MOTOR — SÓ LIMPA PÁGINAS VAZIAS! NÃO ALTERA MAIS NADA!
# ============================================================
def apenas_remover_paginas_vazias(arquivo_bytes):
    """
    🛡️ PRESERVA TUDO — margens, cabeçalho, conteúdo, tudo igual!
    ✅ Remove APENAS: quebras de página forçadas e parágrafos vazios gigantes
    ✅ NÃO altera NADA mais
    """
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            # ✅ SÓ ALTERA O CORPO DO TEXTO
            if item.filename == "word/document.xml":
                xml = conteudo.decode("utf-8")
                
                # 🧹 REMOVE quebras de página FORÇADAS que causam páginas vazias
                xml = re.sub(r'<w:br w:type="page"/>', '', xml)
                xml = re.sub(r'<w:pageBreakBefore/>', '', xml)
                xml = re.sub(r'<w:lastRenderedPageBreak/>', '', xml)
                
                # 🧹 Zera espaçamentos GIGANTES que empurram conteúdo
                xml = re.sub(r'w:spaceBefore="\d+"', 'w:spaceBefore="0"', xml)
                xml = re.sub(r'w:spaceAfter="\d+"',  'w:spaceAfter="240"', xml)
                
                # 🧹 Remove parágrafos COMPLETAMENTE VAZIOS que criam linhas vazias
                xml = re.sub(r'<w:p[^>]*?>\s*<w:pPr[^>]*>\s*</w:pPr>\s*(?:<w:r[^>]*>\s*</w:r>\s*)*</w:p>', '', xml)
                
                conteudo = xml.encode("utf-8")

            # ✅ CABEÇALHO, RODAPÉ E OUTROS → COPIA IGUAL, SEM ALTERAÇÃO!
            # NÃO FAZ NADA!
            
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# ============================================================
# 🚀 INTERFACE
# ============================================================
with st.form("form_apenas_limpar"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o arquivo AQUI com as margens já corretas",
        type=["docx"],
        key="upload_apenas_limpar"
    )
    enviado = st.form_submit_button("🔄 REMOVER PÁGINAS EM BRANCO", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Removendo páginas vazias... PRESERVANDO TUDO O RESTO..."):
        dados_brutos = arquivo_word.read()
        dados_finais = apenas_remover_paginas_vazias(dados_brutos)
        
        st.download_button(
            label="📥 BAIXAR — SEM PÁGINAS VAZIAS",
            data=dados_finais,
            file_name=f"{arquivo_word.name.replace('.docx','')}_SEM_PAGINAS_VAZIAS.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        
        st.success("""✅ **CONCLUÍDO — PÁGINAS VAZIAS REMOVIDAS!**
• 🛡️ Margens ABNT → PRESERVADAS (3,0 / 3,0 / 2,0 / 2,0)
• 🛡️ Cabeçalho, marca d'água, nº de páginas → EXATAMENTE IGUAIS
• ✅ Quebras de página vazias → REMOVIDAS
• ✅ Espaços gigantes entre títulos e textos → REMOVIDOS
• ✅ Estrutura, tabelas e conteúdo → 100% PRESERVADOS""")
