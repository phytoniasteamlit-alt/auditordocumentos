import streamlit as st
import docx
import zipfile
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="Limpador de Espaços - NAQH", page_icon="📄", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("🧹 LIMPAR PÁGINAS EM BRANCO E ESPAÇOS")
st.markdown("""
### ✅ NÃO ALTERA SEU CABEÇALHO!
Apenas remove espaços vazios gigantes e quebras de página problemáticas.
Seu cabeçalho e rodapé ficam EXATAMENTE como você deixou.
""")

# ============================================================
# 🧠 LIMPA ESPAÇOS E QUEBRAS — NÃO TOCA NO CABEÇALHO!
# ============================================================
def limpar_espacos_e_paginas_vazias(arquivo_bytes):
    """
    ✅ NÃO ALTERA cabeçalho e rodapé — PRESERVA 100%
    ✅ Remove espaçamentos gigantes entre títulos e textos
    ✅ Remove quebras de página que causam páginas vazias
    ✅ Normaliza espaçamento dos parágrafos
    """
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            # ✅ SÓ ALTERA O CORPO DO TEXTO
            if item.filename == "word/document.xml":
                xml = conteudo.decode("utf-8")
                
                # ✅ ZERA espaçamento ANTES dos parágrafos → ELIMINA BURACOS GIGANTES
                xml = re.sub(r'w:spaceBefore="\d+"', 'w:spaceBefore="0"', xml)
                
                # ✅ Padroniza espaçamento DEPOIS → espaços uniformes
                xml = re.sub(r'w:spaceAfter="\d+"', 'w:spaceAfter="240"', xml)
                
                # ✅ Remove quebras de página FORÇADAS que causam páginas vazias
                xml = re.sub(r'<w:br w:type="page"/>', '', xml)
                xml = re.sub(r'<w:pageBreakBefore/>', '', xml)
                
                # ✅ Remove parágrafos TOTALMENTE VAZIOS que criam linhas em branco
                xml = re.sub(r'<w:p[^>]*>\s*<w:pPr[^>]*>\s*</w:pPr>\s*</w:p>', '', xml)
                xml = re.sub(r'<w:p[^/>]*/>', '', xml)
                
                conteudo = xml.encode("utf-8")

            # ✅ CABEÇALHOS e RODAPÉS → COPIA IGUAL, NÃO ALTERA NADA!
            # Não aplica nenhuma mudança — preserva o que você ajustou!
            
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# ============================================================
# 🚀 INTERFACE SIMPLES
# ============================================================
with st.form("form_apenas_limpar"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o arquivo AQUI com o cabeçalho JÁ AJUSTADO por você",
        type=["docx"],
        key="upload_apenas_limpar"
    )
    enviado = st.form_submit_button("🔄 LIMPAR ESPAÇOS E PÁGINAS VAZIAS", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Apenas limpando espaços vazios... cabeçalho preservado..."):
        dados_brutos = arquivo_word.read()
        dados_finais = limpar_espacos_e_paginas_vazias(dados_brutos)
        
        st.download_button(
            label="📥 BAIXAR — SEM PÁGINAS VAZIAS",
            data=dados_finais,
            file_name=f"{arquivo_word.name.replace('.docx','')}_Limpo_Sem_Vazios.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        
        st.success("""✅ **CONCLUÍDO! CABEÇALHO PRESERVADO!**
• ✅ Seu cabeçalho e rodapé → NÃO FORAM ALTERADOS
• ✅ Espaços gigantes entre títulos e textos → REMOVIDOS
• ✅ Páginas em branco vazias → ELIMINADAS
• ✅ Espaçamentos normalizados → documento limpo e organizado""")
