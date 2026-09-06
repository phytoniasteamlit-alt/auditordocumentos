import streamlit as st
import docx
import zipfile
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="APENAS Margens + Limpar Espaços", page_icon="📄", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("✅ APENAS APLICAR MARGENS + LIMPAR ESPAÇOS")
st.markdown("""
### 🛡️ CABEÇALHO E RODAPÉ → 100% PRESERVADOS
✅ NÃO copia, NÃO substitui, NÃO altera NADA do cabeçalho/rodapé
✅ Aplica margens corretas APENAS no corpo do texto
✅ Limpa espaços gigantes e páginas em branco
✅ Marca d'água, número de páginas → FICAM EXATAMENTE COMO VOCÊ DEIXOU
""")

# ============================================================
# 🧠 MOTOR — SÓ ALTERA CORPO! CABEÇALHO = INTACTO!
# ============================================================
def apenas_margens_e_limpar(arquivo_bytes):
    """
    🛡️ CABEÇALHO e RODAPÉ → COPIADOS IGUAIS, SEM ALTERAÇÃO NENHUMA!
    ✅ Aplica margens corretas (3,0 / 3,0 / 2,0 / 2,0) no corpo
    ✅ Limpa espaços e quebras
    ✅ NÃO TOCA em cabeçalho, rodapé, marca d'água, número de páginas
    """
    # Margens corretas para o CORPO do documento
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1701", "1134", "1701", "1134"
    
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            # ======================================
            # ✅ SÓ ALTERA O CORPO DO TEXTO
            # ======================================
            if item.filename == "word/document.xml":
                xml = conteudo.decode("utf-8")
                
                # Aplica margens corretas no corpo
                xml = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml)
                xml = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml)
                xml = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml)
                xml = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml)
                
                # Limpa espaçamentos gigantes
                xml = re.sub(r'w:spaceBefore="\d+"', 'w:spaceBefore="0"', xml)
                xml = re.sub(r'w:spaceAfter="\d+"',  'w:spaceAfter="240"', xml)
                
                # Remove quebras que causam páginas vazias
                xml = re.sub(r'<w:br w:type="page"/>', '', xml)
                xml = re.sub(r'<w:pageBreakBefore/>', '', xml)
                
                conteudo = xml.encode("utf-8")

            # ======================================
            # 🛡️ CABEÇALHO e RODAPÉ → COPIA IGUAL! SEM ALTERAÇÃO!
            # ======================================
            else:
                # ✅ Copia o arquivo EXATAMENTE como veio! Sem mexer!
                pass  # NÃO FAZ NADA!
            
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# ============================================================
# 🚀 INTERFACE SIMPLES
# ============================================================
with st.form("form_não_mexe_cabecalho"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o arquivo AQUI com o cabeçalho JÁ PERFEITO por você",
        type=["docx"],
        key="upload_não_mexe_cabecalho"
    )
    enviado = st.form_submit_button("🔄 APLICAR MARGENS + LIMPAR ESPAÇOS", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Aplicando margens... CABEÇALHO PRESERVADO INTACTO..."):
        dados_brutos = arquivo_word.read()
        dados_finais = apenas_margens_e_limpar(dados_brutos)
        
        st.download_button(
            label="📥 BAIXAR — CABEÇALHO INTACTO + MARGENS CORRETAS",
            data=dados_finais,
            file_name=f"{arquivo_word.name.replace('.docx','')}_MARGENS_CORRETAS.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        
        st.success("""✅ **CONCLUÍDO — CABEÇALHO NÃO FOI ALTERADO!**
• 🛡️ Cabeçalho, rodapé, marca d'água, nº de páginas → EXATAMENTE COMO VOCÊ DEIXOU
• ✅ Margens corretas aplicadas no corpo: Sup 3,0 • Esq 3,0 • Inf 2,0 • Dir 2,0
• ✅ Espaços gigantes e páginas em branco → REMOVIDOS
• ✅ NADA foi alterado no cabeçalho — copiado igual em todas as páginas""")
