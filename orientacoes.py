import streamlit as st
import docx
import zipfile
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="Formatador de Documentos NAQH", page_icon="📊", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Formatador Automático - NAQH")
st.markdown("""
### 🧠 Correção Completa de Layout — Norma Zero + Ajuste de Conteúdo
Margens unificadas em corpo/cabeçalhos/rodapés + normalização de tabelas e espaçamentos.
Elimina páginas em branco e desalinhamentos.
""")

# ============================================================
# 📋 SEÇÕES OBRIGATÓRIAS POR TIPO DE DOCUMENTO
# ============================================================
SECOES_POR_TIPO = {
    "POI": ["1. INTRODUÇÃO", "2. OBJETIVO", "3. FINALIDADE", "4. ABRANGÊNCIA", "5. RESPONSABILIDADES", "6. GESTÃO DE RISCO", "7. ANEXOS", "8. REFERÊNCIAS"],
    "POP": ["1. DEFINIÇÃO", "2. APLICABILIDADE", "3. RESPONSABILIDADES", "4. DESCRIÇÃO DAS ETAPAS", "5. REFERÊNCIAS", "6. ANEXOS"],
    "PROG": ["1. REFERENCIAL TEÓRICO", "2. OBJETIVOS", "3. METAS E INDICADORES", "4. DEFINIÇÃO DE METAS", "5. ACOMPANHAMENTO E MONITORAMENTO", "6. AVALIAÇÃO DE RESULTADOS", "7. REFERÊNCIAS", "8. ANEXOS"],
    "PROT": ["1. OBJETIVO", "2. APLICABILIDADE", "3. REFERENCIAL TEÓRICO", "4. CLASSIFICAÇÃO", "5. RESPONSABILIDADES", "6. MEDIDAS PREVENTIVAS", "7. REFERÊNCIAS", "8. ANEXOS"],
    "REG": ["1. FINALIDADE", "2. ÂMBITO", "3. COMPETÊNCIA E ORGANIZAÇÃO", "4. DISPOSIÇÕES GERAIS", "5. DISPOSIÇÕES FINAIS"],
    "NOR": ["1. OBJETIVO", "2. ABRANGÊNCIA", "3. DEFINIÇÕES", "4. COMPETÊNCIAS", "5. PROCEDIMENTOS", "6. DISPOSIÇÕES FINAIS", "7. REFERÊNCIAS"]
}

# ============================================================
# 🧠 MOTOR DE FORMATAÇÃO COMPLETO — MARGENS + LAYOUT
# ============================================================
def formatar_documento_completo(arquivo_bytes):
    """
    ✅ Aplica margens Norma Zero em corpo, cabeçalhos e rodapés (TODAS IGUAIS)
    ✅ Remove espaçamentos fixos que empurram conteúdo
    ✅ Remove quebras de página problemáticas
    ✅ Reduz largura de tabelas para caber na nova margem direita
    """
    # Margens Norma Zero: Sup/Inf 2cm | Esq 2cm | Dir 3cm
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1134", "1134", "1134", "1701"
    
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            # ======================================
            # CORPO DO DOCUMENTO
            # ======================================
            if item.filename == "word/document.xml":
                xml_texto = conteudo.decode("utf-8")
                
                # ✅ Aplica margens em TODAS as seções
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                
                # ✅ NORMALIZA ESPAÇAMENTO ANTES/DEPOIS dos parágrafos (elimina buracos gigantes)
                xml_texto = re.sub(r'w:spaceBefore="\d+"', 'w:spaceBefore="0"', xml_texto)
                xml_texto = re.sub(r'w:spaceAfter="\d+"',  'w:spaceAfter="240"', xml_texto)
                xml_texto = re.sub(r'w:lineSpacing="\d+"', 'w:lineSpacing="276"', xml_texto)  # 1,15 linha
                
                # ✅ REMOVE QUEBRAS DE PÁGINA FORÇADAS que causam páginas vazias
                xml_texto = re.sub(r'<w:br w:type="page"/>', '', xml_texto)
                xml_texto = re.sub(r'<w:pageBreakBefore/>', '', xml_texto)
                
                # ✅ AJUSTA LARGURA DAS TABELAS — reduz 1cm para caber na margem nova
                xml_texto = re.sub(r'w:w="(\d+)"', lambda m: f'w:w="{str(int(m.group(1)) - 500)}"' if 3000 < int(m.group(1)) < 15000 else m.group(0), xml_texto)
                
                conteudo = xml_texto.encode("utf-8")

            # ======================================
            # CABEÇALHOS — MESMAS MARGENS DO CORPO
            # ======================================
            elif item.filename.startswith("word/header"):
                xml_texto = conteudo.decode("utf-8")
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                conteudo = xml_texto.encode("utf-8")

            # ======================================
            # RODAPÉS — MESMAS MARGENS DO CORPO
            # ======================================
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

# ============================================================
# 🔍 TRIAGEM INTELIGENTE
# ============================================================
def identificar_tipo_e_secoes(texto):
    texto = texto.upper()
    if "PROT_" in texto or "PROTOCOLO" in texto:
        return "PROT", SECOES_POR_TIPO["PROT"]
    elif "POP_" in texto or "PROCEDIMENTO OPERACIONAL" in texto:
        return "POP", SECOES_POR_TIPO["POP"]
    elif "NOR_" in texto or "NORMA" in texto:
        return "NOR", SECOES_POR_TIPO["NOR"]
    elif "REG_" in texto or "REGULAMENTO" in texto:
        return "REG", SECOES_POR_TIPO["REG"]
    elif "PROG_" in texto or "PROGRAMA" in texto:
        return "PROG", SECOES_POR_TIPO["PROG"]
    elif "POI_" in texto or "POLÍTICA" in texto or "POLITICA" in texto:
        return "POI", SECOES_POR_TIPO["POI"]
    else:
        return "PROT", SECOES_POR_TIPO["PROT"]

# ============================================================
# 🚀 INTERFACE COM FORMULÁRIO (sem erro)
# ============================================================
with st.form("form_norma_zero_completo"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) aqui",
        type=["docx"],
        key="upload_completo"
    )
    enviado = st.form_submit_button("🔄 ANALISAR E FORMATAR", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Aplicando margens, unificando cabeçalhos e ajustando layout..."):
        dados_brutos = arquivo_word.read()
        
        # Triagem
        doc_triagem = docx.Document(BytesIO(dados_brutos))
        texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:60]])
        texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:2] for r in t.rows for cell in r.cells])
        texto_total_raw = (texto_corpo_raw + " " + texto_tabelas_raw).upper()
        
        sigla_tipo, secoes_esperadas = identificar_tipo_e_secoes(texto_total_raw)
        
        codigo_doc = f"{sigla_tipo}_SCIH000"
        match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|REG|PROG|POI)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
        if match_codigo:
            codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "")
        
        # ✅ Processamento COMPLETO: margens + layout + tabelas
        dados_finais = formatar_documento_completo(dados_brutos)
        
        st.success(f"📋 **DOCUMENTO IDENTIFICADO: {sigla_tipo}**")
        
        st.markdown("### 📑 Estrutura / Seções Obrigatórias:")
        for secao in secoes_esperadas:
            st.write(f"✅ {secao}")
        
        st.markdown("---")
        
        st.download_button(
            label="📥 BAIXAR DOCUMENTO FORMATADO",
            data=dados_finais,
            file_name=f"{codigo_doc}_Formatado_Norma_Zero.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
        
        st.success("""✅ **PROCESSO CONCLUÍDO — LAYOUT UNIFICADO!**
• ✅ Margens Norma Zero em corpo, cabeçalhos e rodapés (TODAS IGUAIS)
• ✅ Espaçamentos normalizados — elimina buracos gigantes
• ✅ Quebras de página problemáticas removidas
• ✅ Largura de tabelas ajustada à margem nova
• ✅ Cabeçalhos alinhados com o corpo
• Logos e tabelas preservados""")
