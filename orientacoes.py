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
### 🧠 Formatação Completa — Norma Zero + Cabeçalho Preservado
Margens unificadas + cabeçalho redimensionado + tabelas ajustadas + sem páginas em branco.
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
# 🧠 MOTOR DE FORMATAÇÃO COMPLETA — CABEÇALHO INCLUSO
# ============================================================
def formatar_documento_completo(arquivo_bytes):
    """
    ✅ Margens Norma Zero em corpo, cabeçalhos e rodapés
    ✅ Redimensiona TABELAS do cabeçalho para caber na margem nova
    ✅ Normaliza espaçamentos e remove quebras problemáticas
    ✅ Preserva número de páginas, versão e validade
    """
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1134", "1134", "1134", "1701"
    REDUCAO_LARGURA = 500  # ~1cm em dxa — reduz tabelas para caber
    
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
                
                # Aplica margens em TODAS as seções
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                
                # Normaliza espaçamentos
                xml_texto = re.sub(r'w:spaceBefore="\d+"', 'w:spaceBefore="0"', xml_texto)
                xml_texto = re.sub(r'w:spaceAfter="\d+"',  'w:spaceAfter="240"', xml_texto)
                xml_texto = re.sub(r'w:lineSpacing="\d+"', 'w:lineSpacing="276"', xml_texto)
                
                # Remove quebras problemáticas
                xml_texto = re.sub(r'<w:br w:type="page"/>', '', xml_texto)
                xml_texto = re.sub(r'<w:pageBreakBefore/>', '', xml_texto)
                
                # ✅ REDUZ LARGURA DE TODAS AS TABELAS em ~1cm
                xml_texto = re.sub(
                    r'(<w:tblW w:w=")(\d+)(" w:type="dxa"/>)',
                    lambda m: f'{m.group(1)}{max(2000, int(m.group(2)) - REDUCAO_LARGURA)}{m.group(3)}',
                    xml_texto
                )
                
                conteudo = xml_texto.encode("utf-8")

            # ======================================
            # ✅ CABEÇALHO — MESMAS MARGENS + TABELA REDUZIDA
            # ======================================
            elif item.filename.startswith("word/header"):
                xml_texto = conteudo.decode("utf-8")
                
                # Aplica margens IGUAIS ao corpo
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                
                # ✅ REDUZ TABELA DO CABEÇALHO em ~1cm → NÃO DESALINHA MAIS!
                xml_texto = re.sub(
                    r'(<w:tblW w:w=")(\d+)(" w:type="dxa"/>)',
                    lambda m: f'{m.group(1)}{max(1500, int(m.group(2)) - REDUCAO_LARGURA)}{m.group(3)}',
                    xml_texto
                )
                
                # ✅ REDUZ LARGURA DAS CÉLULAS da tabela do cabeçalho
                xml_texto = re.sub(
                    r'(<w:tcW w:w=")(\d+)(" w:type="dxa"/>)',
                    lambda m: f'{m.group(1)}{max(500, int(m.group(2)) - 170)}{m.group(3)}',
                    xml_texto
                )
                
                conteudo = xml_texto.encode("utf-8")

            # ======================================
            # RODAPÉS — MESMAS MARGENS
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
# 🚀 INTERFACE COM FORMULÁRIO
# ============================================================
with st.form("form_cabecalho_corrigido"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) aqui",
        type=["docx"],
        key="upload_cabecalho"
    )
    enviado = st.form_submit_button("🔄 ANALISAR E FORMATAR", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Ajustando margens, redimensionando cabeçalho e tabelas..."):
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
        
        # ✅ Processamento COMPLETO com CABEÇALHO AJUSTADO
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
        
        st.success("""✅ **PROCESSO CONCLUÍDO — CABEÇALHO PRESERVADO!**
• ✅ Margens Norma Zero em corpo, cabeçalhos e rodapés
• ✅ Tabela do cabeçalho REDIMENSIONADA — não desalinha mais!
• ✅ Número de páginas, versão e validade PRESERVADOS
• ✅ Tabelas internas ajustadas à nova largura
• ✅ Espaçamentos normalizados e páginas em branco eliminadas""")
