import streamlit as st
import docx
import zipfile
import re
from io import BytesIO
import time

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="Formatador de Documentos NAQH", page_icon="📊", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Formatador Automático - NAQH")
st.markdown("""
### 🧠 CORRIGIDO — Margens ABNT + Recuo 1,25cm + RÁPIDO
✅ Superior: 3,0cm • Esquerda: 3,0cm • Inferior: 2,0cm • Direita: 2,0cm
✅ Recuo de 1,25cm na 1ª linha • Títulos/listas SEM recuo
""")

# ============================================================
# 📋 SEÇÕES OBRIGATÓRIAS
# ============================================================
SECOES_POR_TIPO = {
    "POI": ["1. INTRODUÇÃO", "2. OBJETIVO", "3. FINALIDADE", "4. ABRANGÊNCIA", "5. RESPONSABILIDADES", "6. GESTÃO DE RISCO", "7. ANEXOS", "8. REFERÊNCIAS"],
    "POP": ["1. DEFINIÇÃO", "2. APLICABILIDADE", "3. RESPONSABILIDADES", "4. DESCRIÇÃO DAS ETAPAS", "5. REFERÊNCIAS", "6. ANEXOS"],
    "PROG": ["1. REFERENCIAL TEÓRICO", "2. OBJETIVOS", "3. METAS E INDICADORES", "4. DEFINIÇÃO DE METAS", "5. ACOMPANHAMENTO E MONITORAMENTO", "6. AVALIAÇÃO DE RESULTADOS", "7. REFERÊNCIAS", "8. ANEXOS"],
    "PROT": ["1. OBJETIVO", "2. APLICABILIDADE", "3. REFERENCIAL TEÓRICO", "4. CLASSIFICAÇÃO", "5. RESPONSABILIDADES", "6. MEDIDAS PREVENTIVAS", "7. REFERÊNCIAS", "8. ANEXOS"],
    "REG": ["1. FINALIDADE", "2. ÂMBITO", "3. COMPETÊNCIA E ORGANIZAÇÃO", "4. DISPOSIÇÕES GERAIS", "5. DISPOSIÇÕES FINAIS"],
    "NOR": ["1. OBJETIVO", "2. ABRANGÊNCIA", "3. DEFINIÇÕES", "4. COMPETÊNCIAS", "5. PROCEDIMENTOS", "6. DISPOSIÇÕES FINAIS", "7. REFERÊNCIAS"]
}

# Termos que NÃO devem ter recuo
NAO_RECUAR = [
    "OBJETIVO", "APLICABILIDADE", "REFERENCIAL TEÓRICO", "CLASSIFICAÇÃO",
    "RESPONSABILIDADES", "MEDIDAS PREVENTIVAS", "REFERÊNCIAS", "ANEXOS",
    "DEFINIÇÃO", "FINALIDADE", "ÂMBITO", "COMPETÊNCIA", "PROCEDIMENTOS",
    "DISPOSIÇÕES", "QUADRO", "Figura", "Tabela", "a)", "b)", "c)", "d)", "e)",
    "5.1", "5.2", "5.3", "5.4", "6.1", "•", "●"
]

# ============================================================
# 🧠 MOTOR SIMPLIFICADO — SEM LOOP, SEM TRAVAMENTO
# ============================================================
def formatar_documento_completo(arquivo_bytes):
    """
    Margens: Sup 3,0cm Esq 3,0cm Inf 2,0cm Dir 2,0cm
    Recuo 1,25cm na 1ª linha — forma SIMPLES e RÁPIDA
    """
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1701", "1134", "1701", "1134"
    recuo = "709"  # 1,25cm
    
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            if item.filename == "word/document.xml":
                xml = conteudo.decode("utf-8")
                
                # ✅ Aplica margens — RÁPIDO
                xml = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml)
                xml = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml)
                xml = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml)
                xml = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml)
                
                # ✅ Zera espaçamentos — ELIMINA BURACOS
                xml = re.sub(r'w:spaceBefore="\d+"', 'w:spaceBefore="0"', xml)
                xml = re.sub(r'w:spaceAfter="\d+"',  'w:spaceAfter="240"', xml)
                
                # ✅ Remove recuos antigos
                xml = re.sub(r'<w:ind[^>]+/?>', '', xml)
                
                # ✅ Aplica recuo de forma SIMPLES: adiciona <w:ind w:firstLine="709"/>
                # Depois de cada <w:pPr> que NÃO seja título/lista
                xml = re.sub(
                    r'(<w:pPr>)((?:(?!<w:numPr>|<w:pStyle w:val="Título").){0,150}</w:pPr>)',
                    rf'\1\2<w:ind w:firstLine="{recuo}"/>',
                    xml
                )
                
                # ✅ Remove recuo de títulos e listas
                for termo in NAO_RECUAR:
                    padrao = rf'<w:ind w:firstLine="{recuo}"/>([^<]*?{re.escape(termo)}[^<]*?)</w:pPr>'
                    xml = re.sub(padrao, r'\1</w:pPr>', xml, flags=re.IGNORECASE)
                
                # ✅ Remove quebras problemáticas
                xml = re.sub(r'<w:br w:type="page"/>', '', xml)
                xml = re.sub(r'<w:pageBreakBefore/>', '', xml)
                
                conteudo = xml.encode("utf-8")

            # ✅ CABEÇALHOS e RODAPÉS — mesmas margens
            elif item.filename.startswith("word/header") or item.filename.startswith("word/footer"):
                xml = conteudo.decode("utf-8")
                xml = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml)
                xml = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml)
                xml = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml)
                xml = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml)
                conteudo = xml.encode("utf-8")
            
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# ============================================================
# 🔍 TRIAGEM
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
# 🚀 INTERFACE — SEM TRAVAMENTO
# ============================================================
with st.form("form_final_rapido"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) aqui",
        type=["docx"],
        key="upload_rapido"
    )
    enviado = st.form_submit_button("🔄 ANALISAR E FORMATAR", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Processando... aplicando margens e recuo..."):
        dados_brutos = arquivo_word.read()
        
        doc_triagem = docx.Document(BytesIO(dados_brutos))
        texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:60]])
        texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:2] for r in t.rows for cell in r.cells])
        texto_total_raw = (texto_corpo_raw + " " + texto_tabelas_raw).upper()
        
        sigla_tipo, secoes_esperadas = identificar_tipo_e_secoes(texto_total_raw)
        
        codigo_doc = f"{sigla_tipo}_SCIH000"
        match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|REG|PROG|POI)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
        if match_codigo:
            codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "")
        
        # ✅ Processamento RÁPIDO
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
        
        st.success("""✅ **CONCLUÍDO!**
• ✅ Margens: Sup 3,0cm | Esq 3,0cm | Inf 2,0cm | Dir 2,0cm
• ✅ Recuo 1,25cm na 1ª linha • Títulos SEM recuo
• ✅ Espaçamentos normalizados • Sem páginas em branco""")
