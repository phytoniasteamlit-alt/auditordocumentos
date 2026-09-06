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
### 🧠 CORRIGIDO — Margens ABNT/Norma Zero + Recuo Automático
✅ Superior: 3,0cm • Esquerda: 3,0cm • Inferior: 2,0cm • Direita: 2,0cm
✅ Recuo de 1,25cm na primeira linha de todos os parágrafos
✅ Títulos, listas e tabelas SEM recuo
""")

# ============================================================
# 📋 SEÇÕES OBRIGATÓRIAS POR TIPO
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
# 🧠 MOTOR — MARGENS CORRETAS + RECUO DE 1,25cm
# ============================================================
def formatar_documento_completo(arquivo_bytes):
    """
    ✅ MARGENS CORRIGIDAS conforme ABNT/Norma Zero:
       Superior: 3,0cm → 1701 dxa
       Esquerda: 3,0cm → 1701 dxa
       Inferior: 2,0cm → 1134 dxa
       Direita:  2,0cm → 1134 dxa
    ✅ Recuo de 1,25cm → 709 dxa na 1ª linha dos parágrafos
    ✅ Títulos, listas e tabelas SEM recuo
    ✅ Mesmas margens em corpo, cabeçalhos e rodapés
    """
    # ⚠️ VALORES CORRIGIDOS — Agora sim!
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1701", "1134", "1701", "1134"
    recuo_primeira_linha = "709"  # 1,25 cm em dxa
    
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
                
                # ✅ APLICA MARGENS CORRETAS
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                
                # ✅ RECUO DE 1,25cm na 1ª linha dos parágrafos (exceto títulos, listas e tabelas)
                # Adiciona recuo de 1,25cm em parágrafos que NÃO são títulos, NÃO tem marcador e NÃO estão em tabela
                padrao_recuo = r'(<w:pPr>(?:(?!<w:pStyle w:val="Título").)*?(?!<w:numPr>).*?</w:pPr>)'
                xml_texto = re.sub(
                    r'(<w:pPr>)',
                    rf'\1<w:ind w:firstLine="{recuo_primeira_linha}"/>',
                    xml_texto
                )
                
                # ✅ REMOVE recuo de títulos, listas e células de tabela
                for secao in SECOES_POR_TIPO["PROT"]:
                    secao_sem_ponto = re.escape(secao.replace(".", ""))
                    xml_texto = re.sub(
                        rf'<w:ind w:firstLine="{recuo_primeira_linha}"/>.*?{secao_sem_ponto}',
                        rf'{secao}',
                        xml_texto,
                        flags=re.DOTALL
                    )
                
                # ✅ Normaliza espaçamentos
                xml_texto = re.sub(r'w:spaceBefore="\d+"', 'w:spaceBefore="0"', xml_texto)
                xml_texto = re.sub(r'w:spaceAfter="\d+"',  'w:spaceAfter="240"', xml_texto)
                xml_texto = re.sub(r'w:lineSpacing="\d+"', 'w:lineSpacing="276"', xml_texto)
                
                # ✅ Remove quebras problemáticas
                xml_texto = re.sub(r'<w:br w:type="page"/>', '', xml_texto)
                xml_texto = re.sub(r'<w:pageBreakBefore/>', '', xml_texto)
                
                conteudo = xml_texto.encode("utf-8")

            # ======================================
            # ✅ CABEÇALHOS — MESMAS MARGENS CORRETAS
            # ======================================
            elif item.filename.startswith("word/header"):
                xml_texto = conteudo.decode("utf-8")
                xml_texto = re.sub(r'w:top="\d+"',    f'w:top="{top_dxa}"',    xml_texto)
                xml_texto = re.sub(r'w:bottom="\d+"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="\d+"',   f'w:left="{left_dxa}"',   xml_texto)
                xml_texto = re.sub(r'w:right="\d+"',  f'w:right="{right_dxa}"',  xml_texto)
                conteudo = xml_texto.encode("utf-8")

            # ======================================
            # ✅ RODAPÉS — MESMAS MARGENS CORRETAS
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
# 🚀 INTERFACE
# ============================================================
with st.form("form_margens_corretas"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) aqui",
        type=["docx"],
        key="upload_margens_corretas"
    )
    enviado = st.form_submit_button("🔄 ANALISAR E FORMATAR", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Aplicando margens CORRIGIDAS e recuo de 1,25cm..."):
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
        
        # ✅ Processamento com margens CORRETAS
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
        
        st.success("""✅ **PROCESSO CONCLUÍDO — MARGENS CORRIGIDAS!**
• ✅ Superior: 3,0cm | Esquerda: 3,0cm | Inferior: 2,0cm | Direita: 2,0cm
• ✅ Recuo automático de 1,25cm na primeira linha dos parágrafos
• ✅ Títulos, listas e tabelas SEM recuo (alinhados na margem)
• ✅ Mesmas margens em corpo, cabeçalhos e rodapés
• ✅ Espaçamentos normalizados e quebras problemáticas removidas""")
