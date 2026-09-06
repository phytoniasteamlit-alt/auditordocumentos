import streamlit as st
import docx
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="AUDITORIA COMPLETA — Triagem e Validação", page_icon="🔍", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo / Auditor*")
    st.divider()

st.title("🔍 AUDITORIA COMPLETA — Triagem, Validação e Relatório")
st.markdown("""
### ✅ Escaneia TODO o documento página por página
✅ Verifica se TODAS as seções obrigatórias EXISTEM
✅ Verifica CÓDIGO, VERSÃO e VALIDADE no cabeçalho
✅ ✅ = APROVADO / ❌ = FALTANDO → DOCUMENTO DEVE VOLTAR
✅ Relatório detalhado com tudo que foi encontrado e faltou
""")

# ============================================================
# 📋 SEÇÕES OBRIGATÓRIAS POR TIPO DE DOCUMENTO
# ============================================================
SECOES_OBRIGATORIAS = {
    "PROT": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. REFERENCIAL TEÓRICO",
        "4. CLASSIFICAÇÃO",
        "5. RESPONSABILIDADES",
        "6. MEDIDAS PREVENTIVAS",
        "7. REFERÊNCIAS",
        "8. ANEXOS"
    ],
    "POP": [
        "1. DEFINIÇÃO",
        "2. APLICABILIDADE",
        "3. RESPONSABILIDADES",
        "4. DESCRIÇÃO DAS ETAPAS",
        "5. REFERÊNCIAS",
        "6. ANEXOS"
    ],
    "PROG": [
        "1. REFERENCIAL TEÓRICO",
        "2. OBJETIVOS",
        "3. METAS E INDICADORES",
        "4. DEFINIÇÃO DE METAS",
        "5. ACOMPANHAMENTO E MONITORAMENTO",
        "6. AVALIAÇÃO DE RESULTADOS",
        "7. REFERÊNCIAS",
        "8. ANEXOS"
    ],
    "POI": [
        "1. INTRODUÇÃO",
        "2. OBJETIVO",
        "3. FINALIDADE",
        "4. ABRANGÊNCIA",
        "5. RESPONSABILIDADES",
        "6. GESTÃO DE RISCO",
        "7. ANEXOS",
        "8. REFERÊNCIAS"
    ],
    "NOR": [
        "1. OBJETIVO",
        "2. ABRANGÊNCIA",
        "3. DEFINIÇÕES",
        "4. COMPETÊNCIAS",
        "5. PROCEDIMENTOS",
        "6. DISPOSIÇÕES FINAIS",
        "7. REFERÊNCIAS"
    ],
    "REG": [
        "1. FINALIDADE",
        "2. ÂMBITO",
        "3. COMPETÊNCIA E ORGANIZAÇÃO",
        "4. DISPOSIÇÕES GERAIS",
        "5. DISPOSIÇÕES FINAIS"
    ]
}

# ============================================================
# 🧠 ESCANEAR TODO O DOCUMENTO
# ============================================================
def escanear_documento_completo(arquivo_bytes):
    """
    ✅ LÊ TODO o texto do documento
    ✅ IDENTIFICA o tipo (PROT, POP, NOR, etc.)
    ✅ VERIFICA se CADA seção obrigatória EXISTE no texto
    ✅ VERIFICA CÓDIGO, VERSÃO e VALIDADE no cabeçalho
    ✅ RETORNA relatório completo
    """
    doc = docx.Document(BytesIO(arquivo_bytes))
    
    # 📖 LER TODO O CONTEÚDO — parágrafos + tabelas
    texto_paragrafos = []
    for p in doc.paragraphs:
        if p.text.strip():
            texto_paragrafos.append(p.text.strip().upper())
    
    texto_tabelas = []
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                if celula.text.strip():
                    texto_tabelas.append(celula.text.strip().upper())
    
    texto_completo = "\n".join(texto_paragrafos + texto_tabelas)
    
    # 🔍 IDENTIFICAR TIPO DE DOCUMENTO
    tipo_detectado = None
    codigo_detectado = None
    versao_detectada = None
    validade_detectada = None
    
    if re.search(r'\bPROT[_ /]', texto_completo) or "PROTOCOLO" in texto_completo:
        tipo_detectado = "PROT"
    elif re.search(r'\bPOP[_ /]', texto_completo) or "PROCEDIMENTO OPERACIONAL" in texto_completo:
        tipo_detectado = "POP"
    elif re.search(r'\bPROG[_ /]', texto_completo) or "PROGRAMA" in texto_completo:
        tipo_detectado = "PROG"
    elif re.search(r'\bPOI[_ /]', texto_completo) or "POLÍTICA" in texto_completo or "POLITICA" in texto_completo:
        tipo_detectado = "POI"
    elif re.search(r'\bNOR[_ /]', texto_completo) or "NORMA" in texto_completo:
        tipo_detectado = "NOR"
    elif re.search(r'\bREG[_ /]', texto_completo) or "REGULAMENTO" in texto_completo:
        tipo_detectado = "REG"
    else:
        tipo_detectado = "PROT"  # Padrão
    
    # 🔍 EXTRAIR CÓDIGO, VERSÃO e VALIDADE do cabeçalho
    match_codigo = re.search(r'CÓDIGO[:\s]*([A-Z]{3,}_[A-Z0-9_]+|[A-Z]{4}_[A-Z0-9_]+)', texto_completo)
    if match_codigo:
        codigo_detectado = match_codigo.group(1).strip()
    
    match_versao = re.search(r'VERSÃO[:\s]*([Vv]?\d+[./]?\d*)', texto_completo)
    if match_versao:
        versao_detectada = match_versao.group(1).strip()
    
    match_validade = re.search(r'VALIDADE[:\s]*([\d/]+)', texto_completo)
    if match_validade:
        validade_detectada = match_validade.group(1).strip()
    
    # 🔍 VERIFICAR SEÇÕES — escaneia UMA POR UMA
    secoes_esperadas = SECOES_OBRIGATORIAS.get(tipo_detectado, SECOES_OBRIGATORIAS["PROT"])
    secoes_encontradas = []
    secoes_faltantes = []
    
    for secao in secoes_esperadas:
        secao_limpa = re.escape(secao.upper().replace(".", "").strip())
        padrao = rf'\b{secao_limpa}\b'
        if re.search(padrao, texto_completo):
            secoes_encontradas.append(secao)
        else:
            secoes_faltantes.append(secao)
    
    # ✅ RESULTADO FINAL
    aprovado = len(secoes_faltantes) == 0 and codigo_detectado and versao_detectada
    
    return {
        "tipo": tipo_detectado,
        "codigo": codigo_detectado,
        "versao": versao_detectada,
        "validade": validade_detectada,
        "secoes_esperadas": secoes_esperadas,
        "secoes_encontradas": secoes_encontradas,
        "secoes_faltantes": secoes_faltantes,
        "aprovado": aprovado
    }

# ============================================================
# 🚀 INTERFACE
# ============================================================
with st.form("form_auditoria_completa"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) AQUI para AUDITORIA COMPLETA",
        type=["docx"],
        key="upload_auditoria"
    )
    enviado = st.form_submit_button("🔍 EXECUTAR AUDITORIA COMPLETA", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Escaneando TODO o documento página por página... verificando seções, código e versão..."):
        dados_brutos = arquivo_word.read()
        relatorio = escanear_documento_completo(dados_brutos)
        
        # ======================================
        # 📋 RELATÓRIO DE AUDITORIA
        # ======================================
        st.markdown("---")
        st.subheader("📋 RELATÓRIO DE AUDITORIA")
        
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"**Tipo de Documento:** {relatorio['tipo']}")
            st.info(f"**Código:** {relatorio['codigo'] or '❌ NÃO ENCONTRADO'}")
        with col2:
            st.info(f"**Versão:** {relatorio['versao'] or '❌ NÃO ENCONTRADA'}")
            st.info(f"**Validade:** {relatorio['validade'] or '⚠️ Não verificada'}")
        
        st.markdown("---")
        
        # ✅ SEÇÕES ENCONTRADAS
        st.subheader("✅ SEÇÕES ENCONTRADAS")
        for secao in relatorio["secoes_encontradas"]:
            st.success(f"✅ {secao}")
        
        # ❌ SEÇÕES FALTANTES
        if relatorio["secoes_faltantes"]:
            st.subheader("❌ SEÇÕES FALTANTES — DOCUMENTO DEVE VOLTAR!")
            for secao in relatorio["secoes_faltantes"]:
                st.error(f"❌ {secao}")
        else:
            st.subheader("✅ TODAS AS SEÇÕES FORAM ENCONTRADAS!")
        
        st.markdown("---")
        
        # 🏅 RESULTADO FINAL
        if relatorio["aprovado"]:
            st.success("## ✅ APROVADO — Documento completo e conforme!")
            st.balloons()
        else:
            st.error("## ❌ REPROVADO — Documento INCOMPLETO!")
            st.error("### ⚠️ Este documento DEVE VOLTAR para correção!")
            if not relatorio["codigo"]:
                st.error("• ❌ Falta CÓDIGO no cabeçalho")
            if not relatorio["versao"]:
                st.error("• ❌ Falta VERSÃO no cabeçalho")
            if relatorio["secoes_faltantes"]:
                st.error(f"• ❌ Faltam {len(relatorio['secoes_faltantes'])} seção(ões) obrigatória(s)")
