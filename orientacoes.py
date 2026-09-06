import streamlit as st
import docx
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="AUDITORIA — NOMES REAIS DO SEU DOCUMENTO", page_icon="🔍", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo / Auditor*")
    st.divider()

st.title("🔍 AUDITORIA — NOMES EXATOS DO SEU PADRÃO")
st.markdown("""
### ✅ Nomes das seções COPIADOS do SEU documento real
✅ Código e Versão lidos da TABELA do cabeçalho
✅ Data de aprovação → NÃO é obrigatória
✅ Busca pelo INÍCIO do título → encontra mesmo com texto maior
""")

# ============================================================
# 📋 SEÇÕES OBRIGATÓRIAS — NOMES EXATOS CONFORME SEU DOCUMENTO
# ============================================================
SECOES_OBRIGATORIAS = {
    "PROT": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. REFERENCIAL TEÓRICO",
        "4. CLASSIFICAÇÃO",
        "5. RESPONSABILIDADES",
        "6. MEDIDAS OBRIGATÓRIAS",
        "7. ESTRATÉGIAS",
        "8. REFERÊNCIAS"
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
# 🧠 ESCANEAR — Busca inteligente + Cabeçalho em TABELA
# ============================================================
def escanear_documento_completo(arquivo_bytes):
    doc = docx.Document(BytesIO(arquivo_bytes))
    
    # 📖 LER TODO O CONTEÚDO — parágrafos e TABELAS (cabeçalho é tabela!)
    texto_paragrafos = []
    for p in doc.paragraphs:
        texto_paragrafos.append(p.text.strip().upper())
    
    texto_tabelas = []
    for tabela in doc.tables:
        for linha in tabela.rows:
            texto_linha = " ".join([celula.text.strip() for celula in linha.cells])
            texto_tabelas.append(texto_linha.upper())
    
    texto_completo = "\n".join(texto_paragrafos + texto_tabelas)
    
    # 🔍 IDENTIFICAR TIPO
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
        tipo_detectado = "PROT"
    
    # 🔍 EXTRAIR CÓDIGO — busca no formato da TABELA do cabeçalho
    # Ex: "Código: PROT_SCIH005"
    match_codigo = re.search(r'CÓDIGO[:\s]*[:]?\s*([A-Z]{3,}_[A-Z0-9_]+)', texto_completo)
    if match_codigo:
        codigo_detectado = match_codigo.group(1).strip()
    
    # 🔍 EXTRAIR VERSÃO — busca no formato da TABELA
    # Ex: "Versão: 5.1"
    match_versao = re.search(r'VERSÃO[:\s]*[:]?\s*[Vv]?(\d+[./]?\d*)', texto_completo)
    if match_versao:
        versao_detectada = match_versao.group(1).strip()
    
    # 🔍 EXTRAIR VALIDADE (não é obrigatória)
    match_validade = re.search(r'VALIDADE[:\s]*[:]?\s*([\d/]+)', texto_completo)
    if match_validade:
        validade_detectada = match_validade.group(1).strip()
    
    # 🔍 VERIFICAR SEÇÕES — Busca pelo INÍCIO do número + nome
    secoes_esperadas = SECOES_OBRIGATORIAS.get(tipo_detectado, SECOES_OBRIGATORIAS["PROT"])
    secoes_encontradas = []
    secoes_faltantes = []
    
    for secao in secoes_esperadas:
        secao_upper = secao.upper().strip()
        # ✅ Busca se QUALQUER linha COMEÇA com o número + nome da seção
        padrao = rf'^{re.escape(secao_upper)}'
        
        encontrada = False
        for linha in texto_paragrafos:
            if re.search(padrao, linha.strip()):
                encontrada = True
                break
        if not encontrada:
            for linha in texto_tabelas:
                if re.search(padrao, linha.strip()):
                    encontrada = True
                    break
        
        if encontrada:
            secoes_encontradas.append(secao)
        else:
            secoes_faltantes.append(secao)
    
    # ✅ APROVADO = Tem TUDO obrigatório: Código + Versão + Todas as seções
    # ⚠️ Data de aprovação NÃO é obrigatória!
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
with st.form("form_auditoria_final"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) AQUI",
        type=["docx"],
        key="upload_auditoria_final"
    )
    enviado = st.form_submit_button("🔍 EXECUTAR AUDITORIA COMPLETA", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Escaneando... buscando código da tabela do cabeçalho... verificando seções..."):
        dados_brutos = arquivo_word.read()
        relatorio = escanear_documento_completo(dados_brutos)
        
        st.markdown("---")
        st.subheader("📋 RELATÓRIO DE AUDITORIA")
        
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"**Tipo de Documento:** {relatorio['tipo']}")
            st.success(f"**Código:** {relatorio['codigo'] or '❌ NÃO ENCONTRADO'}")
        with col2:
            st.success(f"**Versão:** {relatorio['versao'] or '❌ NÃO ENCONTRADA'}")
            if relatorio['validade']:
                st.info(f"**Validade:** {relatorio['validade']}")
            else:
                st.info("**Validade:** ⚠️ Não obrigatória / não encontrada")
        
        st.markdown("---")
        
        st.subheader("✅ SEÇÕES ENCONTRADAS")
        for secao in relatorio["secoes_encontradas"]:
            st.success(f"✅ {secao}")
        
        if relatorio["secoes_faltantes"]:
            st.subheader("❌ SEÇÕES FALTANTES")
            for secao in relatorio["secoes_faltantes"]:
                st.error(f"❌ {secao}")
        else:
            st.subheader("✅ TODAS AS SEÇÕES FORAM ENCONTRADAS!")
        
        st.markdown("---")
        
        if relatorio["aprovado"]:
            st.success("## ✅ APROVADO — Documento COMPLETO e CONFORME!")
            st.balloons()
        else:
            st.error("## ❌ REPROVADO — Verifique os itens acima!")
            if not relatorio["codigo"]:
                st.error("• ❌ Falta CÓDIGO no cabeçalho")
            if not relatorio["versao"]:
                st.error("• ❌ Falta VERSÃO no cabeçalho")
            if relatorio["secoes_faltantes"]:
                st.error(f"• ❌ Faltam {len(relatorio['secoes_faltantes'])} seção(ões) obrigatória(s)")
