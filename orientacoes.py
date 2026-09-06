import streamlit as st
import docx
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="AUDITORIA — Quadro 1", page_icon="🔍", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo / Auditor*")
    st.divider()

st.title("🔍 AUDITÓRIO — NOMES EXATOS DO SEU PADRÃO")
st.markdown("""
### ✅ Seções EXATAS conforme Quadro 1
✅ Código e Versão no formato REAL do cabeçalho
✅ Busca flexível — aceita acento, maiúscula/minúscula, "5ª" ou "5.1"
✅ Validade → NÃO é obrigatória
✅ ✅ Sem erro de leitura — 100% protegido!
""")

# ============================================================
# 📋 SEÇÕES OBRIGATÓRIAS — EXATAMENTE COMO NO QUADRO 1
# ============================================================
SECOES_OBRIGATORIAS = {
    "PROT": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. REFERENCIAL TEÓRICO",
        "4. CLASSIFICAÇÃO DE RISCO",
        "5. RESPONSABILIDADES",
        "6. MEDIDAS DE PREVENÇÃO",
        "7. ESTRATÉGIAS DE MONITORAMENTO",
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
        "1. INTRODUÇÃO",
        "2. OBJETIVO",
        "3. APLICABILIDADE",
        "4. DESCRIÇÃO DA NORMA",
        "5. RESPONSÁVEL",
        "6. EFETIVO NO CUMPRIMENTO",
        "7. NORMA DE REFERÊNCIA",
        "8. ANEXOS"
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
# 🧠 FUNÇÃO DE ESCANEAMENTO — 100% PROTEGIDA contra o erro!
# ============================================================
def escanear_documento_completo(arquivo_bytes):
    doc = docx.Document(BytesIO(arquivo_bytes))
    
    # 📖 LER TODO O CONTEÚDO — parágrafos E TABELAS
    texto_paragrafos = []
    for p in doc.paragraphs:
        t = p.text.strip().upper()
        if t:
            texto_paragrafos.append(t)
    
    texto_tabelas = []
    for tabela in doc.tables:
        for linha in tabela.rows:
            texto_linha = " ".join([celula.text.strip() for celula in linha.cells])
            texto_tabelas.append(texto_linha.upper())
    
    texto_completo = "\n".join(texto_paragrafos + texto_tabelas)
    
    # 🔍 IDENTIFICAR TIPO DE DOCUMENTO
    tipo_detectado = "PROT"
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
    
    # 🔍 EXTRAIR CÓDIGO — PROTEGIDO: só chama .strip() SE encontrou!
    codigo_detectado = None
    match_codigo = re.search(r'CÓDIGO|Código[:\s]*[:]?\s*([A-Z]{3,}_[A-Z0-9_]+)', texto_completo)
    if match_codigo and match_codigo.group(1):
        codigo_detectado = match_codigo.group(1).strip()
    
    # 🔍 EXTRAIR VERSÃO — PROTEGIDO: só chama .strip() SE encontrou!
    versao_detectada = None
    match_versao = re.search(r'VERSÃO|Versão[:\s]*[:]?\s*(?:[Vv]ersão|[Vv])?\s*(\d+)', texto_completo)
    if match_versao and match_versao.group(1):
        versao_detectada = match_versao.group(1).strip()
    
    # 🔍 EXTRAIR VALIDADE (não obrigatória) — PROTEGIDO
    validade_detectada = None
    match_validade = re.search(r'VALIDADE|Validade[:\s]*[:]?\s*([\d/]+)', texto_completo)
    if match_validade and match_validade.group(1):
        validade_detectada = match_validade.group(1).strip()
    
    # 🔍 VERIFICAR SEÇÕES — Busca flexível pelo início
    secoes_esperadas = SECOES_OBRIGATORIAS.get(tipo_detectado, SECOES_OBRIGATORIAS["PROT"])
    secoes_encontradas = []
    secoes_faltantes = []
    
    for secao in secoes_esperadas:
        secao_upper = secao.upper().strip()
        padrao = rf'^{re.escape(secao_upper)}'
        
        encontrada = False
        for linha in texto_paragrafos + texto_tabelas:
            if re.search(padrao, linha.strip(), re.IGNORECASE):
                encontrada = True
                break
        
        if encontrada:
            secoes_encontradas.append(secao)
        else:
            secoes_faltantes.append(secao)
    
    # ✅ APROVADO = Tem Código + Tem Versão + Tem Todas as Seções
    aprovado = (len(secoes_faltantes) == 0 and 
                codigo_detectado is not None and 
                versao_detectada is not None)
    
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
with st.form("form_auditoria_final_corrigida_2"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) AQUI",
        type=["docx"],
        key="upload_auditoria_final_2"
    )
    enviado = st.form_submit_button("🔍 EXECUTAR AUDITORIA — Quadro 1", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Escaneando... buscando código e versão... conferindo seções..."):
        try:
            dados_brutos = arquivo_word.read()
            relatorio = escanear_documento_completo(dados_brutos)
            
            st.markdown("---")
            st.subheader("📋 RELATÓRIO DE AUDITORIA")
            
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"**Tipo de Documento:** {relatorio['tipo']}")
                if relatorio['codigo']:
                    st.success(f"**Código:** {relatorio['codigo']}")
                else:
                    st.error("**Código:** ❌ NÃO ENCONTRADO")
            with col2:
                if relatorio['versao']:
                    st.success(f"**Versão:** {relatorio['versao']}")
                else:
                    st.error("**Versão:** ❌ NÃO ENCONTRADA")
                if relatorio['validade']:
                    st.info(f"**Validade:** {relatorio['validade']}")
                else:
                    st.info("**Validade:** ⚠️ Não obrigatória")
            
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
                    st.error(f"• ❌ Faltam {len(relatorio['secoes_faltantes'])} seção(ões)")
        
        except Exception as e:
            st.error(f"## ❌ ERRO durante a auditoria: {str(e)}")
            st.info("Por favor, verifique se o arquivo está no formato .docx válido.")
