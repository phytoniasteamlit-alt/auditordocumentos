import streamlit as st
import docx
import re
from io import BytesIO

st.set_page_config(page_title="AUDITORIA — ACHA TUDO!", page_icon="🔍", layout="wide")

st.title("🔍 AUDITORIA — VERSÃO CORRIGIDA")
st.markdown("### ✅ Código completo • Versão com '5ª' • Seções com espaços e #")

# ============================================================
# 📋 SEÇÕES — NOMES EXATOS
# ============================================================
SECOES_POR_TIPO = {
    "PROT": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. REFERENCIAL TEÓRICO",
        "4. CLASSIFICAÇÃO DAS CIRURGIAS",
        "5. RESPONSABILIDADES",
        "6. MEDIDAS OBRIGATÓRIAS DE PREVENÇÃO",
        "7. ESTRATÉGIAS DE MONITORAMENTO",
        "8. REFERÊNCIAS"
    ],
    "POP": [
        "1. DEFINIÇÃO",
        "2. APLICABILIDADE",
        "3. RESPONSÁVEL",
        "4. DESCRIÇÃO DA EXECUÇÃO",
        "5. MATERIAIS UTILIZADOS",
        "6. TARIFA",
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
        "2. APLICABILIDADE",
        "3. DESCRIÇÃO DA NORMA",
        "4. RESPONSÁVEL",
        "5. EFETIVO NO CUMPRIMENTO",
        "6. NORMA DE REFERÊNCIA",
        "7. ANEXOS"
    ],
    "REG": [
        "1. FINALIDADE",
        "2. ÂMBITO",
        "3. COMPETÊNCIA E ORGANIZAÇÃO",
        "4. DISPOSIÇÕES GERAIS",
        "5. DISPOSIÇÕES FINAIS"
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
    "PLAN": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. DESCRIÇÃO DO CENÁRIO DE RISCO",
        "4. MEDIDAS DE CONTINGÊNCIA",
        "5. ESTRATÉGIAS DE RESPOSTA",
        "6. REFERÊNCIAS",
        "7. ANEXOS"
    ],
    "ROT": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. DESCRIÇÃO DA ROTINA",
        "4. RESPONSÁVEL",
        "5. ETAPAS DE EXECUÇÃO",
        "6. REFERÊNCIAS",
        "7. ANEXOS"
    ]
}

# ============================================================
# 🧠 FUNÇÃO DE AUDITORIA — CORRIGIDA
# ============================================================
def auditar_documento(arquivo_bytes):
    doc = docx.Document(BytesIO(arquivo_bytes))
    
    # ✅ LÊ TUDO e LIMPA: remove #, espaços excessivos, quebras de linha
    texto_bruto = ""
    for p in doc.paragraphs:
        texto_bruto += p.text.upper() + "\n"
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                texto_bruto += celula.text.upper() + " "
    
    # ✅ LIMPA O TEXTO: remove #, espaços múltiplos, tabulações
    texto = re.sub(r'[#\s]+', ' ', texto_bruto).strip()
    
    # 🔍 IDENTIFICAR TIPO
    tipo_detectado = "PROT"  # Padrão
    for tipo in SECOES_POR_TIPO.keys():
        if re.search(rf'\b{tipo}[_ /]', texto) or re.search(rf'\b{tipo}\b', texto):
            tipo_detectado = tipo
            break
    
    # 🔍 CÓDIGO — CORRIGIDO: pega TUDO (PROT_SCH005 completo!)
    codigo_detectado = None
    match_codigo = re.search(rf'{tipo_detectado}[_ ]?[A-Z0-9]+', texto)
    if match_codigo:
        codigo_detectado = re.sub(r'\s+', '', match_codigo.group(0))  # remove espaços
    
    # 🔍 VERSÃO — CORRIGIDO: aceita "5ª", "5", "V5.1"
    versao_detectada = None
    match_versao = re.search(r'VERSÃO[:\s]*[:]?\s*(?:VERSÃO|[Vv])?\s*(\d+)', texto)
    if match_versao and match_versao.group(1):
        versao_detectada = match_versao.group(1).strip()
    
    # 🔍 VALIDADE
    validade_detectada = None
    match_validade = re.search(r'VALIDADE[:\s]*[:]?\s*([\d/]+)', texto)
    if match_validade and match_validade.group(1):
        validade_detectada = match_validade.group(1).strip()
    
    # 🔍 SEÇÕES — CORRIGIDO: compara TEXTO LIMPO
    secoes_esperadas = SECOES_POR_TIPO[tipo_detectado]
    secoes_encontradas = []
    secoes_faltantes = []
    
    for secao in secoes_esperadas:
        secao_limpa = re.sub(r'[#\s]+', ' ', secao.upper()).strip()
        if secao_limpa in texto:
            secoes_encontradas.append(secao)
        else:
            secoes_faltantes.append(secao)
    
    aprovado = (len(secoes_faltantes) == 0 and 
                codigo_detectado is not None and 
                versao_detectada is not None)
    
    return {
        "tipo": tipo_detectado,
        "codigo": codigo_detectado,
        "versao": versao_detectada,
        "validade": validade_detectada,
        "secoes_encontradas": secoes_encontradas,
        "secoes_faltantes": secoes_faltantes,
        "aprovado": aprovado
    }

# ============================================================
# 🚀 INTERFACE
# ============================================================
with st.form("auditoria_final_corrigida"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) AQUI",
        type=["docx"]
    )
    enviado = st.form_submit_button("🔍 EXECUTAR AUDITORIA", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo: **{arquivo_word.name}**")
    
    with st.spinner("Escaneando..."):
        try:
            rel = auditar_documento(arquivo_word.read())
            
            st.markdown("---")
            c1, c2 = st.columns(2)
            with c1:
                st.info(f"**Tipo:** {rel['tipo']}")
                st.success(f"**Código:** {rel['codigo'] if rel['codigo'] else '❌ NÃO ENCONTRADO'}")
            with c2:
                st.success(f"**Versão:** {rel['versao'] if rel['versao'] else '❌ NÃO ENCONTRADA'}")
                st.info(f"**Validade:** {rel['validade'] if rel['validade'] else '⚠️ Não obrigatória'}")
            
            st.markdown("---")
            st.subheader("✅ SEÇÕES ENCONTRADAS")
            for s in rel["secoes_encontradas"]:
                st.success(f"✅ {s}")
            
            if rel["secoes_faltantes"]:
                st.subheader("❌ SEÇÕES FALTANTES")
                for s in rel["secoes_faltantes"]:
                    st.error(f"❌ {s}")
            else:
                st.subheader("✅ TODAS AS SEÇÕES ENCONTRADAS!")
            
            st.markdown("---")
            if rel["aprovado"]:
                st.success("## ✅ APROVADO!")
                st.balloons()
            else:
                st.error("## ❌ REPROVADO")
        
        except Exception as e:
            st.error(f"## ❌ ERRO: {str(e)}")
