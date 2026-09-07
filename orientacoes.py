import streamlit as st
import docx
import re
import unicodedata
from io import BytesIO

st.set_page_config(page_title="AUDITORIA — SEM ACENTO!", page_icon="🔍", layout="wide")

st.title("🔍 AUDITORIA — BUSCA SEM ACENTO")
st.markdown("### ✅ Ignora acento, maiúscula/minúscula e espaços → ACHA TUDO!")

# ============================================================
# 🧹 FUNÇÃO: REMOVER ACENTOS E NORMALIZAR TEXTO
# ============================================================
def limpar_texto(texto):
    """Remove acentos, converte para maiúsculo, remove espaços extras"""
    if not texto:
        return ""
    # Normaliza e remove acentos mantendo letras e números limpos
    sem_acento = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    # Substitui quebras de linha por espaços para evitar palavras coladas
    sem_acento = sem_acento.replace('\n', ' ').replace('\r', ' ')
    # Força maiúsculo e limpa espaços múltiplos
    return re.sub(r'\s+', ' ', sem_acento.upper().strip())

# ============================================================
# 📋 SEÇÕES — NOMES BASE (sem acento na busca)
# ============================================================
SECOES_POR_TIPO = {
    "PROT": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. REFERENCIAL TEORICO",
        "4. CLASSIFICACAO",
        "5. RESPONSABILIDADES",
        "6. MEDIDAS OBRIGATORIAS",
        "7. ESTRATEGIAS DE MONITORAMENTO",
        "8. REFERENCIAS"
    ],
    "POP": [
        "1. DEFINICAO",
        "2. APLICABILIDADE",
        "3. RESPONSAVEL",
        "4. DESCRICAO DA EXECUCAO",
        "5. MATERIAIS UTILIZADOS",
        "6. TARIFA",
        "7. REFERENCIAS",
        "8. ANEXOS"
    ],
    "POI": [
        "1. INTRODUCAO",
        "2. OBJETIVO",
        "3. FINALIDADE",
        "4. ABRANGENCIA",
        "5. RESPONSABILIDADES",
        "6. GESTAO DE RISCO",
        "7. ANEXOS",
        "8. REFERENCIAS"
    ],
    "NOR": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. DESCRICAO DA NORMA",
        "4. RESPONSAVEL",
        "5. EFETIVO NO CUMPRIMENTO",
        "6. NORMA DE REFERENCIA",
        "7. ANEXOS"
    ],
    "REG": [
        "1. FINALIDADE",
        "2. AMBITO",
        "3. COMPETENCIA E ORGANIZACAO",
        "4. DISPOSICOES GERAIS",
        "5. DISPOSICOES FINAIS"
    ],
    "PROG": [
        "1. REFERENCIAL TEORICO",
        "2. OBJETIVOS",
        "3. METAS E INDICADORES",
        "4. DEFINICAO DE METAS",
        "5. ACOMPANHAMENTO E MONITORAMENTO",
        "6. AVALIACAO DE RESULTADOS",
        "7. REFERENCIAS",
        "8. ANEXOS"
    ],
    "PLAN": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. DESCRICAO DO CENARIO DE RISCO",
        "4. MEDIDAS DE CONTINGENCIA",
        "5. ESTRATEGIAS DE RESPOSTA",
        "6. REFERENCIAS",
        "7. ANEXOS"
    ],
    "ROT": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. DESCRICAO DA ROTINA",
        "4. RESPONSAVEL",
        "5. ETAPAS DE EXECUCAO",
        "6. REFERENCIAS",
        "7. ANEXOS"
    ]
}

# ============================================================
# 🧠 FUNÇÃO DE AUDITORIA — CORRIGIDA!
# ============================================================
def auditar_documento(arquivo_bytes):
    doc = docx.Document(BytesIO(arquivo_bytes))
    
    # Extrai o texto preservando espaços entre parágrafos e células de tabela
    elementos_texto = []
    for p in doc.paragraphs:
        if p.text.strip():
            elementos_texto.append(p.text)
            
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                if celula.text.strip():
                    elementos_texto.append(celula.text)
    
    # Junta tudo com um espaço duplo seguro para evitar colisão de dados da tabela
    texto_bruto = "  ".join(elementos_texto)
    texto = limpar_texto(texto_bruto)
    
    # 🔍 IDENTIFICAR TIPO DO DOCUMENTO
    tipo_detectado = "PROT"
    for tipo in SECOES_POR_TIPO.keys():
        if re.search(rf'\b{tipo}\b', texto):
            tipo_detectado = tipo
            break
    
    # 🔍 CÓDIGO — Busca flexível e abrangente (aceita letras, números, underlines e traços)
    codigo_detectado = None
    match_codigo = re.search(r'CODIGO[\s:]+([A-Z0-9_|-]+)', texto)
    if match_codigo:
        codigo_detectado = match_codigo.group(1).strip()
    
    # 🔍 VERSÃO — Captura o dígito puro numérico isolado
    versao_detectada = None
    match_versao = re.search(r'VERSAO[\s:]+(\d+)', texto)
    if match_versao:
        versao_detectada = match_versao.group(1).strip()
    
    # 🔍 VALIDADE
    validade_detectada = None
    match_validade = re.search(r'VALIDADE[\s:]*([\d/X]+)', texto)
    if match_validade:
        validade_detectada = match_validade.group(1).strip()
    
    # 🔍 SEÇÕES — COMPARA SEM ACENTO E SEM PREFIXO NUMÉRICO
    secoes_esperadas = SECOES_POR_TIPO[tipo_detectado]
    secoes_encontradas = []
    secoes_faltantes = []
    
    for secao in secoes_esperadas:
        # Remove a numeração inicial do título (ex: "1. OBJETIVO" vira "OBJETIVO")
        secao_sem_numero = re.sub(r'^\d+\.\s*', '', secao)
        secao_limpa = limpar_texto(secao_sem_numero)
        
        # Procura o nome do título isolado como uma palavra inteira
        if re.search(rf'\b{re.escape(secao_limpa)}\b', texto):
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
# 🚀 INTERFACE STREAMLIT
# ============================================================
with st.form("auditoria_sem_acento_final"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) AQUI",
        type=["docx"]
    )
    enviado = st.form_submit_button("🔍 EXECUTAR AUDITORIA — SEM ACENTO", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo: **{arquivo_word.name}**")
    
    with st.spinner("Escaneando... ignorando acentos e diferenças..."):
        try:
            rel = auditar_documento(arquivo_word.read())
            
            st.markdown("---")
            c1, c2 = st.columns(2)
            with c1:
                st.info(f"**Tipo:** {rel['tipo']}")
                if rel['codigo']:
                    st.success(f"**Código:** {rel['codigo']}")
                else:
                    st.error("**Código:** ❌ NÃO ENCONTRADO")
            with c2:
                if rel['versao']:
                    st.success(f"**Versão:** {rel['versao']}")
                else:
                    st.error("**Versão:** ❌ NÃO ENCONTRADA")
                if rel['validade']:
                    st.info(f"**Validade:** {rel['validade']}")
                else:
                    st.info("**Validade:** ⚠️ Não obrigatória")
            
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
                st.success("## ✅ APROVADO — DOCUMENTO CONFORME!")
                st.balloons()
            else:
                st.error("## ❌ REPROVADO — Verifique os itens acima")
        
        except Exception as e:
            st.error(f"## ❌ ERRO INESPERADO: {str(e)}")
