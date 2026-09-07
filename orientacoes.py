import streamlit as st
import docx
import re
import unicodedata
from io import BytesIO

st.set_page_config(page_title="AUDITORIA — CABEÇALHO DETECTADO!", page_icon="🔍", layout="wide")

st.title("🔍 AUDITORIA — DETECTA CABEÇALHO DA TABELA")
st.markdown("### ✅ Busca exatamente no formato: Código: PROT_SCH005 • Versão: 5ª")

# ============================================================
# 🧹 REMOVER ACENTOS
# ============================================================
def limpar_texto(texto):
    if not texto:
        return ""
    sem_acento = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return re.sub(r'\s+', ' ', sem_acento.upper().strip())

# ============================================================
# 📋 SEÇÕES — SEM ACENTO
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
    ]
}

# ============================================================
# 🧠 FUNÇÃO PRINCIPAL — DETECTA CABEÇALHO DA TABELA
# ============================================================
def auditar_documento(arquivo_bytes):
    doc = docx.Document(BytesIO(arquivo_bytes))
    
    # ✅ LER CÉLULA POR CÉLULA DA TABELA DO CABEÇALHO
    codigo_detectado = None
    versao_detectada = None
    validade_detectada = None
    texto_completo = ""
    
    # 🔍 LER TABELAS — AQUI ESTÃO CÓDIGO E VERSÃO!
    for tabela in doc.tables:
        for linha in tabela.rows:
            # Junta o texto da linha e limpa
            linha_texto = " ".join([cel.text for cel in linha.cells])
            texto_completo += linha_texto + " "
            
            # ✅ BUSCA CÓDIGO — exatamente no formato "Código: PROT_SCH005"
            cod_match = re.search(r'CÓDIGO|Código[:\s]*[:]?\s*([A-Z]{3,4}_[A-Z0-9]+)', linha_texto, re.IGNORECASE)
            if cod_match and cod_match.group(1):
                codigo_detectado = cod_match.group(1).strip()
            
            # ✅ BUSCA VERSÃO — exatamente no formato "Versão: 5ª"
            ver_match = re.search(r'VERSÃO|Versão[:\s]*[:]?\s*(?:[Vv]|versão)?\s*(\d+)', linha_texto, re.IGNORECASE)
            if ver_match and ver_match.group(1):
                versao_detectada = ver_match.group(1).strip()
            
            # ✅ BUSCA VALIDADE
            val_match = re.search(r'VALIDADE|Validade[:\s]*[:]?\s*([\d/]+)', linha_texto, re.IGNORECASE)
            if val_match and val_match.group(1):
                validade_detectada = val_match.group(1).strip()
    
    # ✅ LER PARÁGRAFOS do corpo
    for p in doc.paragraphs:
        texto_completo += p.text + " "
    
    # ✅ LIMPA TODO O TEXTO para busca de seções
    texto_limpo = limpar_texto(texto_completo)
    
    # 🔍 IDENTIFICAR TIPO
    tipo_detectado = "PROT"
    for tipo in SECOES_POR_TIPO.keys():
        if re.search(rf'\b{tipo}[_ /]', texto_limpo) or re.search(rf'\b{tipo}\b', texto_limpo):
            tipo_detectado = tipo
            break
    
    # 🔍 VERIFICAR SEÇÕES
    secoes_esperadas = SECOES_POR_TIPO[tipo_detectado]
    secoes_encontradas = []
    secoes_faltantes = []
    
    for secao in secoes_esperadas:
        secao_limpa = limpar_texto(secao)
        if re.search(rf'\b{re.escape(secao_limpa)}\b', texto_limpo):
            secoes_encontradas.append(secao)
        else:
            secoes_faltantes.append(secao)
    
    # ✅ APROVADO
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
with st.form("auditoria_cabecalho_tabela"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) AQUI",
        type=["docx"]
    )
    enviado = st.form_submit_button("🔍 EXECUTAR AUDITORIA — LER CABEÇALHO", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo: **{arquivo_word.name}**")
    
    with st.spinner("Lendo tabela do cabeçalho... buscando Código e Versão..."):
        try:
            rel = auditar_documento(arquivo_word.read())
            
            st.markdown("---")
            st.subheader("📋 RELATÓRIO DE AUDITORIA")
            
            c1, c2 = st.columns(2)
            with c1:
                st.info(f"**Tipo de Documento:** {rel['tipo']}")
                if rel['codigo']:
                    st.success(f"**Código:** {rel['codigo']} ✅")
                else:
                    st.error("**Código:** ❌ NÃO ENCONTRADO")
            with c2:
                if rel['versao']:
                    st.success(f"**Versão:** {rel['versao']} ✅")
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
            st.error(f"## ❌ ERRO: {str(e)}")
