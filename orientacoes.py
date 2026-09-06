import streamlit as st
import docx
import re
from io import BytesIO

st.set_page_config(page_title="AUDITORIA COMPLETA — Todos os Tipos", page_icon="🔍", layout="wide")

st.title("🔍 AUDITORIA COMPLETA — Todos os Tipos de Documento")
st.markdown("""
### ✅ Reconhece: PROT, POP, POI, NOR, REG, PROG, PLAN e ROT
✅ Seções EXATAMENTE conforme Quadro 1
✅ Busca simples — lê tudo e verifica se está presente
✅ Código e Versão detectados de forma flexível
✅ Validade → NÃO é obrigatória
""")

# ============================================================
# 📋 SEÇÕES OBRIGATÓRIAS — TODOS OS TIPOS conforme Quadro 1
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
# 🧠 FUNÇÃO DE AUDITORIA — SIMPLES E DIRETA
# ============================================================
def auditar_documento(arquivo_bytes):
    doc = docx.Document(BytesIO(arquivo_bytes))
    
    # ✅ LÊ TUDO — junta parágrafos e TODAS as tabelas em UMA string MAIÚSCULA
    texto_completo = ""
    for p in doc.paragraphs:
        texto_completo += p.text.upper() + "\n"
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                texto_completo += celula.text.upper() + " "
    
    # 🔍 IDENTIFICAR TIPO DE DOCUMENTO
    tipo_detectado = None
    for tipo in SECOES_POR_TIPO.keys():
        if re.search(rf'\b{tipo}[_ /]', texto_completo) or re.search(rf'\b{tipo}\b', texto_completo):
            tipo_detectado = tipo
            break
    
    if not tipo_detectado:
        tipo_detectado = "PROT"  # Padrão
    
    # 🔍 EXTRAIR CÓDIGO — flexível: PROT_..., POP_..., etc.
    codigo_detectado = None
    match_codigo = re.search(rf'{tipo_detectado}[_ ][A-Z0-9_]+', texto_completo)
    if match_codigo:
        codigo_detectado = match_codigo.group(0).strip()
    
    # 🔍 EXTRAIR VERSÃO — aceita "5ª", "5", "V5.1"
    versao_detectada = None
    match_versao = re.search(r'VERSÃO|VERSÃO[:\s]*(\d+)', texto_completo)
    if match_versao and match_versao.group(1):
        versao_detectada = match_versao.group(1).strip()
    
    # 🔍 EXTRAIR VALIDADE (não obrigatória)
    validade_detectada = None
    match_validade = re.search(r'VALIDADE[:\s]*([\d/]+)', texto_completo)
    if match_validade and match_validade.group(1):
        validade_detectada = match_validade.group(1).strip()
    
    # 🔍 VERIFICAR SEÇÕES — busca simples: "o texto existe lá?"
    secoes_esperadas = SECOES_POR_TIPO[tipo_detectado]
    secoes_encontradas = []
    secoes_faltantes = []
    
    for secao in secoes_esperadas:
        secao_limpa = secao.upper().strip()
        if secao_limpa in texto_completo:
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
with st.form("form_auditoria_completa_todos_tipos"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) AQUI",
        type=["docx"],
        key="upload_auditoria_completa"
    )
    enviado = st.form_submit_button("🔍 EXECUTAR AUDITORIA COMPLETA", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Escaneando... identificando tipo... verificando seções..."):
        try:
            dados_brutos = arquivo_word.read()
            relatorio = auditar_documento(dados_brutos)
            
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
                    st.error(f"• ❌ Faltam {len(relatorio['secoes_faltantes'])} seção(ões) obrigatória(s)")
        
        except Exception as e:
            st.error(f"## ❌ ERRO durante a auditoria: {str(e)}")
            st.info("Por favor, verifique se o arquivo está no formato .docx válido.")
