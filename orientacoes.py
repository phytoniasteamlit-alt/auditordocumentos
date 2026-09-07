import streamlit as st
import docx
import re
import unicodedata
from io import BytesIO

st.set_page_config(page_title="AUDITORIA — FINAL COMPLETA", page_icon="🔍", layout="wide")

st.title("🔍 AUDITORIA DE DOCUMENTOS — SISTEMA COMPLETO")
st.markdown("### ✅ Detecta Cabeçalho da Tabela • Ignora Acentos • Reconhece Todos os Tipos")

# ============================================================
# 🧹 FUNÇÃO: REMOVER ACENTOS E NORMALIZAR TEXTO
# ============================================================
def limpar_texto(texto):
    if not texto:
        return ""
    sem_acento = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return re.sub(r'\s+', ' ', sem_acento.upper().strip())

# ============================================================
# 📋 SEÇÕES POR TIPO DE DOCUMENTO
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
# 🧹 APLICAR MARGENS PADRÃO ABNT (3/3/2/2 cm)
# ============================================================
def injetar_margens_via_xml_puro(doc_bytes):
    doc = docx.Document(BytesIO(doc_bytes))
    # Aplica margens ABNT: Sup 3cm, Inf 2cm, Esq 3cm, Dir 2cm
    sections = doc.sections
    for sec in sections:
        sec.top_margin = docx.shared.Cm(3)
        sec.bottom_margin = docx.shared.Cm(2)
        sec.left_margin = docx.shared.Cm(3)
        sec.right_margin = docx.shared.Cm(2)
    output = BytesIO()
    doc.save(output)
    return output.getvalue()

# ============================================================
# 📄 GERAR FICHA DE VERIFICAÇÃO
# ============================================================
def gerar_ficha_naqh(tipo, codigo, versao, encontradas, faltantes, aprovado):
    ficha = []
    ficha.append("=" * 60)
    ficha.append("FICHA DE VERIFICAÇÃO DE DOCUMENTO")
    ficha.append("=" * 60)
    ficha.append(f"Tipo de Documento: {tipo}")
    ficha.append(f"Código: {codigo}")
    ficha.append(f"Versão: {versao}")
    ficha.append("")
    ficha.append("--- SEÇÕES ENCONTRADAS ---")
    for s in encontradas:
        ficha.append(f"✓ {s}")
    ficha.append("")
    if faltantes:
        ficha.append("--- SEÇÕES FALTANTES ---")
        for s in faltantes:
            ficha.append(f"✗ {s}")
    else:
        ficha.append("--- TODAS AS SEÇÕES FORAM ENCONTRADAS ---")
    ficha.append("")
    ficha.append("=" * 60)
    status = "APROVADO" if aprovado else "REPROVADO / COM PENDÊNCIAS"
    ficha.append(f"STATUS FINAL: {status}")
    ficha.append("=" * 60)
    return "\n".join(ficha).encode("utf-8")

# ============================================================
# 🧠 FUNÇÃO PRINCIPAL — LÊ TABELA DO CABEÇALHO
# ============================================================
def auditar_documento(arquivo_bytes):
    doc = docx.Document(BytesIO(arquivo_bytes))
    
    codigo_detectado = None
    versao_detectada = None
    validade_detectada = None
    texto_completo = ""
    
    # 🔍 LER TABELAS — onde fica o Cabeçalho!
    for tabela in doc.tables:
        for linha in tabela.rows:
            linha_texto = " ".join([cel.text for cel in linha.cells])
            texto_completo += linha_texto + " "
            
            # ✅ BUSCA CÓDIGO — aceita "Código:" ou "CÓDIGO:"
            cod_match = re.search(r'CÓDIGO|Código[:\s]*[:]?\s*([A-Z]{3,4}_[A-Z0-9]+)', linha_texto, re.IGNORECASE)
            if cod_match and cod_match.group(1):
                codigo_detectado = cod_match.group(1).strip()
            
            # ✅ BUSCA VERSÃO — aceita "Versão: 5ª" → pega só o número
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
    
    return {
        "tipo": tipo_detectado,
        "codigo": codigo_detectado or "NÃO DETECTADO",
        "versao": versao_detectada or "NÃO DETECTADA",
        "validade": validade_detectada,
        "secoes_encontradas": secoes_encontradas,
        "secoes_faltantes": secoes_faltantes
    }

# ============================================================
# 🚀 INTERFACE PRINCIPAL
# ============================================================
with st.form("auditoria_completa_final"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) AQUI",
        type=["docx"]
    )
    enviado = st.form_submit_button("🔍 EXECUTAR AUDITORIA COMPLETA", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Lendo tabela do cabeçalho... verificando seções... aplicando margens..."):
        try:
            dados_brutos = arquivo_word.read()
            rel = auditar_documento(dados_brutos)
            
            tipo_detectado = rel["tipo"]
            codigo_doc = rel["codigo"]
            versao_doc = rel["versao"]
            validade_doc = rel["validade"]
            secoes_encontradas = rel["secoes_encontradas"]
            secoes_faltantes = rel["secoes_faltantes"]
            
            st.markdown("---")
            st.subheader("📋 RELATÓRIO DE AUDITORIA")
            
            c1, c2 = st.columns(2)
            with c1:
                st.info(f"**Tipo de Documento:** {tipo_detectado}")
                if codigo_doc != "NÃO DETECTADO":
                    st.success(f"**Código:** {codigo_doc} ✅")
                else:
                    st.error("**Código:** ❌ NÃO DETECTADO")
            with c2:
                if versao_doc != "NÃO DETECTADA":
                    st.success(f"**Versão:** {versao_doc} ✅")
                else:
                    st.error("**Versão:** ❌ NÃO DETECTADA")
                if validade_doc:
                    st.info(f"**Validade:** {validade_doc}")
                else:
                    st.info("**Validade:** ⚠️ Não obrigatória")
            
            st.markdown("---")
            st.subheader("✅ SEÇÕES ENCONTRADAS")
            for s in secoes_encontradas:
                st.success(f"✅ {s}")
            
            if secoes_faltantes:
                st.subheader("❌ SEÇÕES FALTANTES")
                for s in secoes_faltantes:
                    st.error(f"❌ {s}")
            else:
                st.subheader("✅ TODAS AS SEÇÕES FORAM ENCONTRADAS!")
            
            # ============================================================
            # ✅ PROCESSAMENTO FINAL E DOWNLOADS
            # ============================================================
            dados_finais = injetar_margens_via_xml_puro(dados_brutos)
            documento_aprovado = (len(secoes_faltantes) == 0 and 
                                  versao_doc != "NÃO DETECTADA" and 
                                  codigo_doc != "NÃO DETECTADO")
            ficha_naqh_bytes = gerar_ficha_naqh(
                tipo_detectado, codigo_doc, versao_doc, 
                secoes_encontradas, secoes_faltantes, documento_aprovado
            )
            
            st.markdown("---")
            if documento_aprovado:
                st.success("🎉 **DOCUMENTO APROVADO COM SUCESSO!** Tudo pronto para download.")
                st.balloons()
            else:
                st.warning("⚠️ **DOCUMENTO FORMATADO COM PENDÊNCIAS!** Verifique as seções ou metadados ausentes.")
                
            st.markdown("### 📥 ÁREA DE DOWNLOADS DO PROCESSO")
            
            # BOTÃO DE DOWNLOAD — DOCUMENTO FORMATADO
            nome_arquivo_doc = f"{codigo_doc}_Formatado.docx" if codigo_doc != "NÃO DETECTADO" else "Documento_Formatado.docx"
            st.download_button(
                label="📥 DOWNLOAD DO DOCUMENTO FORMATADO (.DOCX)", 
                data=dados_finais, 
                file_name=nome_arquivo_doc, 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # BOTÃO DE DOWNLOAD — FICHA DE VERIFICAÇÃO
            nome_arquivo_txt = f"Ficha_Verificacao_{codigo_doc}.txt" if codigo_doc != "NÃO DETECTADO" else "Ficha_Verificacao.txt"
            st.download_button(
                label="📄 DOWNLOAD DA FICHA DE VERIFICAÇÃO (.TXT)", 
                data=ficha_naqh_bytes, 
                file_name=nome_arquivo_txt, 
                mime="text/plain"
            )
        
        except Exception as e:
            st.error(f"## ❌ ERRO durante o processamento: {str(e)}")
            st.info("Verifique se o arquivo está no formato .docx válido.")
