import streamlit as st
import docx
import re
import unicodedata
from io import BytesIO

st.set_page_config(page_title="AUDITORIA — Detecção + Manual", page_icon="🔍", layout="wide")

st.title("🔍 AUDITORIA — Detecção Automática + Confirmação Manual")
st.markdown("### ✅ Sistema tenta detectar → Você confirma/corrige manualmente se precisar")

# ============================================================
# 🧹 REMOVER ACENTOS
# ============================================================
def limpar_texto(texto):
    if not texto:
        return ""
    sem_acento = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return re.sub(r'\s+', ' ', sem_acento.upper().strip())

# ============================================================
# 📋 SEÇÕES — EXATAMENTE CONFORME O SEU QUADRO OFICIAL
# ============================================================
SECOES_POR_TIPO = {
    "PROT": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. REFERENCIAL TEORICO",
        "4. DESCRICAO DO PROTOCOLO",
        "5. ESTRATEGIAS DE MONITORAMENTO",
        "6. REFERENCIAS",
        "7. APENDICES",
        "8. ANEXOS"
    ],
    "POP": [
        "1. DEFINICAO",
        "2. APLICABILIDADE",
        "3. RESPONSAVEL PELA EXECUCAO",
        "4. MATERIAIS UTILIZADOS",
        "5. DESCRICAO DA TAREFA",
        "6. ATIVIDADES",
        "7. REFERENCIAS",
        "8. ANEXOS"
    ],
    "POI": [
        "1. INTRODUCAO",
        "2. OBJETIVO",
        "3. PRINCIPIOS",
        "4. DIRETRIZES",
        "5. RESPONSABILIDADES",
        "6. ESTRATEGIA DE MONITORAMENTO",
        "7. REFERENCIAS",
        "8. APENDICES E ANEXOS"
    ],
    "NOR": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. DESCRICAO DA NORMA",
        "4. RESPONSAVEL",
        "5. EFEITOS NO CUMPRIMENTO",
        "6. NORMA DE REFERENCIA",
        "7. ANEXOS"
    ],
    "REG": [
        "1. DA FINALIDADE",
        "2. DA COMPOSICAO — MEMBROS",
        "3. DO MANDATO",
        "4. DO FUNCIONAMENTO E ORGANIZACAO",
        "5. DAS ATRIBUICOES",
        "6. DISPOSICOES FINAIS"
    ],
    "PROG": [
        "1. REFERENCIAL TEORICO",
        "2. PADRONIZACAO DE ROTINAS",
        "3. ESTRATEGIAS DE MONITORAMENTO",
        "4. DESCRICAO DO PROGRAMA",
        "5. MEDIDAS EDUCACIONAIS",
        "6. REFERENCIAS",
        "7. APENDICES",
        "8. ANEXOS"
    ],
    "PLAN": [
        "1. OBJETIVO",
        "2. APLICABILIDADE",
        "3. DEFINICAO DE TERMOS",
        "4. IDENTIFICACAO DE RISCO ATUAL",
        "5. MEDIDAS DE CONTINGENCIA",
        "6. REFERENCIAS",
        "7. APENDICES",
        "8. ANEXOS"
    ],
    "ROT": [
        "1. DEFINICAO",
        "2. OBJETIVO",
        "3. APLICABILIDADE",
        "4. DESCRICAO DA ROTINA",
        "5. APENDICES"
    ],
    "MAN": [
        "1. CAPA",
        "2. ELABORADORES",
        "3. COLABORADORES",
        "4. SUMARIO",
        "5. APRESENTACAO",
        "6. DESCRICAO",
        "7. REFERENCIAS",
        "8. APENDICES E ANEXOS"
    ]
}

# ============================================================
# 🧹 APLICAR MARGENS ABNT (3/3/2/2)
# ============================================================
def injetar_margens_via_xml_puro(doc_bytes):
    doc = docx.Document(BytesIO(doc_bytes))
    for sec in doc.sections:
        sec.top_margin = docx.shared.Cm(3)
        sec.bottom_margin = docx.shared.Cm(2)
        sec.left_margin = docx.shared.Cm(3)
        sec.right_margin = docx.shared.Cm(2)
    output = BytesIO()
    doc.save(output)
    return output.getvalue()

# ============================================================
# 📄 GERAR FICHA
# ============================================================
def gerar_ficha_naqh(tipo, codigo, versao, encontradas, faltantes, aprovado):
    ficha = [
        "=" * 60,
        "FICHA DE VERIFICAÇÃO DE DOCUMENTO",
        "=" * 60,
        f"Tipo de Documento: {tipo}",
        f"Código: {codigo}",
        f"Versão: {versao}",
        "",
        "--- SEÇÕES ENCONTRADAS ---"
    ]
    for s in encontradas:
        ficha.append(f"✓ {s}")
    ficha.append("")
    if faltantes:
        ficha.append("--- SEÇÕES FALTANTES ---")
        for s in faltantes:
            ficha.append(f"✗ {s}")
    else:
        ficha.append("--- TODAS AS SEÇÕES ENCONTRADAS ---")
    ficha.extend(["", "=" * 60, f"STATUS: {'APROVADO' if aprovado else 'COM PENDÊNCIAS'}", "=" * 60])
    return "\n".join(ficha).encode("utf-8")

# ============================================================
# 🧠 FUNÇÃO DE DETECÇÃO AUTOMÁTICA
# ============================================================
def auditar_documento(arquivo_bytes):
    doc = docx.Document(BytesIO(arquivo_bytes))
    
    codigo_detectado = None
    versao_detectada = None
    validade_detectada = None
    texto_completo = ""
    
    # 🔍 LER TABELAS — ONDE ESTÁ O CABEÇALHO!
    for tabela in doc.tables:
        for linha in tabela.rows:
            linha_texto = " ".join([cel.text for cel in linha.cells])
            texto_completo += linha_texto + " "
            
            # ✅ CÓDIGO — EXATO: "Código: PROT_SCHI005"
            cod_match = re.search(r'CÓDIGO|Código[:\s]*[:]?\s*([A-Z]{2,5}[_\s][A-Z0-9]+)', linha_texto, re.IGNORECASE)
            if cod_match and cod_match.group(1):
                codigo_detectado = re.sub(r'\s+', '_', cod_match.group(1).strip())
            
            # ✅ VERSÃO — EXATO: "Versão: 5ª" → pega só o número
            ver_match = re.search(r'VERSÃO|Versão[:\s]*[:]?\s*(\d+)', linha_texto, re.IGNORECASE)
            if ver_match and ver_match.group(1):
                versao_detectada = ver_match.group(1).strip()
            
            # ✅ VALIDADE
            val_match = re.search(r'VALIDADE|Validade[:\s]*[:]?\s*([\d/]+)', linha_texto, re.IGNORECASE)
            if val_match and val_match.group(1):
                validade_detectada = val_match.group(1).strip()
    
    # ✅ LER CORPO DO DOCUMENTO
    for p in doc.paragraphs:
        texto_completo += p.text + " "
    
    texto_limpo = limpar_texto(texto_completo)
    
    # 🔍 DETECTAR TIPO
    tipo_detectado = "PROT"
    if re.search(r'\bPROTOCOLO\b', texto_limpo):
        tipo_detectado = "PROT"
    elif re.search(r'\bPOP\b', texto_limpo):
        tipo_detectado = "POP"
    elif re.search(r'\bPOLÍTICA|POLITICA|POI\b', texto_limpo):
        tipo_detectado = "POI"
    elif re.search(r'\bNORMA|NOR\b', texto_limpo):
        tipo_detectado = "NOR"
    elif re.search(r'\bREGIMENTO|REG\b', texto_limpo):
        tipo_detectado = "REG"
    elif re.search(r'\bPROGRAMA|PROG\b', texto_limpo):
        tipo_detectado = "PROG"
    elif re.search(r'\bPLANO|PLAN\b', texto_limpo):
        tipo_detectado = "PLAN"
    elif re.search(r'\bROTINA|ROT\b', texto_limpo):
        tipo_detectado = "ROT"
    elif re.search(r'\bMANUAL|MAN\b', texto_limpo):
        tipo_detectado = "MAN"
    
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
        "codigo": codigo_detectado,
        "versao": versao_detectada,
        "validade": validade_detectada,
        "secoes_encontradas": secoes_encontradas,
        "secoes_faltantes": secoes_faltantes
    }

# ============================================================
# 🚀 INTERFACE PRINCIPAL
# ============================================================
with st.form("auditoria_com_confirmacao_manual"):
    arquivo_word = st.file_uploader(
        "📂 Arraste o documento WORD (.docx) AQUI",
        type=["docx"]
    )
    enviado = st.form_submit_button("🔍 EXECUTAR DETECÇÃO", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Arquivo carregado: **{arquivo_word.name}**")
    
    with st.spinner("Detectando informações..."):
        try:
            dados_brutos = arquivo_word.read()
            rel = auditar_documento(dados_brutos)
            
            # ============================================================
            # ✅ ÁREA DE CONFIRMAÇÃO / CORREÇÃO MANUAL
            # ============================================================
            st.markdown("---")
            st.subheader("📋 INFORMAÇÕES DETECTADAS — CONFIRME OU CORRIJA ABAIXO")
            
            col1, col2 = st.columns(2)
            with col1:
                tipo_detectado = st.selectbox(
                    "**Tipo de Documento**",
                    options=list(SECOES_POR_TIPO.keys()),
                    index=list(SECOES_POR_TIPO.keys()).index(rel["tipo"])
                )
                codigo_doc = st.text_input(
                    "**Código do Documento**",
                    value=rel["codigo"] or "",
                    placeholder="Ex: PROT_SCHI005"
                )
            with col2:
                versao_doc = st.text_input(
                    "**Versão**",
                    value=rel["versao"] or "",
                    placeholder="Ex: 5"
                )
                validade_doc = st.text_input(
                    "**Validade**",
                    value=rel["validade"] or "",
                    placeholder="Ex: xx/xx/2030 (opcional)"
                )
            
            # ============================================================
            # ✅ SEÇÕES — MARQUE MANUALMENTE AS ENCONTRADAS
            # ============================================================
            st.markdown("---")
            st.subheader("✅ SEÇÕES — CONFIRME QUAIS ESTÃO PRESENTES")
            
            secoes_esperadas = SECOES_POR_TIPO[tipo_detectado]
            secoes_encontradas = []
            secoes_faltantes = []
            
            st.markdown("Marque as seções que existem no documento:")
            for secao in secoes_esperadas:
                # Verifica se foi detectada automaticamente
                detectada_auto = secao in rel["secoes_encontradas"]
                # Checkbox para confirmação manual
                confirmada = st.checkbox(
                    secao,
                    value=detectada_auto,
                    key=f"secao_{secao}"
                )
                if confirmada:
                    secoes_encontradas.append(secao)
                else:
                    secoes_faltantes.append(secao)
            
            # ============================================================
            # ✅ PROCESSAMENTO FINAL E DOWNLOADS
            # ============================================================
            st.markdown("---")
            if not st.form_submit_button("✅ CONFIRMAR E GERAR RESULTADO", type="primary"):
                st.info("👆 Clique em CONFIRMAR acima após revisar os dados")
                st.stop()
            
            # Garante valores padrão
            codigo_doc = codigo_doc or "NÃO DETECTADO"
            versao_doc = versao_doc or "NÃO DETECTADA"
            
            dados_finais = injetar_margens_via_xml_puro(dados_brutos)
            documento_aprovado = (len(secoes_faltantes) == 0 and 
                                  versao_doc != "NÃO DETECTADA" and 
                                  codigo_doc != "NÃO DETECTADO")
            ficha_bytes = gerar_ficha_naqh(
                tipo_detectado, codigo_doc, versao_doc, 
                secoes_encontradas, secoes_faltantes, documento_aprovado
            )
            
            # ============================================================
            # 📊 RESULTADO FINAL
            # ============================================================
            st.markdown("---")
            st.subheader("📋 RELATÓRIO FINAL CONFIRMADO")
            
            c1, c2 = st.columns(2)
            with c1:
                st.info(f"**Tipo:** {tipo_detectado}")
                if codigo_doc != "NÃO DETECTADO":
                    st.success(f"**Código:** {codigo_doc} ✅")
                else:
                    st.error("**Código:** ❌ NÃO INFORMADO")
            with c2:
                if versao_doc != "NÃO DETECTADA":
                    st.success(f"**Versão:** {versao_doc} ✅")
                else:
                    st.error("**Versão:** ❌ NÃO INFORMADA")
                if validade_doc:
                    st.info(f"**Validade:** {validade_doc}")
            
            st.markdown("---")
            st.subheader("✅ SEÇÕES CONFIRMADAS")
            for s in secoes_encontradas:
                st.success(f"✅ {s}")
            
            if secoes_faltantes:
                st.subheader("❌ SEÇÕES FALTANTES")
                for s in secoes_faltantes:
                    st.error(f"❌ {s}")
            else:
                st.subheader("✅ TODAS AS SEÇÕES CONFIRMADAS!")
            
            st.markdown("---")
            if documento_aprovado:
                st.success("🎉 **DOCUMENTO APROVADO COM SUCESSO!**")
                st.balloons()
            else:
                st.warning("⚠️ **DOCUMENTO COM PENDÊNCIAS!** Verifique os dados.")
            
            # ============================================================
            # 📥 DOWNLOADS
            # ============================================================
            st.markdown("### 📥 ÁREA DE DOWNLOADS")
            nome_doc = f"{codigo_doc}_Formatado.docx" if codigo_doc != "NÃO DETECTADO" else "Documento_Formatado.docx"
            nome_txt = f"Ficha_Verificacao_{codigo_doc}.txt" if codigo_doc != "NÃO DETECTADO" else "Ficha_Verificacao.txt"
            
            st.download_button(
                "📥 DOWNLOAD DO DOCUMENTO FORMATADO (.DOCX)",
                dados_finais, nome_doc,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.download_button(
                "📄 DOWNLOAD DA FICHA DE VERIFICAÇÃO (.TXT)",
                ficha_bytes, nome_txt, "text/plain"
            )
        
        except Exception as e:
            st.error(f"## ❌ ERRO: {str(e)}")
