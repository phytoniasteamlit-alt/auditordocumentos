import streamlit as st
import docx
import re
import unicodedata
from io import BytesIO

# ============================================================
# 🎯 CONFIGURAÇÕES
# ============================================================
st.set_page_config(page_title="AUDITOR NAQH NMZ", page_icon="👨‍💻", layout="wide")

TEMPO_POR_DOC_MANUAL = 20
TEMPO_POR_DOC_AUTOMATICO = 1

MARGEM_SUP_ESPERADA = 3.0
MARGEM_INF_ESPERADA = 2.0
MARGEM_ESQ_ESPERADA = 3.0
MARGEM_DIR_ESPERADA = 2.0

FONTE_CORPO = "Calibri"
TAMANHO_CORPO = 11
FONTE_TABELAS = "Calibri"
TAMANHO_TABELAS = 10

# ✅ CORRIGIDO: 1 cm = 567 twips (valor EXATO)
TWIPS_PARA_CM = 567.0

# ============================================================
# 📋 SEÇÕES POR TIPO DE DOCUMENTO
# ============================================================
SECOES_POR_TIPO = {
    "PROT": ["1. OBJETIVO", "2. APLICABILIDADE", "3. REFERENCIAL TEÓRICO",
             "4. DESCRIÇÃO DO PROTOCOLO", "5. ESTRATÉGIAS DE MONITORAMENTO",
             "6. REFERÊNCIAS", "7. APÊNDICES", "8. ANEXOS"],
    "POP": ["1. DEFINIÇÃO", "2. APLICABILIDADE", "3. RESPONSÁVEL PELA EXECUÇÃO",
             "4. MATERIAIS UTILIZADOS", "5. DESCRIÇÃO DA TAREFA",
             "6. ATIVIDADES", "7. REFERÊNCIAS", "8. ANEXOS"],
    "POI": ["1. INTRODUÇÃO", "2. OBJETIVO", "3. PRINCÍPIOS", "4. DIRETRIZES",
            "5. RESPONSABILIDADES", "6. ESTRATÉGIA DE MONITORAMENTO",
            "7. REFERÊNCIAS", "8. APÊNDICES E ANEXOS"],
    "NOR": ["1. OBJETIVO", "2. APLICABILIDADE", "3. DESCRIÇÃO DA NORMA",
            "4. RESPONSÁVEL", "5. EFEITOS NO CUMPRIMENTO",
            "6. NORMA DE REFERÊNCIA", "7. ANEXOS"],
    "REG": ["1. DA FINALIDADE", "2. DA COMPOSIÇÃO — MEMBROS", "3. DO MANDATO",
            "4. DO FUNCIONAMENTO E ORGANIZAÇÃO", "5. DAS ATRIBUIÇÕES",
            "6. DISPOSIÇÕES FINAIS"],
    "PROG": ["1. REFERENCIAL TEÓRICO", "2. PADRONIZAÇÃO DE ROTINAS",
             "3. ESTRATÉGIAS DE MONITORAMENTO", "4. DESCRIÇÃO DO PROGRAMA",
             "5. MEDIDAS EDUCACIONAIS", "6. REFERÊNCIAS",
             "7. APÊNDICES", "8. ANEXOS"],
    "PLAN": ["1. OBJETIVO", "2. APLICABILIDADE", "3. DEFINIÇÃO DE TERMOS",
             "4. IDENTIFICAÇÃO DE RISCO ATUAL", "5. MEDIDAS DE CONTINGÊNCIA",
             "6. REFERÊNCIAS", "7. APÊNDICES", "8. ANEXOS"],
    "ROT": ["1. DEFINIÇÃO", "2. OBJETIVO", "3. APLICABILIDADE",
            "4. DESCRIÇÃO DA ROTINA", "5. APÊNDICES"],
    "MAN": ["1. CAPA", "2. ELABORADORES", "3. COLABORADORES", "4. SUMÁRIO",
            "5. APRESENTAÇÃO", "6. DESCRIÇÃO", "7. REFERÊNCIAS",
            "8. APÊNDICES E ANEXOS"]
}

# ============================================================
# 🧹 FUNÇÕES AUXILIARES — MAIS FLEXÍVEIS
# ============================================================
def limpar_texto(texto):
    if not texto: return ""
    # Remove acentos
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII','ignore').decode('ASCII')
    # Remove pontos, espaços múltiplos, tabulações e deixa tudo maiúsculo
    texto = re.sub(r'[\s.\t]+', ' ', texto).upper().strip()
    # Remove o número inicial da seção (ex: "1. OBJETIVO" → "OBJETIVO")
    texto = re.sub(r'^\d+\s*', '', texto).strip()
    return texto

def formatar_tempo(minutos_total):
    h = int(minutos_total // 60)
    m = int(minutos_total % 60)
    if h > 0: return f"{h}h {m}min" if m>0 else f"{h}h"
    return f"{m}min"

# ============================================================
# 📏 VERIFICAR MARGENS — ✅ CORRIGIDO O ERRO PRINCIPAL!
# ============================================================
def verificar_margens(doc):
    sec = doc.sections[0]
    
    def cm_de_twips(twips_valor):
        if twips_valor is None or twips_valor == 0:
            return 0.0
        # ✅ DIVISÃO CORRETA twips / 567 = cm
        valor_cm = twips_valor / TWIPS_PARA_CM
        return round(valor_cm, 2)
    
    m_sup = cm_de_twips(sec.top_margin)
    m_inf = cm_de_twips(sec.bottom_margin)
    m_esq = cm_de_twips(sec.left_margin)
    m_dir = cm_de_twips(sec.right_margin)
    
    tol = 0.15  # Margem de tolerância maior para não errar por pouca coisa
    return {
        "sup": m_sup, "inf": m_inf, "esq": m_esq, "dir": m_dir,
        "ok_sup": abs(m_sup - MARGEM_SUP_ESPERADA) < tol,
        "ok_inf": abs(m_inf - MARGEM_INF_ESPERADA) < tol,
        "ok_esq": abs(m_esq - MARGEM_ESQ_ESPERADA) < tol,
        "ok_dir": abs(m_dir - MARGEM_DIR_ESPERADA) < tol,
        "todas_ok": (abs(m_sup - MARGEM_SUP_ESPERADA) < tol and
                     abs(m_inf - MARGEM_INF_ESPERADA) < tol and
                     abs(m_esq - MARGEM_ESQ_ESPERADA) < tol and
                     abs(m_dir - MARGEM_DIR_ESPERADA) < tol)
    }

# ============================================================
# ✍️ VERIFICAR FONTE
# ============================================================
def verificar_fonte(doc):
    cont_corpo = {"total":0, "fonte_ok":0, "tam_ok":0, "fontes":{}, "tams":{}}
    cont_tab = {"total":0, "fonte_ok":0, "tam_ok":0, "fontes":{}, "tams":{}}
    tem_registro_historico = False

    for p in doc.paragraphs:
        for run in p.runs:
            cont_corpo["total"] += 1
            nome = run.font.name or FONTE_CORPO
            tam_pt = round(run.font.size.pt,1) if run.font.size else TAMANHO_CORPO
            cont_corpo["fontes"][nome] = cont_corpo["fontes"].get(nome,0)+1
            cont_corpo["tams"][tam_pt] = cont_corpo["tams"].get(tam_pt,0)+1
            if nome == FONTE_CORPO: cont_corpo["fonte_ok"] += 1
            if tam_pt == TAMANHO_CORPO: cont_corpo["tam_ok"] += 1

    for tb in doc.tables:
        texto_tb = ""
        for ln in tb.rows:
            texto_linha = ""
            for cel in ln.cells:
                texto_linha += cel.text + " "
                for p in cel.paragraphs:
                    for run in p.runs:
                        cont_tab["total"] += 1
                        nome = run.font.name or FONTE_TABELAS
                        tam_pt = round(run.font.size.pt,1) if run.font.size else TAMANHO_TABELAS
                        cont_tab["fontes"][nome] = cont_tab["fontes"].get(nome,0)+1
                        cont_tab["tams"][tam_pt] = cont_tab["tams"].get(tam_pt,0)+1
                        if nome == FONTE_TABELAS: cont_tab["fonte_ok"] += 1
                        if tam_pt == TAMANHO_TABELAS: cont_tab["tam_ok"] += 1
            texto_tb += texto_linha
        if "REGISTRO HISTORICO" in limpar_texto(texto_tb):
            tem_registro_historico = True

    def pct(ok, tot): return round((ok/tot*100),1) if tot else 100
    crit = 70

    return {
        "corpo": {
            "fontes": cont_corpo["fontes"], "tamanhos": cont_corpo["tams"],
            "pct_fonte": pct(cont_corpo["fonte_ok"], cont_corpo["total"]),
            "pct_tam": pct(cont_corpo["tam_ok"], cont_corpo["total"]),
            "fonte_ok": pct(cont_corpo["fonte_ok"], cont_corpo["total"]) >= crit,
            "tam_ok": pct(cont_corpo["tam_ok"], cont_corpo["total"]) >= crit,
        },
        "tabelas": {
            "fontes": cont_tab["fontes"], "tamanhos": cont_tab["tams"],
            "pct_fonte": pct(cont_tab["fonte_ok"], cont_tab["total"]),
            "pct_tam": pct(cont_tab["tam_ok"], cont_tab["total"]),
            "fonte_ok": pct(cont_tab["fonte_ok"], cont_tab["total"]) >= crit,
            "tam_ok": pct(cont_tab["tam_ok"], cont_tab["total"]) >= crit,
            "tem_registro_historico": tem_registro_historico
        }
    }

# ============================================================
# 🔍 ESCANEAR DOCUMENTO — ✅ BUSCA CÓDIGO+VERSÃO MELHORADA
# ============================================================
def escanear(doc_bytes):
    doc = docx.Document(BytesIO(doc_bytes)))
    texto_completo = ""
    codigo = versao = None

    # Tabelas (cabeçalho) — busca CÓDIGO e VERSÃO
    for tb in doc.tables:
        for ln in tb.rows:
            texto_linha = " ".join([cel.text for cel in ln.cells])
            texto_completo += texto_linha + " "
            
            # ✅ BUSCA MAIS FLEXÍVEL: aceita variações de formatação
            if not codigo:
                m_cod = re.search(r'C[ÓO]DIGO\s*[:：=\-]?\s*([A-Z0-9_\-\/\.]+)', texto_linha.upper())
                if m_cod:
                    codigo = m_cod.group(1).strip()
            
            if not versao:
                m_ver = re.search(r'VERS[AÃ]O\s*[:：=\-]?\s*(\d+(?:[.\-]\d+)*)', texto_linha.upper())
                if m_ver:
                    versao = m_ver.group(1).strip()

    # Corpo do texto
    for p in doc.paragraphs:
        texto_completo += p.text + " "

    texto_limpo = limpar_texto(texto_completo)

    # ✅ DETECTAR TIPO
    tipo = None
    if re.search(r'\bPROTOCOLO\b', texto_limpo): tipo = "PROT"
    elif re.search(r'\bPOP\b|\bPROCEDIMENTO OPERACIONAL\b', texto_limpo): tipo = "POP"
    elif re.search(r'\bPOL[IÍ]TICA\b|\bPOI\b', texto_limpo): tipo = "POI"
    elif re.search(r'\bNORMA\b|\bNOR\b', texto_limpo): tipo = "NOR"
    elif re.search(r'\bREGIMENTO\b|\bREGULAMENTO\b|\bREG\b', texto_limpo): tipo = "REG"
    elif re.search(r'\bPROGRAMA\b|\bPROG\b', texto_limpo): tipo = "PROG"
    elif re.search(r'\bPLANO\b|\bPLAN\b', texto_limpo): tipo = "PLAN"
    elif re.search(r'\bROTINA\b|\bROT\b', texto_limpo): tipo = "ROT"
    elif re.search(r'\bMANUAL\b|\bMAN\b', texto_limpo): tipo = "MAN"
    else: tipo = "PROT"

    # ✅ SEÇÕES — COMPARAÇÃO FLEXÍVEL (ignora número e formatação)
    secoes_esperadas = SECOES_POR_TIPO[tipo]
    encontradas = []
    faltantes = []
    
    for secao in secoes_esperadas:
        secao_limpa = limpar_texto(secao)
        # Procura a seção APENAS pelo nome (sem o número)
        if secao_limpa in texto_limpo:
            encontradas.append(secao)
        else:
            faltantes.append(secao)

    return {
        "tipo": tipo, "codigo": codigo, "versao": versao,
        "margens": verificar_margens(doc),
        "fonte": verificar_fonte(doc),
        "secoes_esperadas": secoes_esperadas,
        "secoes_enc": encontradas, "secoes_falt": faltantes
    }

# ============================================================
# 🧹 APLICAR MARGENS
# ============================================================
def aplicar_margens(doc_bytes):
    doc = docx.Document(BytesIO(doc_bytes))
    for sec in doc.sections:
        sec.top_margin = docx.shared.Cm(MARGEM_SUP_ESPERADA)
        sec.bottom_margin = docx.shared.Cm(MARGEM_INF_ESPERADA)
        sec.left_margin = docx.shared.Cm(MARGEM_ESQ_ESPERADA)
        sec.right_margin = docx.shared.Cm(MARGEM_DIR_ESPERADA)
    saida = BytesIO()
    doc.save(saida)
    return saida.getvalue()

# ============================================================
# 🚀 INTERFACE PRINCIPAL
# ============================================================
st.markdown("""
    <style>
    .header-container{display:flex;justify-content:space-between;align-items:center;background:#0F172A;padding:15px 25px;border-radius:12px;margin-bottom:20px}
    .header-text h1{color:#FFF;margin:0;font-size:28px}
    .header-text p{color:#94A3B8;margin:5px 0 0 0}
    .header-emoji{font-size:50px}
    .relogio-box{background:linear-gradient(135deg,#0F766E,#14B8A6);padding:20px 25px;border-radius:12px;color:#FFF;margin:15px 0}
    .economia-grande{font-size:32px;font-weight:bold}
    .base-tempo{background:#1E293B;padding:12px 18px;border-radius:8px;color:#CBD5E1;font-size:13px;margin-top:8px}
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="header-container">
        <div class="header-text">
            <h1>AUDITOR NAQH NMZ DE ALTA PRECISÃO</h1>
            <p>Ezequias Santos — Agente Administrativo | Margens: Esq 3,0 / Dir 2,0 / Sup 3,0 / Inf 2,0 cm • Calibri 11 (corpo) / 10 (tabelas)</p>
        </div>
        <div class="header-emoji">👨‍💻</div>
    </div>
""", unsafe_allow_html=True)

arquivos = st.file_uploader("📂 Envie o(s) documento(s) (.docx)", type=["docx"], accept_multiple_files=True)

if arquivos:
    qtd = len(arquivos)
    tempo_manual = qtd * TEMPO_POR_DOC_MANUAL
    tempo_auditor = qtd * TEMPO_POR_DOC_AUTOMATICO
    tempo_economizado = tempo_manual - tempo_auditor

    st.markdown(f"""
    <div class="relogio-box">
        <div style="display:flex;justify-content:space-between;align-items:center;">
            <div><h3 style="margin:0">⏱️ ECONOMIA DE TEMPO</h3><p style="margin:5px 0 0 0;opacity:0.9;">{qtd} documento(s)</p></div>
            <div class="economia-grande">{formatar_tempo(tempo_economizado)}</div>
        </div>
        <div style="display:flex;gap:30px;margin-top:12px;font-size:14px;">
            <span>📝 Manual: <b>{formatar_tempo(tempo_manual)}</b></span>
            <span>⚡ Auditor: <b>{formatar_tempo(tempo_auditor)}</b></span>
            <span>✅ Redução: <b>{round((tempo_economizado/tempo_manual)*100)}%</b></span>
        </div>
    </div>
    <div class="base-tempo">
        📌 Base: ~{TEMPO_POR_DOC_MANUAL}min/doc manualmente vs ~{TEMPO_POR_DOC_AUTOMATICO}min com o Auditor
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    for idx, arq in enumerate(arquivos, 1):
        st.subheader(f"📄 {arq.name}")
        try:
            dados = arq.read()
            r = escanear(dados)

            st.markdown("### 📋 DADOS DO DOCUMENTO")
            c1,c2 = st.columns(2)
            with c1: st.info(f"**Tipo Detectado:** {r['tipo']}")
            with c2:
                if r['codigo']: st.success(f"**Código:** {r['codigo']}")
                else: st.error("**Código:** ❌ NÃO ENCONTRADO")
            if r['versao']: st.success(f"**Versão:** {r['versao']}")
            else: st.error("**Versão:** ❌ NÃO ENCONTRADA")

            st.markdown("---")
            st.markdown("### 📏 MARGENS — Esperado: Sup=3,0 / Inf=2,0 / Esq=3,0 / Dir=2,0 cm")
            m = r["margens"]
            col1,col2,col3,col4 = st.columns(4)
            col1.metric("Superior", f"{m['sup']} cm", "✅" if m['ok_sup'] else "❌")
            col2.metric("Inferior", f"{m['inf']} cm", "✅" if m['ok_inf'] else "❌")
            col3.metric("Esquerda", f"{m['esq']} cm", "✅" if m['ok_esq'] else "❌")
            col4.metric("Direita", f"{m['dir']} cm", "✅" if m['ok_dir'] else "❌")
            st.success("✅ TODAS AS MARGENS CONFORME") if m["todas_ok"] else st.error("❌ MARGENS NÃO CONFORME")

            st.markdown("---")
            st.markdown("### ✍️ CORPO — Calibri 11pt")
            f = r["fonte"]["corpo"]
            st.write("Fontes encontradas:", ", ".join(f"{n} ({q})" for n,q in f["fontes"].items()))
            st.write("Tamanhos encontrados:", ", ".join(f"{t}pt ({q})" for t,q in f["tamanhos"].items()))
            st.metric("Conformidade", f"{f['pct_fonte']}%", "✅" if f["fonte_ok"] else "❌")

            st.markdown("---")
            st.markdown("### ✍️ TABELAS — Calibri 10pt")
            ft = r["fonte"]["tabelas"]
            if ft["fontes"]:
                st.write("Fontes encontradas:", ", ".join(f"{n} ({q})" for n,q in ft["fontes"].items()))
                st.write("Tamanhos encontrados:", ", ".join(f"{t}pt ({q})" for t,q in ft["tamanhos"].items()))
                st.metric("Conformidade", f"{ft['pct_fonte']}%", "✅" if ft["fonte_ok"] else "❌")
            if ft["tem_registro_historico"]:
                st.info("📋 Registro Histórico detectado")

            st.markdown("---")
            st.markdown(f"### ✅ CONFERÊNCIA DE SEÇÕES — Tipo: {r['tipo']}")
            st.info("💡 Marque manualmente o que encontrou no documento. O sistema já verificou.")

            secoes_usuario = []
            for secao in r["secoes_esperadas"]:
                sistema_encontrou = secao in r["secoes_enc"]
                marcado = st.checkbox(
                    f"{secao} {'✅ (DETECTADO)' if sistema_encontrou else '⚠️ NÃO DETECTADO'}",
                    value=sistema_encontrou,
                    key=f"chk_{idx}_{limpar_texto(secao)}"
                )
                secoes_usuario.append((secao, marcado, sistema_encontrou))

            st.markdown("#### 📊 RESULTADO DA CONFERÊNCIA")
            for secao, marcado, sistema_encontrou in secoes_usuario:
                if marcado and sistema_encontrou:
                    st.success(f"✅ {secao} — CONFIRMADO")
                elif marcado and not sistema_encontrou:
                    st.warning(f"⚠️ {secao} — VOCÊ ENCONTROU, SISTEMA NÃO DETECTOU!")
                elif not marcado and sistema_encontrou:
                    st.info(f"ℹ️ {secao} — Sistema detectou, você NÃO MARCOU")
                else:
                    st.error(f"❌ {secao} — NÃO ENCONTRADO")

            st.markdown("---")
            aprov = m["todas_ok"] and r["codigo"] and r["versao"] and len(r["secoes_falt"])==0
            st.success("✅ DOCUMENTO APROVADO CONFORME NORMA ZERO") if aprov else st.warning("⚠️ DOCUMENTO COM PENDÊNCIAS")

            dados_format = aplicar_margens(dados)
            nome_arq = f"{r['codigo'] or 'DOC'}_Formatado.docx"
            st.download_button("📥 Baixar formatado", dados_format, nome_arq,
                              "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl{idx}")

        except Exception as e:
            st.error(f"❌ ERRO: {str(e)}")

        st.markdown("---")
