import streamlit as st
import docx
import re
import unicodedata
from io import BytesIO
from docx.shared import Cm

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

LIMIAR_CONFORMIDADE = 70

# ============================================================
# 📋 SEÇÕES POR TIPO
# ============================================================
SECOES_POR_TIPO = {
    "PROT": {
        "obrigatorias": [
            "OBJETIVO", "APLICAÇÃO", "REFERENCIAL TEÓRICO",
            "CLASSIFICAÇÃO DAS CIRURGIAS POR POTENCIAL DE CONTAMINAÇÃO",
            "FATORES DE RISCO", "BENEFÍCIOS E RISCOS", "PRINCÍPIOS GERAIS",
            "CRITÉRIOS DE ELEGIBILIDADE", "ESQUEMA PADRÃO",
            "ESQUEMAS POR TIPO DE CIRURGIA"
        ],
        "opcionais": ["APÊNDICES", "ANEXOS"]
    },
    "POP": {
        "obrigatorias": [
            "DEFINIÇÃO", "APLICABILIDADE", "RESPONSÁVEL PELA EXECUÇÃO",
            "MATERIAIS UTILIZADOS", "DESCRIÇÃO DA TAREFA",
            "ATIVIDADES", "REFERÊNCIAS"
        ],
        "opcionais": ["ANEXOS"]
    },
    "POI": {
        "obrigatorias": [
            "INTRODUÇÃO", "OBJETIVO", "PRINCÍPIOS", "DIRETRIZES",
            "RESPONSABILIDADES", "ESTRATÉGIA DE MONITORAMENTO", "REFERÊNCIAS"
        ],
        "opcionais": ["APÊNDICES E ANEXOS"]
    },
    "NOR": {
        "obrigatorias": [
            "OBJETIVO", "APLICABILIDADE", "DESCRIÇÃO DA NORMA",
            "RESPONSÁVEL", "EFEITOS NO CUMPRIMENTO", "NORMA DE REFERÊNCIA"
        ],
        "opcionais": ["ANEXOS"]
    },
    "REG": {
        "obrigatorias": [
            "DA FINALIDADE", "DA COMPOSIÇÃO — MEMBROS", "DO MANDATO",
            "DO FUNCIONAMENTO E ORGANIZAÇÃO", "DAS ATRIBUIÇÕES", "DISPOSIÇÕES FINAIS"
        ],
        "opcionais": []
    },
    "PROG": {
        "obrigatorias": [
            "REFERENCIAL TEÓRICO", "PADRONIZAÇÃO DE ROTINAS",
            "ESTRATÉGIAS DE MONITORAMENTO", "DESCRIÇÃO DO PROGRAMA",
            "MEDIDAS EDUCACIONAIS", "REFERÊNCIAS"
        ],
        "opcionais": ["APÊNDICES", "ANEXOS"]
    },
    "PLAN": {
        "obrigatorias": [
            "OBJETIVO", "APLICABILIDADE", "DEFINIÇÃO DE TERMOS",
            "IDENTIFICAÇÃO DE RISCO ATUAL", "MEDIDAS DE CONTINGÊNCIA", "REFERÊNCIAS"
        ],
        "opcionais": ["APÊNDICES", "ANEXOS"]
    },
    "ROT": {
        "obrigatorias": [
            "DEFINIÇÃO", "OBJETIVO", "APLICABILIDADE", "DESCRIÇÃO DA ROTINA"
        ],
        "opcionais": ["APÊNDICES"]
    },
    "MAN": {
        "obrigatorias": [
            "CAPA", "ELABORADORES", "COLABORADORES", "SUMÁRIO",
            "APRESENTAÇÃO", "DESCRIÇÃO", "REFERÊNCIAS"
        ],
        "opcionais": ["APÊNDICES E ANEXOS"]
    }
}

# ============================================================
# 🧹 FUNÇÕES AUXILIARES
# ============================================================
def limpar_texto(texto):
    if not texto: return ""
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    texto = re.sub(r'[\s.\-_\t]+', ' ', texto).upper().strip()
    texto = re.sub(r'^\d+\s*', '', texto).strip()
    return texto

def formatar_tempo(minutos_total):
    try: minutos_total = int(minutos_total)
    except: minutos_total = 0
    h = minutos_total // 60
    m = minutos_total % 60
    if h > 0: return f"{h}h {m}min" if m > 0 else f"{h}h"
    return f"{m}min"

def calcular_pct(ok, total):
    return round((ok / total * 100), 1) if total > 0 else 100

def detectar_pelo_nome(nome_arquivo):
    """Detecta tipo pelo nome do arquivo — MAIOR PRIORIDADE"""
    nl = limpar_texto(nome_arquivo)
    mapeia_nome = [
        ("POP", r'\bPOP\b'),
        ("POI", r'\bPOI\b|\bPOLITICA\b'),
        ("NOR", r'\bNOR\b|\bNORMA\b'),
        ("PLAN", r'\bPLAN\b|\bPLANO\b'),
        ("PROT", r'\bPROT\b|\bPROTOCOLO\b'),
        ("REG", r'\bREG\b|\bREGIMENTO\b'),
        ("PROG", r'\bPROG\b|\bPROGRAMA\b'),
        ("ROT", r'\bROT\b|\bROTINA\b'),
        ("MAN", r'\bMAN\b|\bMANUAL\b'),
    ]
    for t, p in mapeia_nome:
        if re.search(p, nl):
            return t
    return None

# ============================================================
# ✍️ VERIFICAR FONTE
# ============================================================
def verificar_fonte(doc):
    cont_corpo = {"total":0, "fonte_ok":0, "tam_ok":0, "fontes":{}, "tams":{}}
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
        for ln in tb.rows:
            for cel in ln.cells:
                for p in cel.paragraphs:
                    for run in p.runs:
                        cont_corpo["total"] += 1
                        nome = run.font.name or FONTE_TABELAS
                        tam_pt = round(run.font.size.pt,1) if run.font.size else TAMANHO_TABELAS
                        cont_corpo["fontes"][nome] = cont_corpo["fontes"].get(nome,0)+1
                        cont_corpo["tams"][tam_pt] = cont_corpo["tams"].get(tam_pt,0)+1
                        if nome == FONTE_TABELAS: cont_corpo["fonte_ok"] += 1
                        if tam_pt == TAMANHO_TABELAS: cont_corpo["tam_ok"] += 1
    pct_f = calcular_pct(cont_corpo["fonte_ok"], cont_corpo["total"])
    pct_t = calcular_pct(cont_corpo["tam_ok"], cont_corpo["total"])
    return {
        "corpo": {
            "fontes": cont_corpo["fontes"], "tamanhos": cont_corpo["tams"],
            "pct_fonte": pct_f, "pct_tam": pct_t,
            "fonte_ok": pct_f >= LIMIAR_CONFORMIDADE,
            "tam_ok": pct_t >= LIMIAR_CONFORMIDADE,
        }
    }

# ============================================================
# 🔍 ESCANEAR DOCUMENTO — ORDEM DE DETECÇÃO CORRIGIDA
# ============================================================
def escanear(doc_bytes):
    doc = docx.Document(BytesIO(doc_bytes))
    texto_completo = ""
    codigo = versao = None

    for tb in doc.tables:
        for ln in tb.rows:
            tx = " ".join([cel.text for cel in ln.cells])
            texto_completo += tx + " "
            if not codigo:
                m = re.search(r'C[ÓO]DIGO\s*[:：=\-]?\s*([A-Z0-9_\-\/\.]+)', tx.upper())
                if m: codigo = m.group(1).strip()
            if not versao:
                m = re.search(r'VERS[AÃ]O\s*[:：=\-]?\s*(\d+(?:[.\-]\d+)*)', tx.upper())
                if m: versao = m.group(1).strip()

    for p in doc.paragraphs:
        texto_completo += p.text + " "
        txu = p.text.upper()
        if not codigo:
            m = re.search(r'C[ÓO]DIGO\s*[:：=\-]?\s*([A-Z0-9_\-\/\.]+)', txu)
            if m: codigo = m.group(1).strip()
        if not versao:
            m = re.search(r'VERS[AÃ]O\s*[:：=\-]?\s*(\d+(?:[.\-]\d+)*)', txu)
            if m: versao = m.group(1).strip()

    texto_limpo = limpar_texto(texto_completo)

    # ✅ ORDEM DE PRIORIDADE CORRIGIDA — POP ANTES DE PLAN!
    tipo = "PROT"
    mapeamento = [
        ("POP", r'\bPOP\b|\bPROCEDIMENTO OPERACIONAL\b'),
        ("POI", r'\bPOL[IÍ]TICA\b|\bPOI\b'),
        ("NOR", r'\bNORMA\b|\bNOR\b'),
        ("REG", r'\bREGIMENTO\b|\bREGULAMENTO\b|\bREG\b'),
        ("PROG", r'\bPROGRAMA\b|\bPROG\b'),
        ("PLAN", r'\bPLANO\b|\bPLAN\b'),
        ("ROT", r'\bROTINA\b|\bROT\b'),
        ("MAN", r'\bMANUAL\b|\bMAN\b'),
        ("PROT", r'\bPROTOCOLO\b'),
    ]
    for t, padrao in mapeamento:
        if re.search(padrao, texto_limpo):
            tipo = t
            break

    secoes_tipo = SECOES_POR_TIPO.get(tipo, SECOES_POR_TIPO["PROT"])
    obr_enc, obr_falt, opc_enc, opc_falt = [], [], [], []
    for s in secoes_tipo["obrigatorias"]:
        if limpar_texto(s) in texto_limpo: obr_enc.append(s)
        else: obr_falt.append(s)
    for s in secoes_tipo["opcionais"]:
        if limpar_texto(s) in texto_limpo: opc_enc.append(s)
        else: opc_falt.append(s)

    return {
        "tipo": tipo, "codigo": codigo, "versao": versao,
        "fonte": verificar_fonte(doc),
        "secoes_obrig_enc": obr_enc, "secoes_obrig_falt": obr_falt,
        "secoes_opc_enc": opc_enc, "secoes_opc_falt": opc_falt,
    }

# ============================================================
# 🧹 APLICAR MARGENS
# ============================================================
def aplicar_margens(doc_bytes):
    try:
        doc = docx.Document(BytesIO(doc_bytes))
        for sec in doc.sections:
            sec.top_margin = Cm(MARGEM_SUP_ESPERADA)
            sec.bottom_margin = Cm(MARGEM_INF_ESPERADA)
            sec.left_margin = Cm(MARGEM_ESQ_ESPERADA)
            sec.right_margin = Cm(MARGEM_DIR_ESPERADA)
        saida = BytesIO()
        doc.save(saida)
        return saida.getvalue()
    except Exception as e:
        st.warning(f"Não foi possível ajustar margens: {str(e)}")
        return doc_bytes

# ============================================================
# 📄 GERAR FICHA
# ============================================================
def gerar_ficha_verificacao(nome_arq, tipo, cod, ver, fonte_ok, secoes_dados, itens_manuais):
    obr_enc, obr_falt, opc_enc, opc_falt = secoes_dados
    status_final = "✅ SIM" if (fonte_ok and len(obr_falt)==0 and all(itens_manuais.values())) else "⚠️ COM PENDÊNCIAS"
    ficha = f"""
==================================================
          FICHA DE VERIFICAÇÃO — AUDITOR NAQH NMZ
==================================================
Arquivo: {nome_arq}
Tipo de Documento: {tipo}
--------------------------------------------------
DADOS DE IDENTIFICAÇÃO
--------------------------------------------------
Código:     {cod or 'NÃO INFORMADO'}
Versão:     {ver or 'NÃO INFORMADO'}
--------------------------------------------------
CONFORMIDADE TÉCNICA — NORMA ZERO
--------------------------------------------------
MARGENS (Sup {MARGEM_SUP_ESPERADA} / Inf {MARGEM_INF_ESPERADA} / Esq {MARGEM_ESQ_ESPERADA} / Dir {MARGEM_DIR_ESPERADA} cm): ✅ APLICADAS NO DOWNLOAD
FONTE ({FONTE_CORPO} {TAMANHO_CORPO}pt / {FONTE_TABELAS} {TAMANHO_TABELAS}pt): {'✅ CONFORME' if fonte_ok else '❌ NÃO CONFORME'}
--------------------------------------------------
CONFERÊNCIA MANUAL — CABEÇALHO
--------------------------------------------------
"""
    for item, ok in itens_manuais.items():
        ficha += f"{item}: {'✅ CONFERIDO' if ok else '❌ NÃO CONFERIDO'}\n"
    ficha += f"""
--------------------------------------------------
SEÇÕES DO DOCUMENTO
--------------------------------------------------
Obrigatórias encontradas: {len(obr_enc)} / {len(obr_enc)+len(obr_falt)}
Obrigatórias faltantes:   {len(obr_falt)}
Opcionais encontradas:    {len(opc_enc)} / {len(opc_enc)+len(opc_falt)}
--------------------------------------------------
RESULTADO FINAL
--------------------------------------------------
APROVADO CONFORME NORMA ZERO: {status_final}
==================================================
Ezequias Santos — Agente Administrativo
Auditor NAQH NMZ
==================================================
"""
    return ficha.encode('utf-8')

# ============================================================
# 🚀 INTERFACE PRINCIPAL
# ============================================================
st.markdown("""
    <style>
    .header-container{{display:flex;justify-content:space-between;align-items:center;background:#0F172A;padding:15px 25px;border-radius:12px;margin-bottom:20px}}
    .header-text h1{{color:#FFF;margin:0;font-size:28px}}
    .header-text p{{color:#94A3B8;margin:5px 0 0 0}}
    .header-emoji{{font-size:50px}}
    .relogio-box{{background:linear-gradient(135deg,#0F766E,#14B8A6);padding:20px 25px;border-radius:12px;color:#FFF;margin:15px 0}}
    .economia-grande{{font-size:32px;font-weight:bold}}
    .base-tempo{{background:#1E293B;padding:12px 18px;border-radius:8px;color:#CBD5E1;font-size:13px;margin-top:8px}}
    </style>
""", unsafe_allow_html=True)

st.markdown(f"""
    <div class="header-container">
        <div class="header-text">
            <h1>AUDITOR NAQH NMZ DE ALTA PRECISÃO</h1>
            <p>Ezequias Santos — Agente Administrativo | Margens: Esq {MARGEM_ESQ_ESPERADA} / Dir {MARGEM_DIR_ESPERADA} / Sup {MARGEM_SUP_ESPERADA} / Inf {MARGEM_INF_ESPERADA} cm • Calibri {TAMANHO_CORPO} / {TAMANHO_TABELAS}</p>
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

            # ✅ PRIORIDADE: detectar pelo NOME do arquivo
            tipo_nome = detectar_pelo_nome(arq.name)
            if tipo_nome:
                r["tipo"] = tipo_nome
                st.success(f"✅ Tipo definido pelo nome do arquivo: **{tipo_nome}**")

            st.markdown("### 📋 DADOS DO DOCUMENTO")
            c1, c2 = st.columns(2)
            with c1:
                st.info(f"**Tipo Detectado:** {r['tipo']}")

            with c2:
                codigo_final = st.text_input("**CÓDIGO**", value=r['codigo'] or "",
                    placeholder="Digite o código aqui", key=f"cod{idx}")
                if r['codigo']: st.success(f"✅ Detectado: {r['codigo']}")
                elif codigo_final: st.info("✍️ Digitado manualmente")
                else: st.warning("⚠️ Não detectado — digite acima")

            versao_final = st.text_input("**VERSÃO / SÉRIE**", value=r['versao'] or "",
                placeholder="Digite a versão/série aqui", key=f"ver{idx}")
            if r['versao']: st.success(f"✅ Detectado: {r['versao']}")
            elif versao_final: st.info("✍️ Digitado manualmente")
            else: st.warning("⚠️ Não detectado — digite acima")

            st.markdown("---")
            st.info(f"📏 **MARGENS NORMA ZERO:** Serão aplicadas ao baixar ✅")
            st.markdown("---")

            st.markdown(f"### ✍️ FONTE — {FONTE_CORPO} {TAMANHO_CORPO}pt / {FONTE_TABELAS} {TAMANHO_TABELAS}pt")
            f = r["fonte"]["corpo"]
            st.write("Fontes:", ", ".join(f"{n} ({q})" for n,q in f["fontes"].items()) or "Nenhuma")
            st.write("Tamanhos:", ", ".join(f"{t}pt ({q})" for t,q in f["tamanhos"].items()) or "Nenhum")
            fonte_ok = f["fonte_ok"] and f["tam_ok"]
            st.metric("Conformidade", f"{f['pct_fonte']}%", "✅" if fonte_ok else "❌")
            st.markdown("---")

            st.markdown("### ✅ CONFERÊNCIA MANUAL — CABEÇALHO")
            st.info("💡 Confira e marque:")
            itens_cabecalho = {
                "Cabeçalho padrão": st.checkbox("✅ Cabeçalho", key=f"cab{idx}"),
                "Número de páginas": st.checkbox("✅ Páginas", key=f"pag{idx}"),
                "Versão/Série no cabeçalho": st.checkbox("✅ Versão", key=f"verif{idx}"),
                "Logo do Hospital": st.checkbox("✅ Logo", key=f"logo{idx}"),
                "Marca d'água": st.checkbox("✅ Marca d'água", key=f"marca{idx}"),
            }
            st.markdown("---")

            st.markdown(f"### 📑 SEÇÕES — Tipo: {r['tipo']}")
            st.info("💡 Apêndices e Anexos são OPCIONAIS")
            secoes_usuario_ok = True
            st.markdown("**🔴 Obrigatórias:**")
            for s in r["secoes_obrig_enc"]: st.success(f"✅ {s}")
            for s in r["secoes_obrig_falt"]:
                marc = st.checkbox(f"⚠️ {s} — Não detectado (marque se encontrou)",
                                   key=f"obr_falt_{idx}_{limpar_texto(s)}")
                if not marc: secoes_usuario_ok = False
            if r["secoes_opc_enc"] or r["secoes_opc_falt"]:
                st.markdown("**🟡 Opcionais:**")
                for s in r["secoes_opc_enc"]: st.info(f"🟡 {s} — Presente")
                for s in r["secoes_opc_falt"]: st.info(f"🟡 {s} — Opcional, sem problema")

            st.markdown("---")
            cabecalho_ok = all(itens_cabecalho.values())
            aprov = bool(codigo_final and versao_final and secoes_usuario_ok and cabecalho_ok and fonte_ok)
            if aprov: st.success("✅ APROVADO CONFORME NORMA ZERO")
            else: st.warning("⚠️ COM PENDÊNCIAS — verifique itens acima")

            cod_nome = re.sub(r'[<>:"/\\|?*º°]', '-', codigo_final.strip()) if codigo_final else "DOC"
            ver_nome = re.sub(r'[<>:"/\\|?*º°]', '_', versao_final.strip()) if versao_final else "0"
            nome_base = f"{cod_nome}_v{ver_nome}"

            dados_format = aplicar_margens(dados)
            st.download_button(f"📥 BAIXAR: {nome_base}.docx", dados_format,
                f"{nome_base}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_doc_{idx}", type="primary")

            ficha = gerar_ficha_verificacao(arq.name, r['tipo'], codigo_final, versao_final,
                fonte_ok, (r["secoes_obrig_enc"], r["secoes_obrig_falt"], r["secoes_opc_enc"], r["secoes_opc_falt"]),
                itens_cabecalho)
            st.download_button(f"📋 BAIXAR FICHA: {nome_base}_FICHA.txt", ficha,
                f"{nome_base}_FICHA.txt", "text/plain", key=f"dl_ficha_{idx}")

        except Exception as e:
            st.error(f"❌ Erro ao processar {arq.name}: {str(e)}")
        st.markdown("---")
