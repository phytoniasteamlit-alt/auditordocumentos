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

# ✅ VALOR EXATO: 1 cm = 566.928 twips
TWIPS_PARA_CM = 566.928

# ============================================================
# 📋 SEÇÕES POR TIPO — APÊNDICES E ANEXOS SÃO OPCIONAIS
# ============================================================
SECOES_POR_TIPO = {
    "PROT": {
        "obrigatorias": [
            "OBJETIVO",
            "APLICAÇÃO",
            "REFERENCIAL TEÓRICO",
            "CLASSIFICAÇÃO DAS CIRURGIAS POR POTENCIAL DE CONTAMINAÇÃO",
            "FATORES DE RISCO",
            "BENEFÍCIOS E RISCOS",
            "PRINCÍPIOS GERAIS",
            "CRITÉRIOS DE ELEGIBILIDADE",
            "ESQUEMA PADRÃO",
            "ESQUEMAS POR TIPO DE CIRURGIA"
        ],
        "opcionais": ["APÊNDICES", "ANEXOS"]
    },
    "POP": {
        "obrigatorias": ["DEFINIÇÃO", "APLICABILIDADE", "RESPONSÁVEL PELA EXECUÇÃO",
                        "MATERIAIS UTILIZADOS", "DESCRIÇÃO DA TAREFA",
                        "ATIVIDADES", "REFERÊNCIAS"],
        "opcionais": ["ANEXOS"]
    },
    "POI": {
        "obrigatorias": ["INTRODUÇÃO", "OBJETIVO", "PRINCÍPIOS", "DIRETRIZES",
                        "RESPONSABILIDADES", "ESTRATÉGIA DE MONITORAMENTO",
                        "REFERÊNCIAS"],
        "opcionais": ["APÊNDICES E ANEXOS"]
    },
    "NOR": {
        "obrigatorias": ["OBJETIVO", "APLICABILIDADE", "DESCRIÇÃO DA NORMA",
                        "RESPONSÁVEL", "EFEITOS NO CUMPRIMENTO",
                        "NORMA DE REFERÊNCIA"],
        "opcionais": ["ANEXOS"]
    },
    "REG": {
        "obrigatorias": ["DA FINALIDADE", "DA COMPOSIÇÃO — MEMBROS", "DO MANDATO",
                        "DO FUNCIONAMENTO E ORGANIZAÇÃO", "DAS ATRIBUIÇÕES",
                        "DISPOSIÇÕES FINAIS"],
        "opcionais": []
    },
    "PROG": {
        "obrigatorias": ["REFERENCIAL TEÓRICO", "PADRONIZAÇÃO DE ROTINAS",
                        "ESTRATÉGIAS DE MONITORAMENTO", "DESCRIÇÃO DO PROGRAMA",
                        "MEDIDAS EDUCACIONAIS", "REFERÊNCIAS"],
        "opcionais": ["APÊNDICES", "ANEXOS"]
    },
    "PLAN": {
        "obrigatorias": ["OBJETIVO", "APLICABILIDADE", "DEFINIÇÃO DE TERMOS",
                        "IDENTIFICAÇÃO DE RISCO ATUAL", "MEDIDAS DE CONTINGÊNCIA",
                        "REFERÊNCIAS"],
        "opcionais": ["APÊNDICES", "ANEXOS"]
    },
    "ROT": {
        "obrigatorias": ["DEFINIÇÃO", "OBJETIVO", "APLICABILIDADE",
                        "DESCRIÇÃO DA ROTINA"],
        "opcionais": ["APÊNDICES"]
    },
    "MAN": {
        "obrigatorias": ["CAPA", "ELABORADORES", "COLABORADORES", "SUMÁRIO",
                        "APRESENTAÇÃO", "DESCRIÇÃO", "REFERÊNCIAS"],
        "opcionais": ["APÊNDICES E ANEXOS"]
    }
}

# ============================================================
# 🧹 FUNÇÕES AUXILIARES
# ============================================================
def limpar_texto(texto):
    if not texto: return ""
    # Remove acentos
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII','ignore').decode('ASCII')
    # Remove espaços múltiplos, tab, quebra de linha, pontos, hífens
    texto = re.sub(r'[\s.\-_\t]+', ' ', texto).upper().strip()
    # Remove número no início (ex: "4. TEXTO" → "TEXTO")
    texto = re.sub(r'^\d+\s*', '', texto).strip()
    return texto

def formatar_tempo(minutos_total):
    h = int(minutos_total // 60)
    m = int(minutos_total % 60)
    if h > 0: return f"{h}h {m}min" if m>0 else f"{h}h"
    return f"{m}min"

# ============================================================
# 📏 VERIFICAR MARGENS — ✅ CONVERSÃO 100% CORRIGIDA!
# ============================================================
def verificar_margens(doc):
    sec = doc.sections[0]
    
    def cm_de_twips(valor):
        if valor is None or valor == 0:
            return 0.0
        # ✅ CONVERSÃO CORRETA: twips ÷ 566.928 = cm
        return round(valor / TWIPS_PARA_CM, 2)

    m_sup = cm_de_twips(sec.top_margin)
    m_inf = cm_de_twips(sec.bottom_margin)
    m_esq = cm_de_twips(sec.left_margin)
    m_dir = cm_de_twips(sec.right_margin)
    
    tol = 0.5  # Tolerância maior para não errar por poucos mm
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
            for cel in ln.cells:
                texto_tb += cel.text + " "
                for p in cel.paragraphs:
                    for run in p.runs:
                        cont_tab["total"] += 1
                        nome = run.font.name or FONTE_TABELAS
                        tam_pt = round(run.font.size.pt,1) if run.font.size else TAMANHO_TABELAS
                        cont_tab["fontes"][nome] = cont_tab["fontes"].get(nome,0)+1
                        cont_tab["tams"][tam_pt] = cont_tab["tams"].get(tam_pt,0)+1
                        if nome == FONTE_TABELAS: cont_tab["fonte_ok"] += 1
                        if tam_pt == TAMANHO_TABELAS: cont_tab["tam_ok"] += 1
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
# 🔍 ESCANEAR DOCUMENTO — ✅ BUSCA FLEXÍVEL NAS SEÇÕES!
# ============================================================
def escanear(doc_bytes):
    doc = docx.Document(BytesIO(doc_bytes))
    texto_completo = ""
    codigo = versao = None

    # Escaneia tabelas (cabeçalho)
    for tb in doc.tables:
        for ln in tb.rows:
            texto_linha = " ".join([cel.text for cel in ln.cells])
            texto_completo += texto_linha + " "
            if not codigo:
                for pad in [r'C[ÓO]DIGO\s*[:：=\-]?\s*([A-Z0-9_\-\/\.]+)', r'COD[:\s]*([A-Z0-9_\-\/\.]+)']:
                    m = re.search(pad, texto_linha.upper())
                    if m: codigo = m.group(1).strip(); break
            if not versao:
                for pad in [r'VERS[AÃ]O\s*[:：=\-]?\s*(\d+(?:[.\-]\d+)*)', r'VERS[:\s]*([\d.]+)']:
                    m = re.search(pad, texto_linha.upper())
                    if m: versao = m.group(1).strip(); break

    # Escaneia parágrafos
    for p in doc.paragraphs:
        texto_completo += p.text + " "
        texto_upper = p.text.upper()
        if not codigo:
            for pad in [r'C[ÓO]DIGO\s*[:：=\-]?\s*([A-Z0-9_\-\/\.]+)', r'COD[:\s]*([A-Z0-9_\-\/\.]+)']:
                m = re.search(pad, texto_upper)
                if m: codigo = m.group(1).strip(); break
        if not versao:
            for pad in [r'VERS[AÃ]O\s*[:：=\-]?\s*(\d+(?:[.\-]\d+)*)', r'VERS[:\s]*([\d.]+)']:
                m = re.search(pad, texto_upper)
                if m: versao = m.group(1).strip(); break

    # ✅ TEXTO LIMPO PARA BUSCA DE SEÇÕES (flexível!)
    texto_limpo = limpar_texto(texto_completo)

    # Detectar tipo
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

    secoes_tipo = SECOES_POR_TIPO[tipo]
    obr_enc, obr_falt = [], []
    opc_enc, opc_falt = [], []
    
    # ✅ BUSCA FLEXÍVEL: compara só o texto limpo, sem número, sem hífen, sem espaço
    for s in secoes_tipo["obrigatorias"]:
        if limpar_texto(s) in texto_limpo:
            obr_enc.append(s)
        else:
            obr_falt.append(s)
    
    for s in secoes_tipo["opcionais"]:
        if limpar_texto(s) in texto_limpo:
            opc_enc.append(s)
        else:
            opc_falt.append(s)

    return {
        "tipo": tipo, "codigo": codigo, "versao": versao,
        "margens": verificar_margens(doc),
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
            sec.top_margin = docx.shared.Cm(MARGEM_SUP_ESPERADA)
            sec.bottom_margin = docx.shared.Cm(MARGEM_INF_ESPERADA)
            sec.left_margin = docx.shared.Cm(MARGEM_ESQ_ESPERADA)
            sec.right_margin = docx.shared.Cm(MARGEM_DIR_ESPERADA)
        saida = BytesIO()
        doc.save(saida)
        return saida.getvalue()
    except:
        return doc_bytes

# ============================================================
# 📄 GERAR FICHA DE VERIFICAÇÃO
# ============================================================
def gerar_ficha_verificacao(nome_arq, tipo, cod, ver, margens_ok, fonte_ok, secoes_dados, itens_manuais):
    obr_enc, obr_falt, opc_enc, opc_falt = secoes_dados
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
MARGENS (3,0 / 2,0 / 3,0 / 2,0 cm): {'✅ CONFORME' if margens_ok else '❌ NÃO CONFORME'}
FONTE (Calibri 11pt / 10pt):        {'✅ CONFORME' if fonte_ok else '❌ NÃO CONFORME'}

--------------------------------------------------
CONFERÊNCIA MANUAL — ITENS DO CABEÇALHO
--------------------------------------------------
"""
    for item, ok in itens_manuais.items():
        ficha += f"{item}: {'✅ CONFERIDO' if ok else '❌ NÃO CONFERIDO'}\n"

    ficha += f"""
--------------------------------------------------
SEÇÕES DO DOCUMENTO
--------------------------------------------------
Obrigatórias encontradas: {len(obr_enc)} / {len(obr_enc) + len(obr_falt)}
Obrigatórias faltantes:   {len(obr_falt)}
Opcionais encontradas:    {len(opc_enc)} / {len(opc_enc) + len(opc_falt)}

--------------------------------------------------
RESULTADO FINAL
--------------------------------------------------
APROVADO CONFORME NORMA ZERO: {'✅ SIM' if (margens_ok and fonte_ok and len(obr_falt)==0 and all(itens_manuais.values())) else '⚠️ COM PENDÊNCIAS'}

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
            
            # ✅ CÓDIGO — USO O DIGITADO PELO USUÁRIO
            with c2:
                codigo_final = st.text_input("**CÓDIGO**", value=r['codigo'] or "", placeholder="Digite o código aqui", key=f"cod{idx}")
                if r['codigo']: st.success(f"✅ Detectado: {r['codigo']}")
                elif codigo_final: st.info(f"✍️ Digitado manualmente")
                else: st.warning("⚠️ Não detectado — digite acima")

            versao_final = st.text_input("**VERSÃO / SÉRIE**", value=r['versao'] or "", placeholder="Digite a versão/série aqui", key=f"ver{idx}")
            if r['versao']: st.success(f"✅ Detectado: {r['versao']}")
            elif versao_final: st.info(f"✍️ Digitado manualmente")
            else: st.warning("⚠️ Não detectado — digite acima")

            st.markdown("---")
            st.markdown("### 📏 MARGENS — Esperado: Sup=3,0 / Inf=2,0 / Esq=3,0 / Dir=2,0 cm")
            m = r["margens"]
            col1,col2,col3,col4 = st.columns(4)
            col1.metric("Superior", f"{m['sup']} cm", "✅" if m['ok_sup'] else "❌")
            col2.metric("Inferior", f"{m['inf']} cm", "✅" if m['ok_inf'] else "❌")
            col3.metric("Esquerda", f"{m['esq']} cm", "✅" if m['ok_esq'] else "❌")
            col4.metric("Direita", f"{m['dir']} cm", "✅" if m['ok_dir'] else "❌")
            
            if m["todas_ok"]:
                st.success("✅ TODAS AS MARGENS CONFORME NORMA ZERO")
            else:
                st.error("❌ MARGENS NÃO CONFORME — clique em BAIXAR para corrigir")

            st.markdown("---")
            st.markdown("### ✍️ CORPO — Calibri 11pt")
            f = r["fonte"]["corpo"]
            st.write("Fontes encontradas:", ", ".join(f"{n} ({q})" for n,q in f["fontes"].items()))
            st.write("Tamanhos encontrados:", ", ".join(f"{t}pt ({q})" for t,q in f["tamanhos"].items()))
            fonte_ok = f["fonte_ok"] and f["tam_ok"]
            st.metric("Conformidade", f"{f['pct_fonte']}%", "✅" if fonte_ok else "❌")

            st.markdown("---")
            st.markdown("### ✅ CONFERÊNCIA MANUAL — CABEÇALHO E IDENTIFICAÇÃO")
            st.info("💡 Confira visualmente e marque os itens abaixo:")

            itens_cabecalho = {
                "Cabeçalho presente e correto": st.checkbox("✅ Cabeçalho padrão", key=f"cab{idx}"),
                "Número de páginas visível": st.checkbox("✅ Número de páginas", key=f"pag{idx}"),
                "Versão/Série no cabeçalho": st.checkbox("✅ Versão/Série", key=f"verif{idx}"),
                "Logo do Hospital presente": st.checkbox("✅ Logo Hospital", key=f"logo{idx}"),
                "Marca d'água 'HOSPITAL DA CIDADE'": st.checkbox("✅ Marca d'água ao fundo", key=f"marca{idx}"),
            }

            st.markdown("---")
            st.markdown(f"### 📑 SEÇÕES — Tipo: {r['tipo']}")
            st.info("💡 Apêndices e Anexos são OPCIONAIS — não precisa marcar se não houver")

            secoes_usuario_ok = True
            st.markdown("**🔴 Obrigatórias:**")
            for s in r["secoes_obrig_enc"]:
                st.success(f"✅ {s} — DETECTADO")
            for s in r["secoes_obrig_falt"]:
                marc = st.checkbox(f"⚠️ {s} — NÃO DETECTADO (marque se encontrou)", key=f"obr_falt{idx}_{limpar_texto(s)}")
                if not marc: secoes_usuario_ok = False

            if r["secoes_opc_enc"] or r["secoes_opc_falt"]:
                st.markdown("**🟡 Opcionais (Apêndices/Anexos):**")
                for s in r["secoes_opc_enc"]:
                    st.info(f"🟡 {s} — DETECTADO (opcional)")
                for s in r["secoes_opc_falt"]:
                    st.info(f"🟡 {s} — Não encontrado (opcional, sem problema)")

            st.markdown("---")
            faltam_obrig = len(r["secoes_obrig_falt"])
            cabecalho_ok = all(itens_cabecalho.values())
            aprov = m["todas_ok"] and codigo_final and versao_final and secoes_usuario_ok and cabecalho_ok
            
            if aprov:
                st.success("✅ DOCUMENTO APROVADO CONFORME NORMA ZERO")
            else:
                st.warning("⚠️ DOCUMENTO COM PENDÊNCIAS — verifique itens acima")

            # ✅ NOME DO ARQUIVO = CÓDIGO + VERSÃO
            cod_nome = re.sub(r'[<>:"/\\|?*º°]', '-', codigo_final) if codigo_final else "DOC"
            ver_nome = re.sub(r'[<>:"/\\|?*º°]', '_', versao_final) if versao_final else "0"
            nome_base = f"{cod_nome}_v{ver_nome}"

            # ✅ BAIXAR DOCUMENTO FORMATADO
            dados_format = aplicar_margens(dados)
            st.download_button(
                f"📥 BAIXAR: {nome_base}.docx",
                dados_format,
                f"{nome_base}.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_doc{idx}",
                type="primary"
            )

            # ✅ BAIXAR FICHA DE VERIFICAÇÃO
            ficha = gerar_ficha_verificacao(
                arq.name, r['tipo'], codigo_final, versao_final,
                m["todas_ok"], fonte_ok,
                (r["secoes_obrig_enc"], r["secoes_obrig_falt"], r["secoes_opc_enc"], r["secoes_opc_falt"]),
                itens_cabecalho
            )
            st.download_button(
                f"📋 BAIXAR FICHA: {nome_base}_FICHA.txt",
                ficha,
                f"{nome_base}_FICHA.txt",
                "text/plain",
                key=f"dl_ficha{idx}"
            )

        except Exception as e:
            st.error(f"❌ Erro ao processar: {str(e)}")

        st.markdown("---")
