import streamlit as st
import docx
import re
import unicodedata
from io import BytesIO

# ============================================================
# 🎯 CONFIGURAÇÕES DA PÁGINA
# ============================================================
st.set_page_config(page_title="AUDITOR NAQH NMZ", page_icon="👨‍💻", layout="wide")

# ============================================================
# 📏 VALORES EXATOS — ABNT NBR 14724 / NORMA ZERO
# ============================================================
MARGEM_SUP_ESPERADA = 3.0   # cm
MARGEM_INF_ESPERADA = 2.0   # cm
MARGEM_ESQ_ESPERADA = 3.0   # cm ✅
MARGEM_DIR_ESPERADA = 2.0   # cm ✅

FONTE_CORPO = "Calibri"
TAMANHO_CORPO = 11

FONTE_TABELAS = "Calibri"
TAMANHO_TABELAS = 10

# ============================================================
# 🧹 REMOVER ACENTOS
# ============================================================
def limpar_texto(texto):
    if not texto: return ""
    sem_acento = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return re.sub(r'\s+', ' ', sem_acento.upper().strip())

# ============================================================
# 📋 SEÇÕES POR TIPO DE DOCUMENTO
# ============================================================
SECOES_POR_TIPO = {
    "PROT": ["1. OBJETIVO", "2. APLICABILIDADE", "3. REFERENCIAL TEORICO",
             "4. DESCRICAO DO PROTOCOLO", "5. ESTRATEGIAS DE MONITORAMENTO",
             "6. REFERENCIAS", "7. APENDICES", "8. ANEXOS"],
    "POP": ["1. DEFINICAO", "2. APLICABILIDADE", "3. RESPONSAVEL PELA EXECUCAO",
             "4. MATERIAIS UTILIZADOS", "5. DESCRICAO DA TAREFA",
             "6. ATIVIDADES", "7. REFERENCIAS", "8. ANEXOS"],
    "POI": ["1. INTRODUCAO", "2. OBJETIVO", "3. PRINCIPIOS", "4. DIRETRIZES",
            "5. RESPONSABILIDADES", "6. ESTRATEGIA DE MONITORAMENTO",
            "7. REFERENCIAS", "8. APENDICES E ANEXOS"],
    "NOR": ["1. OBJETIVO", "2. APLICABILIDADE", "3. DESCRICAO DA NORMA",
            "4. RESPONSAVEL", "5. EFEITOS NO CUMPRIMENTO",
            "6. NORMA DE REFERENCIA", "7. ANEXOS"],
    "REG": ["1. DA FINALIDADE", "2. DA COMPOSICAO — MEMBROS", "3. DO MANDATO",
            "4. DO FUNCIONAMENTO E ORGANIZACAO", "5. DAS ATRIBUICOES",
            "6. DISPOSICOES FINAIS"],
    "PROG": ["1. REFERENCIAL TEORICO", "2. PADRONIZACAO DE ROTINAS",
             "3. ESTRATEGIAS DE MONITORAMENTO", "4. DESCRICAO DO PROGRAMA",
             "5. MEDIDAS EDUCACIONAIS", "6. REFERENCIAS",
             "7. APENDICES", "8. ANEXOS"],
    "PLAN": ["1. OBJETIVO", "2. APLICABILIDADE", "3. DEFINICAO DE TERMOS",
             "4. IDENTIFICACAO DE RISCO ATUAL", "5. MEDIDAS DE CONTINGENCIA",
             "6. REFERENCIAS", "7. APENDICES", "8. ANEXOS"],
    "ROT": ["1. DEFINICAO", "2. OBJETIVO", "3. APLICABILIDADE",
            "4. DESCRICAO DA ROTINA", "5. APENDICES"],
    "MAN": ["1. CAPA", "2. ELABORADORES", "3. COLABORADORES", "4. SUMARIO",
            "5. APRESENTACAO", "6. DESCRICAO", "7. REFERENCIAS",
            "8. APENDICES E ANEXOS"]
}

# ============================================================
# 📏 VERIFICAR MARGENS
# ============================================================
def verificar_margens(doc):
    sec = doc.sections[0]
    def twips_para_cm(v): return round(v / 567.0, 2) if v else 0.00
    
    m_sup = twips_para_cm(sec.top_margin)
    m_inf = twips_para_cm(sec.bottom_margin)
    m_esq = twips_para_cm(sec.left_margin)
    m_dir = twips_para_cm(sec.right_margin)
    
    tol = 0.05
    ok_sup = abs(m_sup - MARGEM_SUP_ESPERADA) < tol
    ok_inf = abs(m_inf - MARGEM_INF_ESPERADA) < tol
    ok_esq = abs(m_esq - MARGEM_ESQ_ESPERADA) < tol
    ok_dir = abs(m_dir - MARGEM_DIR_ESPERADA) < tol
    
    return {
        "sup": m_sup, "inf": m_inf, "esq": m_esq, "dir": m_dir,
        "ok_sup": ok_sup, "ok_inf": ok_inf, "ok_esq": ok_esq, "ok_dir": ok_dir,
        "todas_ok": ok_sup and ok_inf and ok_esq and ok_dir
    }

# ============================================================
# ✍️ VERIFICA FONTE SEPARADA: CORPO ≠ TABELAS
# ============================================================
def verificar_fonte_separado(doc):
    # --- CORPO DO TEXTO → Calibri 11 ---
    corpo_fontes = {}
    corpo_tamanhos = {}
    corpo_total = corpo_fonte_ok = corpo_tam_ok = 0
    
    for p in doc.paragraphs:
        for run in p.runs:
            corpo_total += 1
            nome = run.font.name or "NÃO DEFINIDA"
            tam = round(run.font.size.pt, 1) if run.font.size else 0
            corpo_fontes[nome] = corpo_fontes.get(nome, 0) + 1
            corpo_tamanhos[tam] = corpo_tamanhos.get(tam, 0) + 1
            if nome == FONTE_CORPO: corpo_fonte_ok += 1
            if tam == TAMANHO_CORPO: corpo_tam_ok += 1
    
    # --- DENTRO DAS TABELAS → Calibri 10 ---
    tabela_fontes = {}
    tabela_tamanhos = {}
    tabela_total = tabela_fonte_ok = tabela_tam_ok = 0
    tem_registro_historico = False
    
    for tabela in doc.tables:
        texto_tabela = ""
        for linha in tabela.rows:
            for celula in linha.cells:
                texto_tabela += cel.text + " "
                for p in celula.paragraphs:
                    for run in p.runs:
                        tabela_total += 1
                        nome = run.font.name or "NÃO DEFINIDA"
                        tam = round(run.font.size.pt, 1) if run.font.size else 0
                        tabela_fontes[nome] = tabela_fontes.get(nome, 0) + 1
                        tabela_tamanhos[tam] = tabela_tamanhos.get(tam, 0) + 1
                        if nome == FONTE_TABELAS: tabela_fonte_ok += 1
                        if tam == TAMANHO_TABELAS: tabela_tam_ok += 1
        
        if "REGISTRO HISTÓRICO" in texto_tabela.upper() or "REGISTRO HISTORICO" in texto_tabela.upper():
            tem_registro_historico = True
    
    def calc_pct(total, ok): return round((ok/total*100),1) if total else 0
    crit = 70
    
    return {
        "corpo": {
            "fontes": corpo_fontes, "tamanhos": corpo_tamanhos,
            "fonte_ok": calc_pct(corpo_total, corpo_fonte_ok) >= crit,
            "tam_ok": calc_pct(corpo_total, corpo_tam_ok) >= crit,
            "pct_fonte": calc_pct(corpo_total, corpo_fonte_ok),
            "pct_tam": calc_pct(corpo_total, corpo_tam_ok),
            "esperado_fonte": FONTE_CORPO, "esperado_tam": TAMANHO_CORPO
        },
        "tabelas": {
            "fontes": tabela_fontes, "tamanhos": tabela_tamanhos,
            "fonte_ok": calc_pct(tabela_total, tabela_fonte_ok) >= crit,
            "tam_ok": calc_pct(tabela_total, tabela_tam_ok) >= crit,
            "pct_fonte": calc_pct(tabela_total, tabela_fonte_ok),
            "pct_tam": calc_pct(tabela_total, tabela_tam_ok),
            "esperado_fonte": FONTE_TABELAS, "esperado_tam": TAMANHO_TABELAS,
            "tem_registro_historico": tem_registro_historico
        }
    }

# ============================================================
# 📊 LISTAR TABELAS
# ============================================================
def listar_tabelas(doc):
    info = []
    tem_cab = False
    for i, tb in enumerate(doc.tables):
        texto = " ".join([c.text.upper() for c in tb.rows[0].cells])
        eh_cab = "CODIGO" in texto or "CÓDIGO" in texto or "VERSÃO" in texto
        eh_hist = "REGISTRO HISTÓRICO" in texto or "REGISTRO HISTORICO" in texto
        if eh_cab: tem_cab = True
        info.append({"n":i+1, "lin":len(tb.rows), "col":len(tb.columns), "cab":eh_cab, "hist":eh_hist})
    return {"total":len(doc.tables), "tem_cabecalho":tem_cab, "detalhes":info}

# ============================================================
# 🧹 APLICAR MARGENS CORRIGIDAS
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
# 📄 GERAR FICHA
# ============================================================
def gerar_ficha(tipo, codigo, versao, m, f, tb, enc, falt, aprov):
    ficha = [
        "="*60, "FICHA DE VERIFICAÇÃO — NORMA ZERO", "="*60,
        f"Tipo: {tipo} | Código: {codigo} | Versão: {versao}", "",
        f"--- 📏 MARGENS (Esperado: Sup=3,0 / Inf=2,0 / Esq=3,0 / Dir=2,0 cm) ---",
        f"Superior: {m['sup']} cm — {'✅' if m['ok_sup'] else '❌'}",
        f"Inferior: {m['inf']} cm — {'✅' if m['ok_inf'] else '❌'}",
        f"Esquerda: {m['esq']} cm — {'✅' if m['ok_esq'] else '❌'}",
        f"Direita: {m['dir']} cm — {'✅' if m['ok_dir'] else '❌'}",
        f"Margens: {'✅ TODAS CONFORME' if m['todas_ok'] else '❌ COM PENDÊNCIAS'}", "",
        f"--- ✍️ CORPO DO TEXTO (Esperado: {FONTE_CORPO} {TAMANHO_CORPO}) ---",
        f"Fontes: {', '.join(f'{n}({q})' for n,q in f['corpo']['fontes'].items())}",
        f"Tamanhos: {', '.join(f'{t}pt({q})' for t,q in f['corpo']['tamanhos'].items())}",
        f"Fonte: {f['corpo']['pct_fonte']}% — {'✅ CONFORME' if f['corpo']['fonte_ok'] else f'❌ ESPERADO: {FONTE_CORPO}'}",
        f"Tamanho: {f['corpo']['pct_tam']}% — {'✅ CONFORME' if f['corpo']['tam_ok'] else f'❌ ESPERADO: {TAMANHO_CORPO}pt'}", "",
        f"--- ✍️ TABELAS / FIGURAS (Esperado: {FONTE_TABELAS} {TAMANHO_TABELAS}) ---",
        f"Fontes: {', '.join(f'{n}({q})' for n,q in f['tabelas']['fontes'].items()) or 'Nenhuma'}",
        f"Tamanhos: {', '.join(f'{t}pt({q})' for t,q in f['tabelas']['tamanhos'].items()) or 'Nenhuma'}",
        f"Fonte: {f['tabelas']['pct_fonte']}% — {'✅ CONFORME' if f['tabelas']['fonte_ok'] else f'❌ ESPERADO: {FONTE_TABELAS}'}",
        f"Tamanho: {f['tabelas']['pct_tam']}% — {'✅ CONFORME' if f['tabelas']['tam_ok'] else f'❌ ESPERADO: {TAMANHO_TABELAS}pt'}",
        f"Registro Histórico: {'✅ ENCONTRADO' if f['tabelas']['tem_registro_historico'] else 'Não encontrado'}", "",
        f"--- 📊 TABELAS ({tb['total']} no total) ---",
        f"Tabela Cabeçalho: {'✅ ENCONTRADA' if tb['tem_cabecalho'] else '❌ NÃO ENCONTRADA'}"
    ]
    for tbinfo in tb["detalhes"]:
        marc = "📌 CABEÇALHO" if tbinfo["cab"] else "📋 REGISTRO HISTÓRICO" if tbinfo["hist"] else f"📊 Tabela {tbinfo['n']}"
        ficha.append(f"{marc}: {tbinfo['lin']} lin × {tbinfo['col']} col")
    ficha.extend(["", "--- 📋 SEÇÕES ENCONTRADAS ---"])
    for s in enc: ficha.append(f"✅ {s}")
    if falt:
        ficha.extend(["", "--- SEÇÕES FALTANTES ---"])
        for s in falt: ficha.append(f"❌ {s}")
    else:
        ficha.append("✅ TODAS AS SEÇÕES ENCONTRADAS")
    ficha.extend(["", "="*60, f"RESULTADO FINAL: {'✅ APROVADO' if aprov else '❌ NÃO CONFORME'}", "="*60])
    return "\n".join(ficha).encode("utf-8")

# ============================================================
# 🧠 ESCANEAR DOCUMENTO — ERRO CORRIGIDO!
# ============================================================
def escanear(doc_bytes):
    doc = docx.Document(BytesIO(doc_bytes))
    texto_completo = ""
    codigo = versao = validade = None
    
    # Ler tabelas (cabeçalho) — ✅ CORRIGIDO o erro 'NoneType' object has no attribute 'strip'
    for tb in doc.tables:
        for ln in tb.rows:
            lt = " ".join([c.text for c in ln.cells])
            texto_completo += lt + " "
            cod = re.search(r'CÓDIGO|Código[:\s]*[:]?\s*([A-Z]{2,5}[_\s]?[A-Z0-9]+)', lt, re.IGNORECASE)
            if cod and cod.group(1) and not codigo: 
                codigo = re.sub(r'\s+','_', cod.group(1).strip())
            ver = re.search(r'VERSÃO|Versão[:\s]*[:]?\s*(\d+)', lt, re.IGNORECASE)
            if ver and ver.group(1) and not versao: 
                versao = ver.group(1).strip()
            val = re.search(r'VALIDADE|Validade[:\s]*[:]?\s*([\d/]+)', lt, re.IGNORECASE)
            if val and val.group(1) and not validade: 
                validade = val.group(1).strip()
    
    # Ler corpo do texto
    for p in doc.paragraphs: texto_completo += p.text + " "
    texto_limpo = limpar_texto(texto_completo)
    
    # Detectar tipo de documento
    tipo = "PROT"
    if re.search(r'\bPROTOCOLO\b', texto_limpo): tipo = "PROT"
    elif re.search(r'\bPOP\b', texto_limpo): tipo = "POP"
    elif re.search(r'\bPOLÍTICA|POLITICA|POI\b', texto_limpo): tipo = "POI"
    elif re.search(r'\bNORMA|NOR\b', texto_limpo): tipo = "NOR"
    elif re.search(r'\bREGIMENTO|REG\b', texto_limpo): tipo = "REG"
    elif re.search(r'\bPROGRAMA|PROG\b', texto_limpo): tipo = "PROG"
    elif re.search(r'\bPLANO|PLAN\b', texto_limpo): tipo = "PLAN"
    elif re.search(r'\bROTINA|ROT\b', texto_limpo): tipo = "ROT"
    elif re.search(r'\bMANUAL|MAN\b', texto_limpo): tipo = "MAN"
    
    # Verificar seções
    secoes_esp = SECOES_POR_TIPO[tipo]
    enc = [s for s in secoes_esp if limpar_texto(s) in texto_limpo]
    falt = [s for s in secoes_esp if limpar_texto(s) not in texto_limpo]
    
    return {
        "tipo": tipo, "codigo": codigo, "versao": versao, "validade": validade,
        "margens": verificar_margens(doc),
        "fonte": verificar_fonte_separado(doc),
        "tabelas": listar_tabelas(doc),
        "secoes_enc": enc, "secoes_falt": falt
    }

# ============================================================
# 🚀 INTERFACE PRINCIPAL — SEU NOME, TÍTULO E BONEQUINHO
# ============================================================

# Cabeçalho com título, nome e bonequinho no canto direito
st.markdown("""
    <style>
    .header-container {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background-color: #0F172A;
        padding: 15px 25px;
        border-radius: 12px;
        margin-bottom: 20px;
    }
    .header-text h1 {
        color: #FFFFFF;
        margin: 0;
        font-size: 28px;
        font-weight: bold;
    }
    .header-text p {
        color: #94A3B8;
        margin: 5px 0 0 0;
        font-size: 15px;
    }
    .header-emoji {
        font-size: 50px;
    }
    </style>
    <div class="header-container">
        <div class="header-text">
            <h1>AUDITOR NAQH NMZ DE ALTA PRECISÃO</h1>
            <p>Ezequias Santos — Agente Administrativo | Margens: Esq 3,0 / Dir 2,0 / Sup 3,0 / Inf 2,0 cm • Corpo Calibri 11 • Tabelas Calibri 10</p>
        </div>
        <div class="header-emoji">👨‍💻</div>
    </div>
""", unsafe_allow_html=True)

# Upload do arquivo
arquivo = st.file_uploader("📂 Envie o documento (.docx)", type=["docx"])

if arquivo:
    st.info(f"✅ Arquivo carregado: **{arquivo.name}**")
    with st.spinner("Verificando conforme Norma Zero..."):
        try:
            dados = arquivo.read()
            r = escanear(dados)
            
            st.markdown("---")
            st.subheader("📋 DADOS DO DOCUMENTO")
            c1, c2 = st.columns(2)
            with c1: st.info(f"**Tipo:** {r['tipo']}")
            with c2: st.success(f"**Código:** {r['codigo'] or '❌ NÃO ENCONTRADO'}")
            st.success(f"**Versão:** {r['versao'] or '❌ NÃO ENCONTRADA'}")
            
            st.markdown("---")
            st.subheader("📏 MARGENS — Esperado: Sup=3,0 / Inf=2,0 / Esq=3,0 / Dir=2,0 cm")
            m = r["margens"]
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Superior", f"{m['sup']} cm", "✅" if m["ok_sup"] else "❌")
            col2.metric("Inferior", f"{m['inf']} cm", "✅" if m["ok_inf"] else "❌")
            col3.metric("Esquerda", f"{m['esq']} cm", "✅" if m["ok_esq"] else "❌")
            col4.metric("Direita", f"{m['dir']} cm", "✅" if m["ok_dir"] else "❌")
            st.success("✅ TODAS AS MARGENS CONFORME") if m["todas_ok"] else st.error("❌ MARGENS NÃO CONFORME")
            
            st.markdown("---")
            st.subheader(f"✍️ CORPO DO TEXTO — Esperado: Calibri 11pt")
            fc = r["fonte"]["corpo"]
            st.write("**Fontes encontradas:**", ", ".join([f"{n} ({q})" for n,q in fc["fontes"].items()]))
            st.write("**Tamanhos encontrados:**", ", ".join([f"{t}pt ({q})" for t,q in fc["tamanhos"].items()]))
            col1, col2 = st.columns(2)
            with col1: st.metric("Fonte Calibri", f"{fc['pct_fonte']}%", "✅ CONFORME" if fc["fonte_ok"] else "❌ DIVERGENTE")
            with col2: st.metric("Tamanho 11pt", f"{fc['pct_tam']}%", "✅ CONFORME" if fc["tam_ok"] else "❌ DIVERGENTE")
            
            st.markdown("---")
            st.subheader(f"✍️ TABELAS / FIGURAS — Esperado: Calibri 10pt")
            ft = r["fonte"]["tabelas"]
            if ft["fontes"]:
                st.write("**Fontes nas Tabelas:**", ", ".join([f"{n} ({q})" for n,q in ft["fontes"].items()]))
                st.write("**Tamanhos nas Tabelas:**", ", ".join([f"{t}pt ({q})" for t,q in ft["tamanhos"].items()]))
                col1, col2 = st.columns(2)
                with col1: st.metric("Fonte Calibri", f"{ft['pct_fonte']}%", "✅ CONFORME" if ft["fonte_ok"] else "⚠️ DIFERENTE")
                with col2: st.metric("Tamanho 10pt", f"{ft['pct_tam']}%", "✅ CONFORME" if ft["tam_ok"] else "⚠️ DIFERENTE")
                if ft["tem_registro_historico"]:
                    st.info("📋 **Tabela de REGISTRO HISTÓRICO detectada** — verifique fonte/tamanho")
            else:
                st.info("Nenhuma tabela com texto detectada")
            
            st.markdown("---")
            st.subheader("📊 TABELAS DO DOCUMENTO")
            tb = r["tabelas"]
            st.write(f"**Total de tabelas:** {tb['total']}")
            st.success("✅ Tabela do Cabeçalho encontrada") if tb["tem_cabecalho"] else st.warning("⚠️ Tabela do Cabeçalho NÃO encontrada")
            for tbinfo in tb["detalhes"]:
                tipo = "📌 CABEÇALHO" if tbinfo["cab"] else "📋 REGISTRO HISTÓRICO" if tbinfo["hist"] else f"📊 Tabela {tbinfo['n']}"
                st.write(f"{tipo}: {tbinfo['lin']} linhas × {tbinfo['col']} colunas")
            
            st.markdown("---")
            st.subheader("📋 SEÇÕES")
            for s in r["secoes_enc"]: st.success(f"✅ {s}")
            if r["secoes_falt"]:
                for s in r["secoes_falt"]: st.error(f"❌ {s}")
            else:
                st.success("✅ TODAS AS SEÇÕES ENCONTRADAS")
            
            st.markdown("---")
            aprov = m["todas_ok"] and r["codigo"] and r["versao"] and not r["secoes_falt"]
            if aprov:
                st.success("🎉 **DOCUMENTO APROVADO CONFORME NORMA ZERO!**")
                st.balloons()
            else:
                st.warning("⚠️ **DOCUMENTO COM PENDÊNCIAS** — verifique os itens acima")
            
            dados_formatados = aplicar_margens(dados)
            ficha_bytes = gerar_ficha(r["tipo"], r["codigo"], r["versao"], m, r["fonte"], tb, r["secoes_enc"], r["secoes_falt"], aprov)
            
            nome_doc = f"{r['codigo']}_Formatado.docx" if r['codigo'] else "Documento_Formatado.docx"
            nome_txt = f"Ficha_Verificacao_{r['codigo']}.txt" if r['codigo'] else "Ficha_Verificacao.txt"
            
            st.download_button("📥 DOWNLOAD — Documento formatado (.DOCX)", dados_formatados, nome_doc, 
                              "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.download_button("📄 DOWNLOAD — Ficha de verificação (.TXT)", ficha_bytes, nome_txt, "text/plain")
        
        except Exception as e:
            st.error(f"## ❌ ERRO: {str(e)}")
