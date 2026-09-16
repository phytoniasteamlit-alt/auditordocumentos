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
    "PLAN": {
        "obrigatorias": [
            "OBJETIVO", "APLICABILIDADE", "DEFINIÇÃO DE TERMOS",
            "IDENTIFICAÇÃO DE RISCO ATUAL", "MEDIDAS DE CONTINGÊNCIA", "REFERÊNCIAS"
        ],
        "opcionais": ["APÊNDICES", "ANEXOS"]
    },
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

# Extrair do NOME do arquivo
def extrair_do_nome(nome_arquivo):
    dados = {"tipo": None, "codigo": None, "versao": None}
    nome_limpo = limpar_texto(nome_arquivo)

    mapeia_nome = [
        ("POP", r'\bPOP\b'), ("POI", r'\bPOI\b'), ("NOR", r'\bNOR\b'),
        ("PLAN", r'\bPLAN\b'), ("PROT", r'\bPROT\b'), ("REG", r'\bREG\b'),
        ("PROG", r'\bPROG\b'), ("ROT", r'\bROT\b'), ("MAN", r'\bMAN\b'),
    ]
    for t, p in mapeia_nome:
        if re.search(p, nome_limpo):
            dados["tipo"] = t
            break

    m = re.search(r'\b([A-Z]{2,}\d+)\b', nome_limpo)
    if m: dados["codigo"] = m.group(1)

    padroes_ver = [
        r'(\d+[ºª]\s*(?:[a-zA-ZÀ-ÿ\s]+)?)',
        r'V(\d+(?:[.\-]\d+)*)',
        r'VERS[AO]?\s*(\d+(?:[.\-]\d+)*)',
        r'(\d+)\s*[ªº]'
    ]
    for p in padroes_ver:
        m = re.search(p, nome_arquivo.upper())
        if m:
            dados["versao"] = m.group(1).strip()
            break
    return dados

# Extrair dados de um texto
def extrair_de_texto(texto):
    dados = {"codigo": None, "versao": None}
    if not texto: return dados
    tx_upper = texto.upper()

    m = re.search(r'C[ÓO]DIGO\s*[:：=\-]?\s*([A-Z0-9_\-\/\.]+)', tx_upper)
    if m: dados["codigo"] = m.group(1).strip()

    m = re.search(r'VERS[AÃ]O\s*[:：=\-]?\s*(\d+[ºª]?\d*[ºª]?|\d+(?:[.\-]\d+)*)', tx_upper)
    if m: dados["versao"] = m.group(1).strip()

    m = re.search(r'VERSÃO\s*(\d+[ºª])', texto)
    if m and not dados["versao"]:
        dados["versao"] = m.group(1).strip()

    return dados

# Ler TODO o conteúdo: Cabeçalho + Corpo + Tabelas + Rodapé
def ler_conteudo_completo(doc):
    texto_completo = ""
    dados_encontrados = {"codigo": None, "versao": None}

    # CABEÇALHO
    for secao in doc.sections:
        cab = secao.header
        for paragrafo in cab.paragraphs:
            texto_completo += paragrafo.text + " "
            d = extrair_de_texto(paragrafo.text)
            if d["codigo"] and not dados_encontrados["codigo"]:
                dados_encontrados["codigo"] = d["codigo"]
            if d["versao"] and not dados_encontrados["versao"]:
                dados_encontrados["versao"] = d["versao"]
        for tabela in cab.tables:
            for linha in tabela.rows:
                texto_cel = " ".join([cel.text for cel in linha.cells])
                texto_completo += texto_cel + " "
                d = extrair_de_texto(texto_cel)
                if d["codigo"] and not dados_encontrados["codigo"]:
                    dados_encontrados["codigo"] = d["codigo"]
                if d["versao"] and not dados_encontrados["versao"]:
                    dados_encontrados["versao"] = d["versao"]

    # CORPO
    for p in doc.paragraphs:
        texto_completo += p.text + " "
        d = extrair_de_texto(p.text)
        if d["codigo"] and not dados_encontrados["codigo"]:
            dados_encontrados["codigo"] = d["codigo"]
        if d["versao"] and not dados_encontrados["versao"]:
            dados_encontrados["versao"] = d["versao"]

    # TABELAS DO CORPO
    for tb in doc.tables:
        for ln in tb.rows:
            texto_cel = " ".join([cel.text for cel in ln.cells])
            texto_completo += texto_cel + " "
            d = extrair_de_texto(texto_cel)
            if d["codigo"] and not dados_encontrados["codigo"]:
                dados_encontrados["codigo"] = d["codigo"]
            if d["versao"] and not dados_encontrados["versao"]:
                dados_encontrados["versao"] = d["versao"]

    # RODAPÉ
    for secao in doc.sections:
        rod = secao.footer
        for paragrafo in rod.paragraphs:
            texto_completo += paragrafo.text + " "
            d = extrair_de_texto(paragrafo.text)
            if d["codigo"] and not dados_encontrados["codigo"]:
                dados_encontrados["codigo"] = d["codigo"]
            if d["versao"] and not dados_encontrados["versao"]:
                dados_encontrados["versao"] = d["versao"]
        for tabela in rod.tables:
            for linha in tabela.rows:
                texto_cel = " ".join([cel.text for cel in linha.cells])
                texto_completo += texto_cel + " "
                d = extrair_de_texto(texto_cel)
                if d["codigo"] and not dados_encontrados["codigo"]:
                    dados_encontrados["codigo"] = d["codigo"]
                if d["versao"] and not dados_encontrados["versao"]:
                    dados_encontrados["versao"] = d["versao"]

    return texto_completo, dados_encontrados

# ============================================================
# ✍️ VERIFICAR FONTE
# ============================================================
def verificar_fonte(doc):
    cont = {"total":0, "fonte_ok":0, "tam_ok":0, "fontes":{}, "tams":{}}
    def cont_run(run, fonte_padrao, tam_padrao):
        nonlocal cont
        cont["total"] += 1
        nome = run.font.name or fonte_padrao
        tam_pt = round(run.font.size.pt, 1) if run.font.size else tam_padrao
        cont["fontes"][nome] = cont["fontes"].get(nome, 0) + 1
        cont["tams"][tam_pt] = cont["tams"].get(tam_pt, 0) + 1
        if nome == fonte_padrao: cont["fonte_ok"] += 1
        if tam_pt == tam_padrao: cont["tam_ok"] += 1

    for secao in doc.sections:
        for p in secao.header.paragraphs:
            for r in p.runs: cont_run(r, FONTE_CORPO, TAMANHO_CORPO)
        for p in secao.footer.paragraphs:
            for r in p.runs: cont_run(r, FONTE_CORPO, TAMANHO_CORPO)

    for p in doc.paragraphs:
        for r in p.runs: cont_run(r, FONTE_CORPO, TAMANHO_CORPO)
    for tb in doc.tables:
        for ln in tb.rows:
            for cel in ln.cells:
                for p in cel.paragraphs:
                    for r in p.runs: cont_run(r, FONTE_TABELAS, TAMANHO_TABELAS)

    pct_f = calcular_pct(cont["fonte_ok"], cont["total"])
    pct_t = calcular_pct(cont["tam_ok"], cont["total"])
    return {
        "fontes": cont["fontes"], "tamanhos": cont["tams"],
        "pct_fonte": pct_f, "pct_tam": pct_t,
        "fonte_ok": pct_f >= LIMIAR_CONFORMIDADE,
        "tam_ok": pct_t >= LIMIAR_CONFORMIDADE,
    }

# ============================================================
# 🔍 ESCANEAR
# ============================================================
def escanear(doc_bytes):
    doc = docx.Document(BytesIO(doc_bytes))
    texto_completo, dados_doc = ler_conteudo_completo(doc)
    texto_limpo = limpar_texto(texto_completo)

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

    secoes = SECOES_POR_TIPO.get(tipo, SECOES_POR_TIPO["PROT"])
    obr_enc, obr_falt, opc_enc, opc_falt = [], [], [], []
    for s in secoes["obrigatorias"]:
        if limpar_texto(s) in texto_limpo: obr_enc.append(s)
        else: obr_falt.append(s)
    for s in secoes["opcionais"]:
        if limpar_texto(s) in texto_limpo: opc_enc.append(s)
        else: opc_falt.append(s)

    return {
        "tipo": tipo,
        "codigo_doc": dados_doc["codigo"],
        "versao_doc": dados_doc["versao"],
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
# 📄 GERAR FICHA — CORRIGIDO ✅
# ============================================================
def gerar_ficha(nome_arq, tipo, cod, ver, fonte_ok, secoes_dados, itens):
    obr_enc, obr_falt, opc_enc, opc_falt = secoes_dados
    status = "✅ SIM" if (fonte_ok and len(obr_falt)==0 and all(itens.values())) else "⚠️ COM PENDÊNCIAS"
    
    linhas_itens = "\n".join(f"{k}: {'✅' if v else '❌'}" for k, v in itens.items())
    linhas_obrig = "\n".join(f"✅ {s}" for s in obr_enc)
    if obr_falt:
        linhas_obrig += "\n" + "\n".join(f"❌ {s}" for s in obr_falt)

    ficha_texto = f"""==================================================
          FICHA DE VERIFICAÇÃO — NAQH NMZ
==================================================
Arquivo: {nome_arq}
Tipo: {tipo}
--------------------------------------------------
Código:     {cod or 'NÃO INFORMADO'}
Versão:     {ver or 'NÃO INFORMADO'}
--------------------------------------------------
MARGENS: Esq {MARGEM_ESQ_ESPERADA} / Dir {MARGEM_DIR_ESPERADA} / Sup {MARGEM_SUP_ESPERADA} / Inf {MARGEM_INF_ESPERADA} cm ✅
FONTE: {FONTE_CORPO} {TAMANHO_CORPO}pt / {FONTE_TABELAS} {TAMANHO_TABELAS}pt → {'✅ CONFORME' if fonte_ok else '❌ NÃO CONFORME'}
--------------------------------------------------
CABEÇALHO CONFERIDO
--------------------------------------------------
{linhas_itens}
--------------------------------------------------
SEÇÕES OBRIGATÓRIAS: {len(obr_enc)}/{len(obr_enc)+len(obr_falt)}
--------------------------------------------------
{linhas_obrig}
--------------------------------------------------
APROVADO CONFORME NORMA ZERO: {status}
==================================================
Ezequias Santos — Agente Administrativo
Auditor NAQH NMZ
==================================================
"""
    return ficha_texto.encode("utf-8")

# ============================================================
# 🚀 INTERFACE PRINCIPAL
# ============================================================
st.markdown("""
    <style>
    .header{{display:flex;gap:20px;background:#0F172A;padding:20px;border-radius:12px;margin-bottom:20px}}
    .header h1{{color:#FFF;margin:0}}
    .relogio{{background:linear-gradient(135deg,#0F766E,#14B8A6);padding:20px;border-radius:12px;color:#FFF;margin:15px 0}}
    </style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header">
    <div>
        <h1>👨‍💻 AUDITOR NAQH NMZ DE ALTA PRECISÃO</h1>
        <p style="color:#94A3B8;margin:5px 0 0 0;">Leitura de Cabeçalho + Corpo + Rodapé • Erro de concatenação corrigido ✅</p>
    </div>
</div>
""", unsafe_allow_html=True)

arquivos = st.file_uploader("📂 Envie o(s) documento(s) (.docx)", type=["docx"], accept_multiple_files=True)

if arquivos:
    qtd = len(arquivos)
    st.markdown(f"""
    <div class="relogio">
        ⏱️ ECONOMIA DE TEMPO — {qtd} documento(s)<br>
        📝 Manual: {qtd*TEMPO_POR_DOC_MANUAL}min | ⚡ Auditor: {qtd*TEMPO_POR_DOC_AUTOMATICO}min
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    for idx, arq in enumerate(arquivos, 1):
        st.subheader(f"📄 {arq.name}")
        try:
            dados = arq.read()
            r = escanear(dados)
            nome = extrair_do_nome(arq.name)

            # Definir tipo
            tipo_final = nome["tipo"] or r["tipo"]
            if nome["tipo"]:
                st.success(f"✅ Tipo pelo nome: **{tipo_final}**")
            else:
                st.info(f"ℹ️ Tipo pelo conteúdo: **{tipo_final}**")

            st.markdown("### 📋 DADOS DETECTADOS")

            # CÓDIGO
            fonte_cod = []
            if r["codigo_doc"]: fonte_cod.append(f"Cabeçalho: {r['codigo_doc']}")
            if nome["codigo"]: fonte_cod.append(f"Nome: {nome['codigo']}")
            valor_cod = r["codigo_doc"] or nome["codigo"] or ""
            codigo_final = st.text_input("**CÓDIGO**", value=valor_cod,
                placeholder="Digite o código", key=f"cod{idx}")
            if fonte_cod:
                st.info(f"📰 Fontes: {' | '.join(fonte_cod)}")

            # VERSÃO
            fonte_ver = []
            if r["versao_doc"]: fonte_ver.append(f"**Cabeçalho: {r['versao_doc']}**")
            if nome["versao"]: fonte_ver.append(f"Nome: {nome['versao']}")
            valor_ver = r["versao_doc"] or nome["versao"] or ""
            versao_final = st.text_input("**VERSÃO / SÉRIE**", value=valor_ver,
                placeholder="Digite a versão", key=f"ver{idx}")
            if fonte_ver:
                st.info(f"📰 Fontes: {' | '.join(fonte_ver)}")

            st.markdown("---")

            # FONTE
            f = r["fonte"]
            st.markdown(f"### ✍️ FONTE — {FONTE_CORPO} {TAMANHO_CORPO}pt / {FONTE_TABELAS} {TAMANHO_TABELAS}pt")
            st.write("Fontes:", ", ".join(f"{n} ({q})" for n, q in f["fontes"].items()) or "Nenhuma")
            st.write("Tamanhos:", ", ".join(f"{t}pt ({q})" for t, q in f["tamanhos"].items()) or "Nenhum")
            fonte_ok = f["fonte_ok"] and f["tam_ok"]
            st.metric("Conformidade", f"{f['pct_fonte']}%", "✅" if fonte_ok else "❌")

            st.markdown("---")

            # CONFERÊNCIA MANUAL
            st.markdown("### ✅ CONFERÊNCIA MANUAL — CABEÇALHO")
            itens = {
                "Cabeçalho padrão": st.checkbox("✅ Cabeçalho", key=f"cab{idx}"),
                "Número de páginas": st.checkbox("✅ Páginas", key=f"pag{idx}"),
                "Versão/Série no cabeçalho": st.checkbox("✅ Versão no cabeçalho", key=f"verif{idx}"),
                "Logo do Hospital": st.checkbox("✅ Logo", key=f"logo{idx}"),
                "Marca d'água": st.checkbox("✅ Marca d'água", key=f"marca{idx}"),
            }

            st.markdown("---")

            # SEÇÕES
            st.markdown(f"### 📑 SEÇÕES — Tipo: {tipo_final}")
            secoes = SECOES_POR_TIPO[tipo_final]
            doc_temp = docx.Document(BytesIO(dados))
            tx_limpo = limpar_texto(ler_conteudo_completo(doc_temp)[0])
            
            obr_enc, obr_falt, opc_enc, opc_falt = [], [], [], []
            for s in secoes["obrigatorias"]:
                if limpar_texto(s) in tx_limpo: obr_enc.append(s)
                else: obr_falt.append(s)
            for s in secoes["opcionais"]:
                if limpar_texto(s) in tx_limpo: opc_enc.append(s)
                else: opc_falt.append(s)

            for s in obr_enc: st.success(f"✅ {s}")
            secoes_ok = True
            for s in obr_falt:
                marc = st.checkbox(f"⚠️ {s} — Marque se encontrou", key=f"sf{idx}_{limpar_texto(s)}")
                if not marc: secoes_ok = False
            if opc_enc:
                st.markdown("**🟡 Opcionais:**")
                for s in opc_enc: st.info(f"🟡 {s}")

            st.markdown("---")

            cab_ok = all(itens.values())
            aprov = bool(codigo_final and versao_final and secoes_ok and cab_ok and fonte_ok)
            if aprov:
                st.success("✅ APROVADO CONFORME NORMA ZERO")
            else:
                st.warning("⚠️ COM PENDÊNCIAS — verifique itens acima")

            # DOWNLOADS
            cod_nome = re.sub(r'[<>:"/\\|?*º°]', '-', codigo_final.strip()) if codigo_final else "DOC"
            ver_nome = re.sub(r'[<>:"/\\|?*º°]', '_', versao_final.strip()) if versao_final else "0"
            base = f"{cod_nome}_v{ver_nome}"

            doc_processado = aplicar_margens(dados)
            st.download_button(
                f"📥 BAIXAR: {base}.docx",
                doc_processado,
                f"{base}.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_doc_{idx}",
                type="primary"
            )

            ficha = gerar_ficha(arq.name, tipo_final, codigo_final, versao_final,
                fonte_ok, (obr_enc, obr_falt, opc_enc, opc_falt), itens)
            st.download_button(
                f"📋 BAIXAR FICHA: {base}_FICHA.txt",
                ficha,
                f"{base}_FICHA.txt",
                "text/plain",
                key=f"dl_ficha_{idx}"
            )

        except Exception as e:
            st.error(f"❌ Erro ao processar {arq.name}: {str(e)}")
        st.markdown("---")
