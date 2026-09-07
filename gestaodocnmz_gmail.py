# ==================================================
# 🛡️ REGRA DE OURO — CÓDIGO + VERSÃO = CHAVE ÚNICA
# ==================================================
# SEM CÓDIGO OU SEM VERSÃO → NÃO ENTRA NO FLUXO → DEVOLVE
# CÓDIGO + VERSÃO IGUAL → MESMO DOCUMENTO MESMA VERSÃO → ATUALIZA APENAS VAZIOS
# CÓDIGO IGUAL + VERSÃO DIFERENTE → DOCUMENTO NOVO → LINHA NOVA
# NADA É APAGADO. TODAS AS VERSÕES SÃO PRESERVADAS.
# ==================================================

def validar_dados_obrigatorios(codigo, versao, data_aprovacao=None):
    """
    VALIDAÇÃO OBRIGATÓRIA antes de qualquer coisa
    Retorna: (valido: bool, motivo: str)
    """
    # Regra Dourada: SEM CÓDIGO → NÃO ENTRA
    if not codigo or str(codigo).strip().upper() in ["", "NAN", "NONE"]:
        return False, "❌ DEVOLVIDO: SEM CÓDIGO — Não pode iniciar verificação"
    
    # Regra Dourada: SEM VERSÃO → NÃO ENTRA
    if not versao or str(versao).strip() in ["", "NAN", "NONE"]:
        return False, "❌ DEVOLVIDO: SEM VERSÃO — Não pode iniciar verificação"
    
    # Tem os dois → PODE ENTRAR
    return True, "✅ VÁLIDO: Código + Versão presentes"


def criar_chave_unica(codigo, versao):
    """Cria a identidade única do documento"""
    cod = str(codigo).strip().upper().replace(" ", "")
    ver = str(versao).strip().replace(" ", "")
    return f"{cod}||VERSAO:{ver}"


def auditoria_comparativa(df_planilha, lista_emails):
    # ── PASSO 1: Construir índice de TODOS os documentos já registrados ──
    existentes = {}  # { CHAVE_UNICA: numero_linha }
    ultimas_versoes = {}  # { CODIGO: ULTIMA_VERSAO }

    for idx, linha in df_planilha.iterrows():
        cod = str(linha.get("Código", "")).strip().upper()
        ver = str(linha.get("Versão", "")).strip()
        
        if cod and ver and cod != "NAN":
            chave = criar_chave_unica(cod, ver)
            existentes[chave] = idx
            
            # Rastrear última versão de cada código
            if cod not in ultimas_versoes:
                ultimas_versoes[cod] = []
            ultimas_versoes[cod].append(int(ver) if ver.isdigit() else ver)

    # ── PASSO 2: Listas de relatórios ──
    novas_linhas = []
    devolvidos_sem_dados = []
    relatorio_faltantes = []
    relatorio_parados = []
    relatorio_nao_devolvidos = []

    data_hoje = datetime.now()

    # ── PASSO 3: Processar cada e-mail recebido ──
    for email in lista_emails:
        assunto = email.get("assunto", "")
        data_email = email.get("data_recebimento")
        corpo = email.get("corpo", "")
        anexos = email.get("anexos", [])

        codigo_doc, versao_doc, data_aprov_doc = None, None, None

        # Extrair de cada anexo (Word/PDF)
        for nome_arq, conteudo in anexos:
            if nome_arq.lower().endswith(".docx"):
                c, v, d = extrair_dados_docx(conteudo)
            elif nome_arq.lower().endswith(".pdf"):
                c, v, d = extrair_dados_pdf(conteudo)
            else:
                continue
            if c: codigo_doc = c
            if v: versao_doc = v
            if d: data_aprov_doc = d

        # ==================================================
        # 🔴 REGRA DE OURO — VALIDAÇÃO OBRIGATÓRIA
        # ==================================================
        valido, motivo = validar_dados_obrigatorios(codigo_doc, versao_doc, data_aprov_doc)
        
        if not valido:
            # ❌ SEM CÓDIGO OU SEM VERSÃO → DEVOLVE → NÃO ENTRA NA PLANILHA
            devolvidos_sem_dados.append({
                "Assunto": assunto,
                "Data Recebimento": data_email.strftime("%d/%m/%Y") if data_email else "Desconhecida",
                "Motivo": motivo,
                "Código recebido": codigo_doc or "---",
                "Versão recebida": versao_doc or "---"
            })
            continue  # ⛔ PARA AQUI. NÃO FAZ MAIS NADA.

        # ✅ TEM CÓDIGO + VERSÃO → PODE SEGUIR
        chave_unica = criar_chave_unica(codigo_doc, versao_doc)

        # ── CASO A: Documento JÁ EXISTE na planilha → SÓ PREENCHE VAZIOS ──
        if chave_unica in existentes:
            # Já está registrado → NÃO CRIAR NOVA LINHA
            # O sistema só vai preencher as células vazias dessa linha existente
            continue

        # ── CASO B: Documento NOVO — mesma versão? Não, nova versão! ──
        # Mesmo Código, Versão Diferente = DOCUMENTO NOVO → LINHA NOVA
        if codigo_doc in [ch.split("||")[0] for ch in existentes.keys()]:
            # É uma NOVA VERSÃO de um documento já existente
            status = "NOVA VERSÃO DETECTADA"
        else:
            # É um DOCUMENTO TOTALMENTE NOVO
            status = "NOVO DOCUMENTO"

        # ── Verificar se veio APROVADO direto ──
        if data_aprov_doc:
            relatorio_faltantes.append({
                "Código": codigo_doc,
                "Versão": versao_doc,
                "Data Aprovação": data_aprov_doc,
                "Status": f"✅ APROVADO — {status}",
                "Data Recebimento": data_email.strftime("%d/%m/%Y") if data_email else "---"
            })
            novas_linhas.append({
                "Código": codigo_doc,
                "Versão": versao_doc,
                "Data de Aprovação": data_aprov_doc,
                "Status": "APROVADO",
                "Data de Recebimento": data_email.strftime("%d/%m/%Y") if data_email else None
            })
            continue

        # ── CASO C: Documento parado sem análise ──
        if data_email and (data_hoje - data_email).days > PRAZO_DOC_PARADO:
            relatorio_parados.append({
                "Código": codigo_doc,
                "Versão": versao_doc,
                "Dias parado sem análise": (data_hoje - data_email).days,
                "Data Recebimento": data_email.strftime("%d/%m/%Y"),
                "Assunto": assunto
            })

        # ── CASO D: Enviado ao Setor sem devolução ──
        if "encaminhado ao setor" in corpo.lower() or "devolvido ao setor" in corpo.lower():
            if data_email and (data_hoje - data_email).days > PRAZO_ALERTA_NAO_DEVOLVIDO:
                relatorio_nao_devolvidos.append({
                    "Código": codigo_doc,
                    "Versão": versao_doc,
                    "Dias sem devolução": (data_hoje - data_email).days,
                    "Data envio ao setor": data_email.strftime("%d/%m/%Y")
                })

    # ==================================================
    # ✅ RETORNA TUDO SEPARADO
    # ==================================================
    return {
        "novas_linhas": novas_linhas,
        "devolvidos_sem_dados": devolvidos_sem_dados,  # ❌ Sem código/versão
        "faltantes_aprovados": relatorio_faltantes,     # ✅ Aprovados que faltavam
        "parados_sem_analise": relatorio_parados,       # ⚠️ Parados
        "sem_devolucao_setor": relatorio_nao_devolvidos  # ⏹ Setor não devolveu
    }
