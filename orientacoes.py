import streamlit as st
import docx 
import pandas as pd
import re
import datetime
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador NAQH", page_icon="🔍", layout="wide")
st.title("Auditor Automatizado - Norma Zero (V44 - Corrigida)")
st.markdown("Este sistema realiza a auditoria técnica rigorosa e gera a Ficha de Verificação oficial baseada na Norma Zero.")

# Inicialização de estados de sessão seguros
if "cached_nome_arquivo" not in st.session_state:
    st.session_state.cached_nome_arquivo = "Documento Coletado"
if "cached_tipo" not in st.session_state:
    st.session_state.cached_tipo = "PROT"
if "cached_porcentagem" not in st.session_state:
    st.session_state.cached_porcentagem = 0

#--- 2. BARRA LATERAL (CONTROLES E DOWNLOAD FIXO) ---
st.sidebar.header("⚙️ Controles de Auditoria")
logo_prefeitura_manual = st.sidebar.checkbox("Autorizar Manualmente: Logomarca do Hospital")
marca_dagua_manual = st.sidebar.checkbox("Autorizar Aprovação da Marca d'água")
autorizar_codigo_manual = st.sidebar.checkbox("Autorizar Manualmente: Código do Documento")
autorizar_fonte_manual = st.sidebar.checkbox("Autorizar Manualmente: Fonte e Tamanho")
autorizar_etapas_manual = st.sidebar.checkbox("Autorizar Manualmente: Descrição das Etapas")

#--- 3. DICIONÁRIO DE REGRAS UNIFICADO ---
REGRAS_NORMA_ZERO = {
    "NOR": {
        "nome": "Norma",
        "obrigatorios": ["INTRODUÇÃO", "OBJETIVO", "APLICABILIDADE", "DESCRIÇÃO DA NORMA", "RESPONSÁVEL", "EFEITOS DO NÃO CUMPRIMENTO", "REFERÊNCIA", "APÊNDICES", "ANEXOS"]
    },
    "POP": {
        "nome": "Procedimento Operacional Padrão",
        "obrigatorios": ["DEFINIÇÃO", "APLICABILIDADE", "RESPONSÁVEL PELA EXECUÇÃO", "MATERIAIS UTILIZADOS", "DESCRIÇÃO DAS ETAPAS", "REFERÊNCIAS"]
    },
    "PROT": {
        "nome": "Protocolo",
        "obrigatorios": ["INTRODUÇÃO", "OBJETIVO", "APLICABILIDADE", "CONTEÚDO CLÍNICO", "RESPONSÁVEL", "REFERÊNCIAS"]
    },
    "MAN": {"nome": "Manual", "obrigatorios": ["INTRODUÇÃO", "CONTEÚDO", "REFERÊNCIAS"]},
    "PLANC": {"nome": "Plano de Contingência", "obrigatorios": ["INTRODUÇÃO", "AÇÕES", "RESPONSÁVEIS"]},
    "POL": {"nome": "Política Institucional", "obrigatorios": ["DIRETRIZES", "OBJETIVOS"]},
    "PROG": {"nome": "Programa", "obrigatorios": ["CRONOGRAMA", "METAS"]},
    "REG": {"nome": "Regimento Interno", "obrigatorios": ["DISPOSIÇÕES GERAIS", "COMPETÊNCIAS"]},
    "ROT": {"nome": "Rotina", "obrigatorios": ["ATIVIDADES", "FLUXO"]}
}

tipo_detectado = st.session_state.cached_tipo
has_apendices_ou_anexos = False
marca_dagua_corpo_ok = marca_dagua_manual
cabecalho_completo = True
codigo_detectado_no_texto = False
registro_historico_ok = False
campos_nota_rh_ok = False
nome_arquivo_doc = st.session_state.cached_nome_arquivo

# Flags de verificação estrita de formatação (Pente Fino)
margens_ok = True
papel_a4_ok = True
fonte_e_tamanho_ok = True
espacamento_linhas_ok = True
alinhamento_ok = True
paragrafo_ok = True
paginacao_ok = True

itens_ficha = [
    "Logomarca do Hospital", "Título do Documento", "Tipo de Documento", "Código do Documento",
    "Versão", "Páginas", "Data da Aprovação / Validade", "Registro Histórico", "Campos Adicionais RH", "Papel",
    "Margens", "Fonte e Tamanho", "Espaçamento entre Linhas", "Alinhamento", "Parágrafo",
    "Figuras/Tabelas", "Paginação", "Marca d'água", "Cabeçalho das Páginas", "Referências", "Apêndices / Anexos",
    "Impressos Assistenciais/Administrativos"
]

itens_estrutura_dinamica = []
valores_estrutura_dinamica = []

#--- 4. FLUXO DE CARREGAMENTO DE ARQUIVOS (WORD.DOCX) ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para auditoria", type=["docx"])

if arquivo_word:
    nome_arquivo_doc = arquivo_word.name
    st.session_state.cached_nome_arquivo = nome_arquivo_doc
    
    for sigla in REGRAS_NORMA_ZERO.keys():
        if f"{sigla}_" in arquivo_word.name.upper() or f"{sigla}" in arquivo_word.name.upper():
            tipo_detectado = sigla
            break
    st.session_state.cached_tipo = tipo_detectado
    
    doc = docx.Document(arquivo_word)
    conteudo_linhas = []
    
    # --- PENTE FINO DE CONFIGURAÇÃO DE PÁGINA (MARGENS E PAPEL) ---
    for secao in doc.sections:
        # Conversão de unidades nativas do Word para centímetros (1 cm = 360000 EMUs)
        margin_top = round(secao.top_margin / 360000, 1) if secao.top_margin else 0
        margin_bottom = round(secao.bottom_margin / 360000, 1) if secao.bottom_margin else 0
        margin_left = round(secao.left_margin / 360000, 1) if secao.left_margin else 0
        margin_right = round(secao.right_margin / 360000, 1) if secao.right_margin else 0
        
        # Validação conforme imagem da caixa de diálogo "Configurar Página"
        if not (margin_top == 3.0 and margin_left == 3.0 and margin_bottom == 2.0 and margin_right == 2.0):
            margens_ok = False
            
        # Validação do tamanho do papel A4 (21.0 cm x 29.7 cm)
        page_width = round(secao.page_width / 360000, 1) if secao.page_width else 0
        page_height = round(secao.page_height / 360000, 1) if secao.page_height else 0
        if not (page_width == 21.0 and page_height == 29.7):
            papel_a4_ok = False

    # --- VARREDURA E PENTE FINO DE TEXTO ---
    fontes_reais = set()
    has_tables_or_images = False
    
    for p in doc.paragraphs:
        conteudo_linhas.append(p.text)
        
        # Análise estrita de formatação por Run de texto
        for r in p.runs:
            if r.font.name:
                fontes_reais.add(r.font.name.upper())
        
        # Validar propriedades do parágrafo (Alinhamento Justificado = 3)
        if p.alignment and p.alignment != docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY:
            alinhamento_ok = False
            
        # Validar espaçamento entre linhas (Norma pede 1.5)
        if p.paragraph_format.line_spacing and round(p.paragraph_format.line_spacing, 1) != 1.5:
            espacamento_linhas_ok = False

    # Leitura técnica das tabelas do documento
    if len(doc.tables) > 0:
        has_tables_or_images = True

    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                conteudo_linhas.append(celula.text)
                for p_celula in celula.paragraphs:
                    for r_celula in p_celula.runs:
                        if r_celula.font.name:
                            fontes_reais.add(r_celula.font.name.upper())
                            
    # Extração de cabeçalhos das seções
    for secao in doc.sections:
        if secao.header:
            for p in secao.header.paragraphs:
                conteudo_linhas.append(p.text)
                
    texto_completo = "\n".join(conteudo_linhas)
    
    if texto_completo:
        p_upper = texto_completo.upper()
        
        # Validações estruturais por string
        if "APÊNDICE" in p_upper or "ANEXO" in p_upper:
            has_apendices_ou_anexos = True
        if "HOSPITAL DA CIDADE" in p_upper or "SOCORRÃO" in p_upper:
            marca_dagua_corpo_ok = True
        if "VERSÃO" in p_upper and "PÁGINAS" in p_upper:
            cabecalho_completo = True
            
        if re.search(r"[A-Z0-9]{2,5}_[A-Z0-9\.\-]+", p_upper) or re.search(r"CÓDIGO:\s*[A-Z0-9\.\-_]+", p_upper):
            codigo_detectado_no_texto = True
            
        if any(p in p_upper for p in ["ELABORAÇÃO", "REVISÃO", "HISTÓRICO", "REGISTRO"]):
            registro_historico_ok = True
            
        # Validação obrigatória da Nota do Print (Campos adicionais do fim da folha de RH)
        if "DATA DE APROVAÇÃO:" in p_upper and "VERSÃO:" in p_upper:
            campos_nota_rh_ok = True
            
        # Cruzamento dinâmico das seções obrigatórias por Tipo
        itens_obrigatorios_do_tipo = REGRAS_NORMA_ZERO[tipo_detectado]["obrigatorios"]
        for item in itens_obrigatorios_do_tipo:
            itens_estrutura_dinamica.append(f"Estrutura: {item.title()}")
            if item in p_upper:
                valores_estrutura_dinamica.append(True)
            else:
                if item == "DESCRIÇÃO DAS ETAPAS" and autorizar_etapas_manual:
                    valores_estrutura_dinamica.append(True)
                else:
                    valores_estrutura_dinamica.append(False)
                    
    st.success(f"✔️ **Análise Concluída para o Tipo:** {REGRAS_NORMA_ZERO[tipo_detectado]['nome']}")

# Aplicação das autorizações manuais via Sidebar
if autorizar_codigo_manual:
    codigo_detectado_no_texto = True
if logo_prefeitura_manual:
    logo_prefeitura_manual = True
if marca_dagua_manual:
    marca_dagua_corpo_ok = True
if autorizar_fonte_manual:
    fonte_e_tamanho_ok = True

if fontes_reais:
    fontes_permitidas = {"CALIBRI", "ARIAL"}
    if not fontes_reais.issubset(fontes_permitidas):
        fonte_e_tamanho_ok = False

if not itens_estrutura_dinamica:
    for item in REGRAS_NORMA_ZERO[st.session_state.cached_tipo]["obrigatorios"]:
        itens_estrutura_dinamica.append(f"Estrutura: {item.title()}")
        valores_estrutura_dinamica.append(True if arquivo_word else False)

# Determinação do status técnico do item Impressos
status_impressos = "NÃO SE APLICA"
if arquivo_word:
    status_impressos = "SIM" if has_tables_or_images else "NÃO SE APLICA"

#--- 5. UNIFICAÇÃO DA CONFIGURAÇÃO DE ITENS ---
f_conf = [
    logo_prefeitura_manual, cabecalho_completo, cabecalho_completo, codigo_detectado_no_texto,
    cabecalho_completo, cabecalho_completo, True, registro_historico_ok, campos_nota_rh_ok, papel_a4_ok,
    margens_ok, fonte_e_tamanho_ok, espacamento_linhas_ok, alinhamento_ok, paragrafo_ok,
    True, paginacao_ok, marca_dagua_corpo_ok, cabecalho_completo, True, has_apendices_ou_anexos,
    status_impressos
]

todos_os_itens_finais = itens_ficha + itens_estrutura_dinamica
todas_as_conf_finais = f_conf + valores_estrutura_dinamica

