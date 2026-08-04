import streamlit as st
import docx
import pandas as pd
import re
from io import BytesIO

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador NAQH", page_icon="🔍", layout="wide")
st.title("Auditor Automatizado - Ficha Oficial NAQH")
st.markdown("Sistema completo focado no apontamento detalhado de desvios tipográficos conforme a Norma Zero.")

NOMES_TIPOS = {
    "NOR": "NORMA (NOR)",
    "POP": "PROCEDIMENTO OPERACIONAL PADRÃO (POP)",
    "PROT": "PROTOCOLO (PROT)",
    "MAN": "MANUAL (MAN)",
    "REG": "REGIMENTO INTERNO (REG)",
    "ROT": "ROTINA (ROT)",
    "POL": "POLÍTICA INSTITUCIONAL (POL)"
}

# --- 2. BARRA LATERAL (CONTROLES E CHECKBOXES) ---
st.sidebar.header("⚙️ Controles de Auditoria (NAQH)")
tem_impressos_inconformes = st.sidebar.checkbox("⚠️ Há imagens/fotos de impressos ilegíveis ou inconformes?")
logo_hospital_manual = st.sidebar.checkbox("Autorizar Manualmente: Logomarca do Hospital")
codigo_manual = st.sidebar.checkbox("Autorizar Manualmente: Código do Documento")
impresso_manual_conforme = st.sidebar.checkbox("Autorizar Manualmente: Itens Impresso (Conforme)")

# --- ESTRUTURA GLOBAL ATUALIZADA (DINÂMICA) ---
nome_arquivo_doc = "Documento Coletado"
tipo_detectado = "PROT"
status_impressos = "SIM"
comentario_impressos = "Conforme apresentado."
erros_formatacao = [] # Esta lista continuará sendo preenchida pelo seu motor de busca

# [O seu fluxo de carregamento e análise rigorosa (WORD .DOCX) continua exatamente igual aqui]
# ...
# (Suas funções e loops que fazem o append em erros_formatacao não mudam)
# ...

# --- CÁLCULO DINÂMICO DA CONFORMIDADE (NOVO) ---
# Remove duplicatas para não penalizar o mesmo erro mais de uma vez
lista_erros_painel = list(set(erros_formatacao))

# Cada erro desconta 5% da nota total. Se não houver erros, garante 100%.
desconto_por_erro = 5
porcentagem_conforme = max(0, 100 - (len(lista_erros_painel) * desconto_por_erro))

# --- CONFIGURAÇÃO DOS DICIONÁRIOS COM OS ESTADOS FINAIS ---
cabecalho_dados = {
    "LOGOMARCA DO HOSPITAL": True if logo_hospital_manual else False,
    "TÍTULO DO DOCUMENTO": True,
    "TIPO DE DOCUMENTO": True,
    "CÓDIGO DO DOCUMENTO": True if codigo_manual else False,
    "VERSÃO": True,
    "PÁGINAS": True
}

historico_dados = {
    "DATA DA APROVAÇÃO / VALIDADE": True,
    "REGISTRO HISTÓRICO DO DOCUMENTO": True
}

# Modifique as linhas finais de renderização para usar a nova porcentagem dinâmica:
st.progress(porcentagem_conforme / 100)
st.subheader(f"📊 {porcentagem_conforme}% Conformidade")
