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

# --- ESTRUTURA GLOBAL PERMANENTE ---
nome_arquivo_doc = "Documento Coletado"
tipo_detectado = "PROT"
status_impressos = "SIM"
comentario_impressos = "Conforme apresentado."
erros_formatacao = []

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

stats_texto = {
    "PAPEL": "SIM", 
    "MARGENS": "SIM",
    "MODELO DA FONTE E TAMANHO (CORPO DO TEXTO)": "SIM",
    "FONTE E TAMANHO (DENTRO DE TABELAS/HISTÓRICO)": "SIM",
    "ESPAÇAMENTO ENTRE LINHAS": "SIM", 
    "ALINHAMENTO": "SIM", 
    "PARÁGRAFO": "SIM",
    "FIGURAS, TABELAS E GRÁFICOS": "SIM", 
    "PAGINAÇÃO": "SIM", 
    "MARCA D'AGUA": "SIM",
    "REFERÊNCIAS": "SIM", 
    "APÊNDICES/ ANEXOS": "OPCIONAL"
}

# --- 3. FLUXO DE CARREGAMENTO (WORD .DOCX) ---
# Mantido no escopo principal para evitar que a tela quebre ou o botão suma
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para auditoria", type=["docx"])

porcentagem_conforme = 100

if arquivo_word:
    nome_arquivo_doc = arquivo_word.name
    doc = docx.Document(arquivo_word)
    
    # VALIDAÇÃO DE MARGENS (Padrão ABNT Trabalhos Acadêmicos)
    for section in doc.sections:
        sup = round(section.top_margin.cm, 1) if section.top_margin else 0
        inf = round(section.bottom_margin.cm, 1) if section.bottom_margin else 0
        esq = round(section.left_margin.cm, 1) if section.left_margin else 0
        dir_m = round(section.right_margin.cm, 1) if section.right_margin else 0
        
        # Regra ABNT atualizada: Superior 3.0, Esquerda 3.0, Inferior 2.0, Direita 2.0
        if sup != 3.0 or esq != 3.0 or inf != 2.0 or dir_m != 2.0:
            stats_texto["MARGENS"] = "NÃO"
            erros_formatacao.append(
                f"❌ **Margens Inconformes**: Detectado Sup: {sup}cm, Inf: {inf}cm, Esq: {esq}cm, Dir: {dir_m}cm. "
                f"Ajuste para o padrão ABNT (Superior e Esquerda: 3cm / Inferior e Direita: 2cm)."
            )

    # [O restante dos seus loops originais de análise de texto/tabelas entra aqui]
    # ...
    
    # --- CÁLCULO DINÂMICO APÓS A LEITURA DO ARQUIVO ---
    lista_erros_painel = list(set(erros_formatacao))
    desconto_por_erro = 5
    porcentagem_conforme = max(0, 100 - (len(lista_erros_painel) * desconto_por_erro))

# --- 4. INTERFACE GRÁFICA ---
st.markdown("---")
st.info(f"📄 **Tipo de Documento Identificado pelo Sistema**: {NOMES_TIPOS.get(tipo_detectado, tipo_detectado.upper())}")
st.subheader("📋 Ficha de Verificação Consolidada (Espelho Oficial)")

col_p1, col_p2 = st.columns(2)
with col_p1:
    st.progress(porcentagem_conforme / 100)
with col_p2:
    st.subheader(f"📊 {porcentagem_conforme}% Conformidade")

def render_linha_ficha(nome_item, status_atual):
    c1, c2 = st.columns(2)
    c1.markdown(f"**{nome_item}**")
    if status_atual == "SIM":
        marcador = "🟩 **[X] SIM**"
    elif status_atual == "NÃO":
        marcador = "🟥 **[X] NÃO**"
    else:
        marcador = "🔷 **[X] OPCIONAL**"
    c2.markdown(marcador)

st.markdown("### 🔹 CABEÇALHO")
for k, v in cabecalho_dados.items():
    render_linha_ficha(k, "SIM" if v else "NÃO")

st.markdown("### 🔹 ITENS TEXTO")
for item, stat in stats_texto.items():
    render_linha_ficha(item, stat)

if arquivo_word and erros_formatacao:
    st.markdown("### ⚠️ Guia de Correção Manual")
    for erro in list(set(erros_formatacao)):
        st.error(erro)
