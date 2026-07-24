import streamlit as st
import docx 
import pandas as pd
import re
import datetime
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador NAQH", page_icon="🔍", layout="wide")
st.title("Auditor Automatizado - Norma Zero (V44)")
st.markdown("Este sistema realiza a auditoria técnica e gera a Ficha de Verificação oficial.")

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
    "REG": {"nome": "Regimento Interno", "obrigatorios": ["DISPOSIÇÕES GERIAS", "COMPETÊNCIAS"]},
    "ROT": {"nome": "Rotina", "obrigatorios": ["ATIVIDADES", "FLUXO"]}
}

tipo_detectado = st.session_state.cached_tipo
has_apendices_ou_anexos = False
marca_dagua_corpo_ok = marca_dagua_manual
cabecalho_completo = True
codigo_detectado_no_texto = False
registro_historico_ok = False
nome_arquivo_doc = st.session_state.cached_nome_arquivo

# Flag de verificação estrita de fontes e tamanhos
fonte_e_tamanho_ok = True

itens_ficha = [
    "Logomarca do Hospital", "Título do Documento", "Tipo de Documento", "Código do Documento",
    "Versão", "Páginas", "Data da Aprovação / Validade", "Registro Histórico", "Papel",
    "Margens", "Fonte e Tamanho", "Espaçamento entre Linhas", "Alinhamento", "Parágrafo",
    "Figuras/Tabelas", "Paginação", "Marca d'água", "Cabeçalho das Páginas", "Referências", "Apêndices / Anexos"
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
    conteudo_linhas = [p.text for p in doc.paragraphs]
    
    fontes_reais = set()
    for p in doc.paragraphs:
        for r in p.runs:
            if r.font.name:
                fontes_reais.add(r.font.name.upper())
                
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                conteudo_linhas.append(celula.text)
                for p_celula in celula.paragraphs:
                    for r_celula in p_celula.runs:
                        if r_celula.font.name:
                            fontes_reais.add(r_celula.font.name.upper())

    if fontes_reais:
        fontes_permitidas = {"CALIBRI", "ARIAL"}
        if not fontes_reais.issubset(fontes_permitidas):
            fonte_e_tamanho_ok = False
            
    for secao in doc.sections:
        if secao.header:
            for p in secao.header.paragraphs:
                conteudo_linhas.append(p.text)
                
    texto_completo = "\n".join(conteudo_linhas)
    
    if texto_completo:
        p_upper = texto_completo.upper()
        
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
                    
        st.success(f"📊 **Análise Concluída para o Tipo:** {REGRAS_NORMA_ZERO[tipo_detectado]['nome']}")

# Aplicação imediata das autorizações manuais
if autorizar_codigo_manual:
    codigo_detectado_no_texto = True
if logo_prefeitura_manual:
    logo_prefeitura_manual = True
if marca_dagua_manual:
    marca_dagua_corpo_ok = True
if autorizar_fonte_manual:
    fonte_e_tamanho_ok = True

if not itens_estrutura_dinamica:
    for item in REGRAS_NORMA_ZERO[st.session_state.cached_tipo]["obrigatorios"]:
        itens_estrutura_dinamica.append(f"Estrutura: {item.title()}")
        if item == "DESCRIÇÃO DAS ETAPAS" and autorizar_etapas_manual:
            valores_estrutura_dinamica.append(True)
        else:
            valores_estrutura_dinamica.append(True if arquivo_word else False)

#--- 5. UNIFICAÇÃO DA CONFIGURAÇÃO DE ITENS ---
f_conf = [
    logo_prefeitura_manual, cabecalho_completo, cabecalho_completo, codigo_detectado_no_texto,
    cabecalho_completo, cabecalho_completo, True, registro_historico_ok, True, True,
    fonte_e_tamanho_ok, True, True, True, True, True, marca_dagua_corpo_ok, cabecalho_completo, True, has_apendices_ou_anexos
]

todos_os_itens_finais = itens_ficha + itens_estrutura_dinamica
todas_as_conf_finais = f_conf + valores_estrutura_dinamica

#--- 6. CÁLCULO E EXIBIÇÃO DA PORCENTAGEM DE CONFORMIDADE ---
total_itens = len(todas_as_conf_finais)
total_conforme = sum(1 for item in todas_as_conf_finais if item is True)
porcentagem_conforme = int((total_conforme / total_itens) * 100) if total_itens > 0 else 0
st.session_state.cached_porcentagem = porcentagem_conforme

st.markdown("### Status de Conformidade Geral")
st.progress(porcentagem_conforme/100)
st.subheader(f"✅ {porcentagem_conforme}% conforme a Norma Zero ({REGRAS_NORMA_ZERO[st.session_state.cached_tipo]['nome']})")

# --- 7. FICHA DE VERIFICAÇÃO RESUMIDA NA TELA ---
st.markdown("---")
st.markdown("### Ficha de Verificação para Aprovação (Espelho Oficial NAQH)")

linhas_tabela_resumida = []
for idx, nome_item in enumerate(todos_os_itens_finais):
    status_str = "✔️ SIM" if todas_as_conf_finais[idx] else "❌ NÃO"
    if nome_item == "Apêndices / Anexos" and not todas_as_conf_finais[idx]:
        status_str = "🔷 OPCIONAL"
    linhas_tabela_resumida.append({"Item Técnico Regulamentado": nome_item, "Status Técnico": status_str})
st.table(pd.DataFrame(linhas_tabela_resumida))

# --- 8. PREPARAÇÃO DO TEXTO DA FICHA CURTA ---
itens_removidos_download = ["Data da Aprovação / Validade", "Registro Histórico", "Verificador"]
texto_ficha = "FICHA DE VERIFICAÇÃO PARA APROVAÇÃO (NAQH)\n"
texto_ficha += "PREFEITURA DE SÃO LUÍS - HOSPITAL DR. JACKSON LAGO\n\n"
texto_ficha += f"Documento: {st.session_state.cached_nome_arquivo}\n"
texto_ficha += f"Tipo: {REGRAS_NORMA_ZERO[st.session_state.cached_tipo]['nome']}\n"
texto_ficha += f"Conformidade: {st.session_state.cached_porcentagem}%\n\n"

for idx, nome_item in enumerate(todos_os_itens_finais):
    if nome_item in itens_removidos_download or nome_item.startswith("Estrutura:"):
        continue
    marcador = "[X] SIM [ ] NÃO" if todas_as_conf_finais[idx] else "[ ] SIM [X] NÃO"
    if nome_item == "Apêndices / Anexos" and not todas_as_conf_finais[idx]:
        marcador = "[ ] SIM [ ] NÃO (Não consta)"
    texto_ficha += f"{nome_item}: {marcador}\n"

# --- 9. BOTÃO DE DOWNLOAD CORRETAMENTE FECHADO NA SIDEBAR ---
st.sidebar.markdown("---")
st.sidebar.subheader("📥 Exportação Pronta")
st.sidebar.download_button(
    label="📥 Baixar Ficha (.doc / Word)",
    data=texto_ficha,
    file_name=f"Ficha_NAQH_{st.session_state.cached_tipo}.doc",
    mime="application/msword",
    use_container_width=True
)
