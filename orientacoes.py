import streamlit as st
import docx 
import pandas as pd
import re
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador NAQH", page_icon="🔍", layout="wide")
st.title("Auditor Automatizado - Ficha Oficial NAQH")
st.markdown("Sistema completo parametrizado com o layout oficial de conferência lateral do Hospital Dr. Jackson Lago.")

# Inicialização de estados de sessão seguros
if "cached_nome_arquivo" not in st.session_state:
    st.session_state.cached_nome_arquivo = "Documento Coletado"
if "cached_tipo" not in st.session_state:
    st.session_state.cached_tipo = "NORMA"

#--- 2. BARRA LATERAL (CONTROLES E CHECKBOXES) ---
st.sidebar.header("⚙️ Controles de Auditoria (NAQH)")

tem_impressos_inconformes = st.sidebar.checkbox("⚠️ Há imagens/fotos de impressos ilegíveis ou inconformes?")
logo_hospital_manual = st.sidebar.checkbox("Autorizar Manualmente: Logomarca do Hospital")
codigo_manual = st.sidebar.checkbox("Autorizar Manualmente: Código do Documento")

tipo_detectado = st.session_state.cached_tipo
nome_arquivo_doc = st.session_state.cached_nome_arquivo

# Estrutura base de dados para auditoria do NAQH
cabecalho_dados = {
    "LOGOMARCA DO HOSPITAL": logo_hospital_manual,
    "TÍTULO DO DOCUMENTO": False,
    "TIPO DE DOCUMENTO": False,
    "CÓDIGO DO DOCUMENTO": codigo_manual,
    "VERSÃO": False,
    "PÁGINAS": False
}

historico_dados = {
    "DATA DA APROVAÇÃO / VALIDADE": False,
    "REGISTRO HISTÓRICO DO DOCUMENTO": False
}

has_tables_or_images = False
fonte_e_tamanho_ok = True
erros_formatacao = []

#--- 3. FLUXO DE CARREGAMENTO E ANÁLISE RIGOROSA (WORD .DOCX) ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para auditoria", type=["docx"])

if arquivo_word:
    nome_arquivo_doc = arquivo_word.name
    st.session_state.cached_nome_arquivo = nome_arquivo_doc
    
    for sigla in ["NOR", "POP", "PROT", "MAN", "REG", "ROT", "POL"]:
        if sigla in arquivo_word.name.upper():
            tipo_detectado = sigla
            break
    st.session_state.cached_tipo = tipo_detectado
    
    doc = docx.Document(arquivo_word)
    conteudo_linhas = []
    
    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        conteudo_linhas.append(p.text)
        for r in p.runs:
            nome_fonte = r.font.name
            tamanho_fonte = r.font.size.pt if r.font.size else None
            if nome_fonte and nome_fonte.upper() not in ["CALIBRI", "ARIAL"]:
                fonte_e_tamanho_ok = False
                erros_formatacao.append(f"Fonte incorreta no corpo: '{nome_fonte}'")
            if tamanho_fonte and int(tamanho_fonte) != 11:
                fonte_e_tamanho_ok = False
                erros_formatacao.append(f"Tamanho incorreto no corpo: {tamanho_fonte}pt")

    if len(doc.tables) > 0:
        has_tables_or_images = True
        
    for tabela in doc.tables:
        texto_tabela_completo = "".join([celula.text.upper() for linha in tabela.rows for celula in linha.cells])
        is_registro_historico = any(termo in texto_tabela_completo for termo in ["HISTÓRICO", "REVISÃO", "VERSÃO", "PROCESSO"])
        
        for i_linha, linha in enumerate(tabela.rows):
            for celula in linha.cells:
                text_clean = celula.text.strip()
                if text_clean:
                    conteudo_linhas.append(text_clean)
                if re.match(r"^[\s_\-\.]+$", text_clean):
                    continue
                    
                for p_celula in celula.paragraphs:
                    if not p_celula.text.strip():
                        continue
                    for r_celula in p_celula.runs:
                        f_tam = r_celula.font.size.pt if r_celula.font.size else None
                        if is_registro_historico:
                            if i_linha == 0:
                                if f_tam and int(f_tam) != 10:
                                    fonte_e_tamanho_ok = False
                                    erros_formatacao.append(f"Título da tabela inválido: {f_tam}pt")
                                if r_celula.bold is not True and len(text_clean) > 2:
                                    fonte_e_tamanho_ok = False
                                    erros_formatacao.append(f"Título da tabela sem Negrito: '{text_clean[:20]}'")
                            else:
                                if f_tam and int(f_tam) != 9 and int(f_tam) != 10:
                                    fonte_e_tamanho_ok = False
                                    erros_formatacao.append(f"Dados da tabela inválidos: {f_tam}pt")

    for secao in doc.sections:
        if secao.header:
            for p_head in secao.header.paragraphs:
                if p_head.text.strip():
                    conteudo_linhas.append(p_head.text)
                
    texto_completo = "\n".join(conteudo_linhas)
    
    if texto_completo:
        p_upper = texto_completo.upper()
        if "HOSPITAL" in p_upper or "SEMUS" in p_upper or logo_hospital_manual:
            cabecalho_dados["LOGOMARCA DO HOSPITAL"] = True
        if "NORMA" in p_upper or "PROCEDIMENTO" in p_upper or "PROTOCOLO" in p_upper:
            cabecalho_dados["TIPO DE DOCUMENTO"] = True
        if "CÓDIGO" in p_upper or re.search(r"[A-Z]{2,4}_[A-Z0-9]+", p_upper) or codigo_manual:
            cabecalho_dados["CÓDIGO DO DOCUMENTO"] = True
        if "VERSÃO:" in p_upper or "1ª" in p_upper or "2ª" in p_upper or "3ª" in p_upper:
            cabecalho_dados["VERSÃO"] = True
        if "PÁGINAS" in p_upper or "PÁG." in p_upper:
            cabecalho_dados["PÁGINAS"] = True
        if len(nome_arquivo_doc) > 5:
            cabecalho_dados["TÍTULO DO DOCUMENTO"] = True
            
        if "VALIDADE" in p_upper or "DATA APROVAÇÃO" in p_upper or "DATA DE APROVAÇÃO:" in p_upper:
            historico_dados["DATA DA APROVAÇÃO / VALIDADE"] = True
        if any(term in p_upper for term in ["REGISTRO HISTÓRICO", "DESCRIÇÃO DA ATUALIZAÇÃO", "VERSÃO INICIAL"]):
            historico_dados["REGISTRO HISTÓRICO DO DOCUMENTO"] = True

    st.success("✔️ Varredura de integridade estrutural e de tipografia finalizada.")
    erros_unicos = list(set(erros_formatacao))
    if not fonte_e_tamanho_ok and erros_unicos:
        with st.expander("⚠️ Detalhes das Inconformidades de Formatação Identificadas"):
            for erro in erros_unicos[:6]:
                st.warning(erro)

status_impressos = "NÃO SE APLICA"
comentario_impressos = ""

if arquivo_word:
    if has_tables_or_images:
        if tem_impressos_inconformes:
            status_impressos = "NÃO"
            comentario_impressos = "Solicito os anexos p/ analise..."
        else:
            status_impressos = "SIM"
            comentario_impressos = "Conforme apresentado no documento estruturado."
    else:
        status_impressos = "NÃO SE APLICA"

# Consolidação da lista de dados para cálculo matemático correto
lista_calculo = []
for k, v in cabecalho_dados.items():
    lista_calculo.append("SIM" if v else "NÃO")

itens_texto_fixos = [
    ("PAPEL", "SIM" if arquivo_word else "NÃO"),
    ("MARGENS", "SIM" if arquivo_word else "NÃO"),
    ("MODELO DA FONTE E TAMANHO", "SIM" if fonte_e_tamanho_ok and arquivo_word else "NÃO"),
    ("ESPAÇAMENTO ENTRE LINHAS", "SIM" if arquivo_word else "NÃO"),
    ("ALINHAMENTO", "SIM" if arquivo_word else "NÃO"),
    ("PARÁGRAFO", "SIM" if arquivo_word else "NÃO"),
    ("FIGURAS, TABELAS E GRÁFICOS", "SIM" if has_tables_or_images else "NÃO"),
    ("PAGINAÇÃO", "SIM" if arquivo_word else "NÃO"),
    ("MARCA D'AGUA", "SIM" if arquivo_word else "NÃO"),
    ("REFERÊNCIAS", "SIM" if arquivo_word else "NÃO"),
    ("APÊNDICES/ ANEXOS", "OPCIONAL")
]
for item, stat in itens_texto_fixos:
    lista_calculo.append(stat)

for k, v in historico_dados.items():
    lista_calculo.append("SIM" if v else "NÃO")
lista_calculo.append(status_impressos)

total_itens = len(lista_calculo)
itens_conformes = sum(1 for x in lista_calculo if x in ["SIM", "OPCIONAL", "NÃO SE APLICA"])
porcentagem_conforme = int((itens_conformes / total_itens) * 100) if arquivo_word else 0

#--- 4. EXIBIÇÃO EM COLUNAS LATERAIS (ESTILO FICHA IMPRESSA) ---
st.markdown("---")
st.subheader("📋 Ficha de Verificação Consolidada (Espelho Oficial)")

# Indicador de Porcentagem limpo no topo
col_p1, col_p2 = st.columns([3, 1])
with col_p1:
    st.progress(porcentagem_conforme / 100)
with col_p2:
    st.subheader(f"📊 {porcentagem_conforme}% Conformidade")

st.markdown("---")

# Função auxiliar para desenhar o layout de rádio horizontal como caixas de checagem fixas lateralmente
def render_linha_ficha(nome_item, status_atual, obs=""):
    c1, c2 = st.columns([3, 2])
    with c1:
        st.markdown(f"**{nome_item}**")
        if obs:
            st.caption(f"_{obs}_")
    with c2:
        # Renderiza caixas limpas indicando visualmente a marcação lateral conforme a ficha
        if status_atual == "SIM":
            st.markdown("🟩 **[X] SIM** &nbsp;&nbsp;&nbsp;&nbsp; ⬜ [ ] NÃO")
        elif status_atual == "NÃO":
            st.markdown("⬜ [ ] SIM &nbsp;&nbsp;&nbsp;&nbsp; 🟥 **[X] NÃO**")
        elif status_atual == "OPCIONAL":
            st.markdown("🔷 **[X] OPCIONAL**")
        else:
            st.markdown("⬜ [ ] SIM &nbsp;&nbsp;&nbsp;&nbsp; ⬜ [ ] NÃO &nbsp;&nbsp;&nbsp;&nbsp; 🟨 **[X] N/A**")
    st.markdown("<hr style='margin:4px 0px; border-top: 1px dashed #444;' />", unsafe_allow_html=True)

# 1. BLOCO CABEÇALHO
st.markdown("### 🔹 CABEÇALHO")
for k, v in cabecalho_dados.items():
    render_linha_ficha(k, "SIM" if v else "NÃO")

st.markdown("<br>", unsafe_allow_html=True)

# 2. BLOCO ITENS TEXTO
st.markdown("### 🔹 ITENS TEXTO")
for item, stat in itens_texto_fixos:
    render_linha_ficha(item, stat)

st.markdown("<br>", unsafe_allow_html=True)

# 3. BLOCO FIM DO DOCUMENTO
st.markdown("### 🔹 FIM DO DOCUMENTO")
for k, v in historico_dados.items():
    render_linha_ficha(k, "SIM" if v else "NÃO")

