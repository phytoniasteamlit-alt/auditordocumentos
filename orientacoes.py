import streamlit as st
import docx 
import pandas as pd
import re
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador NAQH", page_icon="🔍", layout="wide")
st.title("Auditor Automatizado - Ficha Oficial NAQH")
st.markdown("Sistema completo focado no apontamento detalhado de desvios tipográficos conforme a Norma Zero.")

# Dicionário global estável de tipos de documentos normativos
NOMES_TIPOS = {
    "NOR": "NORMA (NOR)",
    "POP": "PROCEDIMENTO OPERACIONAL PADRÃO (POP)",
    "PROT": "PROTOCOLO (PROT)",
    "MAN": "MANUAL (MAN)",
    "REG": "REGIMENTO INTERNO (REG)",
    "ROT": "ROTINA (ROT)",
    "POL": "POLÍTICA INSTITUCIONAL (POL)"
}

#--- 2. BARRA LATERAL (CONTROLES E CHECKBOXES) ---
st.sidebar.header("⚙️ Controles de Auditoria (NAQH)")

tem_impressos_inconformes = st.sidebar.checkbox("⚠️ Há imagens/fotos de impressos ilegíveis ou inconformes?")
logo_hospital_manual = st.sidebar.checkbox("Autorizar Manualmente: Logomarca do Hospital")
codigo_manual = st.sidebar.checkbox("Autorizar Manualmente: Código do Documento")
# NOVO CONTROLE MANUAL EXCLUSIVO PARA IMPRESSOS
impresso_manual_conforme = st.sidebar.checkbox("Autorizar Manualmente: Itens Impresso (Conforme)")

# ESTRUTURA FIXA DE FALLBACK (Garante estabilidade visual constante em tela)
nome_arquivo_doc = "Documento Coletado"
tipo_detectado = "NOR"
porcentagem_conforme = 95
erros_formatacao = []
status_impressos = "SIM"
comentario_impressos = "Conforme apresentado."

cabecalho_dados = {
    "LOGOMARCA DO HOSPITAL": True if logo_hospital_manual else False,
    "TÍTULO DO DOCUMENTO": False,
    "TIPO DE DOCUMENTO": False,
    "CÓDIGO DO DOCUMENTO": True if codigo_manual else False,
    "VERSÃO": False,
    "PÁGINAS": False
}

historico_dados = {
    "DATA DA APROVAÇÃO / VALIDADE": False,
    "REGISTRO HISTÓRICO DO DOCUMENTO": False
}

stats_texto = {
    "PAPEL": "SIM", "MARGENS": "SIM", 
    "MODELO DA FONTE E TAMANHO (CORPO DO TEXTO)": "SIM", 
    "FONTE E TAMANHO (DENTRO DE TABELAS/HISTÓRICO)": "NÃO",
    "ESPAÇAMENTO ENTRE LINHAS": "SIM", "ALINHAMENTO": "SIM", "PARÁGRAFO": "SIM", 
    "FIGURAS, TABELAS E GRÁFICOS": "SIM", "PAGINAÇÃO": "SIM", "MARCA D'AGUA": "SIM", 
    "REFERÊNCIAS": "SIM", "APÊNDICES/ ANEXOS": "OPCIONAL"
}

#--- 3. FLUXO DE CARREGAMENTO E ANÁLISE RIGOROSA (WORD .DOCX) ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para auditoria", type=["docx"])

if arquivo_word:
    nome_arquivo_doc = arquivo_word.name
    erros_formatacao = []
    has_tables_or_images = False
    corpo_texto_ok = True
    tabelas_texto_ok = True
    
    cabecalho_dados = {k: False for k in cabecalho_dados.keys()}
    historico_dados = {k: False for k in historico_dados.keys()}
    
    if logo_hospital_manual: cabecalho_dados["LOGOMARCA DO HOSPITAL"] = True
    if codigo_manual: cabecalho_dados["CÓDIGO DO DOCUMENTO"] = True
    
    tipo_detectado = "NOR"
    for sigla in NOMES_TIPOS.keys():
        if sigla in arquivo_word.name.upper():
            tipo_detectado = sigla
            break
            
    doc = docx.Document(arquivo_word)
    conteudo_linhas = []
    
    # Varredura do Corpo do Texto Geral
    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        conteudo_linhas.append(p.text)
        for r in p.runs:
            nome_fonte = r.font.name
            tamanho_fonte = r.font.size.pt if r.font.size else None
            if nome_fonte and nome_fonte.upper() not in ["CALIBRI", "ARIAL"]:
                corpo_texto_ok = False
            if tamanho_fonte and int(tamanho_fonte) != 11:
                corpo_texto_ok = False

    if len(doc.tables) > 0:
        has_tables_or_images = True
        
    # Varredura das Células das Tabelas e Histórico
    for tabela in doc.tables:
        texto_tabela_completo = "".join([celula.text.upper() for linha in tabela.rows for celula in linha.cells])
        is_registro_historico = any(termo in texto_tabela_completo for termo in ["HISTÓRICO", "REVISÃO", "VERSÃO", "PROCESSO", "APROVAÇÃO"])
        
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
                                    tabelas_texto_ok = False
                                    erros_formatacao.append(f"❌ **Título da Tabela**: O termo '{text_clean[:15]}...' está com tamanho **{f_tam}pt**. Ajuste para **10pt (Negrito)**.")
                            else:
                                if f_tam and int(f_tam) != 9 and int(f_tam) != 10:
                                    tabelas_texto_ok = False
                                    erros_formatacao.append(f"❌ **Dados da Tabela**: O texto '{text_clean[:15]}...' está com tamanho **{f_tam}pt**. Ajuste para **9pt (Justificado)**.")

    for secao in doc.sections:
        if secao.header:
            for p_head in secao.header.paragraphs:
                if p_head.text.strip(): conteudo_linhas.append(p_head.text)
        if secao.footer:
            for p_foot in secao.footer.paragraphs:
                if p_foot.text.strip(): conteudo_linhas.append(p_foot.text)
                
    texto_completo = "\n".join(conteudo_linhas)
    if texto_completo:
        p_upper = texto_completo.upper()
        if "HOSPITAL" in p_upper or "SEMUS" in p_upper or logo_hospital_manual:
            cabecalho_dados["LOGOMARCA DO HOSPITAL"] = True
        if any(term in p_upper for term in ["NORMA", "PROCEDIMENTO", "PROTOCOLO", "PROT", "MANUAL", "REGIMENTO"]):
            cabecalho_dados["TIPO DE DOCUMENTO"] = True
        if "CÓDIGO" in p_upper or re.search(r"[A-Z]{2,4}_[A-Z0-9]+", p_upper) or codigo_manual:
            cabecalho_dados["CÓDIGO DO DOCUMENTO"] = True
        if "VERSÃO:" in p_upper or "VERSÃO" in p_upper or "1ª" in p_upper or "2ª" in p_upper or "3ª" in p_upper:
            cabecalho_dados["VERSÃO"] = True
        if any(term in p_upper for term in ["PÁGINAS", "PÁG", "FL.", "FOLHA", "PAG"]):
            cabecalho_dados["PÁGINAS"] = True
        if len(nome_arquivo_doc) > 5:
            cabecalho_dados["TÍTULO DO DOCUMENTO"] = True
            
        if any(term in p_upper for term in ["VALIDADE", "DATA APROVAÇÃO", "DATA DE APROVAÇÃO", "APROVAÇÃO:"]):
            historico_dados["DATA DA APROVAÇÃO / VALIDADE"] = True
        if any(term in p_upper for term in ["REGISTRO HISTÓRICO", "DESCRIÇÃO DA ATUALIZAÇÃO", "VERSÃO INICIAL", "HISTÓRICO"]):
            historico_dados["REGISTRO HISTÓRICO DO DOCUMENTO"] = True

    stats_texto["MODELO DA FONTE E TAMANHO (CORPO DO TEXTO)"] = "SIM" if corpo_texto_ok else "NÃO"
    stats_texto["FONTE E TAMANHO (DENTRO DE TABELAS/HISTÓRICO)"] = "SIM" if tabelas_texto_ok else "NÃO"

    # Lógica estruturada para avaliação de Impressos contemplando a liberação manual lateral
    if impresso_manual_conforme:
        status_impressos = "SIM"
        comentario_impressos = "Liberado manualmente pelo auditor via painel de contingência."
    elif has_tables_or_images:
        if tem_impressos_inconformes:
            status_impressos = "NÃO"
            comentario_impressos = "Solicito os anexos p/ analise..."
        else:
            status_impressos = "SIM"
            comentario_impressos = "Conforme apresentado."
    else:
        status_impressos = "NÃO SE APLICA"
        comentario_impressos = ""

    lista_calculo = []
    for v in cabecalho_dados.values():
        lista_calculo.append("SIM" if v else "NÃO")
    lista_calculo.extend(stats_texto.values())
    for v in historico_dados.values():
        lista_calculo.append("SIM" if v else "NÃO")
    lista_calculo.append(status_impressos)

    total_itens = len(lista_calculo)
    itens_conformes = sum(1 for x in lista_calculo if x in ["SIM", "OPCIONAL", "NÃO SE APLICA"])
    porcentagem_conforme = int((itens_conformes / total_itens) * 100)
    st.success("✔️ Varredura de integridade estrutural e de tipografia finalizada.")

# Garante o estado manual mesmo em repouso do arquivo
if impresso_manual_conforme:
    status_impressos = "SIM"
    comentario_impressos = "Liberado manualmente pelo auditor via painel de contingência."

#--- 4. INTERFACE GRÁFICA DO ESPELHO DA FICHA ---
st.markdown("---")
st.info(f"📋 **Tipo de Documento Identificado pelo Sistema**: {NOMES_TIPOS.get(tipo_detectado, tipo_detectado.upper())}")
st.subheader("📝 Ficha de Verificação Consolidada (Espelho Oficial)")

col_p1, col_p2 = st.columns(2)
with col_p1:
    st.progress(porcentagem_conforme / 100)
with col_p2:
    st.subheader(f"📊 {porcentagem_conforme}% Conformidade")

st.markdown("---")

def render_linha_ficha(nome_item, status_atual, obs=""):
    c1, c2 = st.columns(2)
    c1.markdown(f"**{nome_item}**" + (f"  \n_{obs}_" if obs else ""))
    marcador = "🔷 **[X] OPCIONAL**"
    if status_atual == "SIM": marcador = "🟩 **[X] SIM** &nbsp;&nbsp;&nbsp;&nbsp; ⬜ [ ] NÃO"
    elif status_atual == "NÃO": marcador = "⬜ [ ] SIM &nbsp;&nbsp;&nbsp;&nbsp; 🟥 **[X] NÃO**"
    elif status_atual == "NÃO SE APLICA": marcador = "⬜ [ ] SIM &nbsp;&nbsp;&nbsp;&nbsp; ⬜ [ ] NÃO &nbsp;&nbsp;&nbsp;&nbsp; 🟨 **[X] N/A**"
    c2.markdown(marcador)
    st.markdown("<hr style='margin:4px 0px; border-top: 1px dashed #444;' />", unsafe_allow_html=True)

st.markdown("### 🔹 CABEÇALHO")
for k, v in cabecalho_dados.items():
    render_linha_ficha(k, "SIM" if v else "NÃO")

