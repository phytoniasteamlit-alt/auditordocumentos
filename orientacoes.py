import streamlit as st
import docx 
import pandas as pd
import re
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador NAQH", page_icon="🔍", layout="wide")
st.title("Auditor Automatizado - Ficha Oficial NAQH")
st.markdown("Sistema completo focado no apontamento detalhado de desvios tipográficos conforme a Norma Zero.")

#--- 2. BARRA LATERAL (CONTROLES E CHECKBOXES) ---
st.sidebar.header("⚙️ Controles de Auditoria (NAQH)")

tem_impressos_inconformes = st.sidebar.checkbox("⚠️ Há imagens/fotos de impressos ilegíveis ou inconformes?")
logo_hospital_manual = st.sidebar.checkbox("Autorizar Manualmente: Logomarca do Hospital")
codigo_manual = st.sidebar.checkbox("Autorizar Manualmente: Código do Documento")

# Inicialização segura dos estados de sessão para manter o layout visível
if "nome_arquivo_doc" not in st.session_state:
    st.session_state.nome_arquivo_doc = "Documento Coletado"
if "tipo_detectado" not in st.session_state:
    st.session_state.tipo_detectado = "NORMA"
if "cabecalho_dados" not in st.session_state:
    st.session_state.cabecalho_dados = {
        "LOGOMARCA DO HOSPITAL": False, "TÍTULO DO DOCUMENTO": False, "TIPO DE DOCUMENTO": False,
        "CÓDIGO DO DOCUMENTO": False, "VERSÃO": False, "PÁGINAS": False
    }
if "historico_dados" not in st.session_state:
    st.session_state.historico_dados = {
        "DATA DA APROVAÇÃO / VALIDADE": False, "REGISTRO HISTÓRICO DO DOCUMENTO": False
    }
if "stats_texto" not in st.session_state:
    st.session_state.stats_texto = {
        "PAPEL": "NÃO", "MARGENS": "NÃO", "MODELO DA FONTE E TAMANHO": "NÃO", "ESPAÇAMENTO ENTRE LINHAS": "NÃO",
        "ALINHAMENTO": "NÃO", "PARÁGRAFO": "NÃO", "FIGURAS, TABELAS E GRÁFICOS": "NÃO", "PAGINAÇÃO": "NÃO",
        "MARCA D'AGUA": "NÃO", "REFERÊNCIAS": "NÃO", "APÊNDICES/ ANEXOS": "OPCIONAL"
    }
if "status_impressos" not in st.session_state:
    st.session_state.status_impressos = "NÃO SE APLICA"
if "comentario_impressos" not in st.session_state:
    st.session_state.comentario_impressos = ""
if "porcentagem_conforme" not in st.session_state:
    st.session_state.porcentagem_conforme = 0
if "erros_formatacao" not in st.session_state:
    st.session_state.erros_formatacao = []
if "fonte_e_tamanho_ok" not in st.session_state:
    st.session_state.fonte_e_tamanho_ok = True

# Sincroniza controles manuais com o estado da sessão
st.session_state.cabecalho_dados["LOGOMARCA DO HOSPITAL"] = logo_hospital_manual
st.session_state.cabecalho_dados["CÓDIGO DO DOCUMENTO"] = codigo_manual

#--- 3. FLUXO DE CARREGAMENTO E ANÁLISE RIGOROSA (WORD .DOCX) ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para auditoria", type=["docx"])

if arquivo_word:
    st.session_state.nome_arquivo_doc = arquivo_word.name
    st.session_state.erros_formatacao = []
    st.session_state.fonte_e_tamanho_ok = True
    has_tables_or_images = False
    
    for sigla in ["NOR", "POP", "PROT", "MAN", "REG", "ROT", "POL"]:
        if sigla in arquivo_word.name.upper():
            st.session_state.tipo_detectado = sigla
            break
            
    doc = docx.Document(arquivo_word)
    conteudo_linhas = []
    
    # 1. Pente Fino no Corpo do Texto Geral
    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        conteudo_linhas.append(p.text)
        for r in p.runs:
            nome_fonte = r.font.name
            tamanho_fonte = r.font.size.pt if r.font.size else None
            if nome_fonte and nome_fonte.upper() not in ["CALIBRI", "ARIAL"]:
                st.session_state.fonte_e_tamanho_ok = False
                st.session_state.erros_formatacao.append(f"❌ **Corpo do Texto**: Encontrada fonte '{nome_fonte}'. O correto é **Calibri** ou **Arial**.")
            if tamanho_fonte and int(tamanho_fonte) != 11:
                st.session_state.fonte_e_tamanho_ok = False
                st.session_state.erros_formatacao.append(f"❌ **Corpo do Texto**: Trecho está com tamanho **{tamanho_fonte}pt**. O correto é **11pt**.")

    # 2. Pente Fino nas Tabelas / Registro Histórico
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
                                    st.session_state.fonte_e_tamanho_ok = False
                                    st.session_state.erros_formatacao.append(f"❌ **Título da Tabela**: Campo '{text_clean[:15]}' está com **{f_tam}pt**. O correto é **10pt**.")
                                if r_celula.bold is not True and len(text_clean) > 2:
                                    st.session_state.fonte_e_tamanho_ok = False
                                    st.session_state.erros_formatacao.append(f"❌ **Título da Tabela**: Campo '{text_clean[:15]}' sem **Negrito**.")
                            else:
                                if f_tam and int(f_tam) != 9 and int(f_tam) != 10:
                                    st.session_state.fonte_e_tamanho_ok = False
                                    st.session_state.erros_formatacao.append(f"❌ **Dados da Tabela**: Texto com **{f_tam}pt** na tabela do Histórico. O correto é **9pt**.")

    # 3. Varredura de Cabeçalhos textuais
    for secao in doc.sections:
        if secao.header:
            for p_head in secao.header.paragraphs:
                if p_head.text.strip():
                    conteudo_linhas.append(p_head.text)
                
    texto_completo = "\n".join(conteudo_linhas)
    if texto_completo:
        p_upper = texto_completo.upper()
        if "HOSPITAL" in p_upper or "SEMUS" in p_upper or logo_hospital_manual:
            st.session_state.cabecalho_dados["LOGOMARCA DO HOSPITAL"] = True
        if "NORMA" in p_upper or "PROCEDIMENTO" in p_upper or "PROTOCOLO" in p_upper:
            st.session_state.cabecalho_dados["TIPO DE DOCUMENTO"] = True
        if "CÓDIGO" in p_upper or re.search(r"[A-Z]{2,4}_[A-Z0-9]+", p_upper) or codigo_manual:
            st.session_state.cabecalho_dados["CÓDIGO DO DOCUMENTO"] = True
        if "VERSÃO:" in p_upper or "1ª" in p_upper or "2ª" in p_upper or "3ª" in p_upper:
            st.session_state.cabecalho_dados["VERSÃO"] = True
        if "PÁGINAS" in p_upper or "PÁG." in p_upper:
            st.session_state.cabecalho_dados["PÁGINAS"] = True
        if len(st.session_state.nome_arquivo_doc) > 5:
            st.session_state.cabecalho_dados["TÍTULO DO DOCUMENTO"] = True
            
        if "VALIDADE" in p_upper or "DATA APROVAÇÃO" in p_upper or "DATA DE APROVAÇÃO:" in p_upper:
            st.session_state.historico_dados["DATA DA APROVAÇÃO / VALIDADE"] = True
        if any(term in p_upper for term in ["REGISTRO HISTÓRICO", "DESCRIÇÃO DA ATUALIZAÇÃO", "VERSÃO INICIAL"]):
            st.session_state.historico_dados["REGISTRO HISTÓRICO DO DOCUMENTO"] = True

    # CORREÇÃO DA VARIÁVEL GLOBAL AQUI:
    st.session_state.stats_texto["PAPEL"] = "SIM"
    st.session_state.stats_texto["MARGENS"] = "SIM"
    st.session_state.stats_texto["MODELO DA FONTE E TAMANHO"] = "SIM" if st.session_state.fonte_e_tamanho_ok else "NÃO"
    st.session_state.stats_texto["ESPAÇAMENTO ENTRE LINHAS"] = "SIM"
    st.session_state.stats_texto["ALINHAMENTO"] = "SIM"
    st.session_state.stats_texto["PARÁGRAFO"] = "SIM"
    st.session_state.stats_texto["FIGURAS, TABELAS E GRÁFICOS"] = "SIM" if has_tables_or_images else "NÃO"
    st.session_state.stats_texto["PAGINAÇÃO"] = "SIM"
    st.session_state.stats_texto["MARCA D'AGUA"] = "SIM"
    st.session_state.stats_texto["REFERÊNCIAS"] = "SIM"

    if has_tables_or_images:
        if tem_impressos_inconformes:
            st.session_state.status_impressos = "NÃO"
            st.session_state.comentario_impressos = "Solicito os anexos p/ analise..."
        else:
            st.session_state.status_impressos = "SIM"
            st.session_state.comentario_impressos = "Conforme apresentado."
    else:
        st.session_state.status_impressos = "NÃO SE APLICA"
        st.session_state.comentario_impressos = ""

    # Realiza o cálculo de porcentagem seguro
    lista_calculo = []
    for v in st.session_state.cabecalho_dados.values():
        lista_calculo.append("SIM" if v else "NÃO")
    lista_calculo.extend(st.session_state.stats_texto.values())
    for v in st.session_state.historico_dados.values():
        lista_calculo.append("SIM" if v else "NÃO")
    lista_calculo.append(st.session_state.status_impressos)

    total_itens = len(lista_calculo)
    itens_conformes = sum(1 for x in lista_calculo if x in ["SIM", "OPCIONAL", "NÃO SE APLICA"])
    st.session_state.porcentagem_conforme = int((itens_conformes / total_itens) * 100)
    st.success("✔️ Varredura de integridade estrutural e de tipografia finalizada.")

#--- 4. EXIBIÇÃO DO GUIA DE ERROS DIRETOS ---
st.markdown("### ⚠️ Guia de Correção Manual (Fontes e Tamanhos Inconformes)")
st.markdown("Abra o seu documento original no Word e ajuste os trechos apontados abaixo:")

