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

# Inicialização e persistência segura de dados (Garante estabilidade visual imediata)
if "nome_arquivo_doc" not in st.session_state:
    st.session_state.nome_arquivo_doc = "Documento Coletado"
if "tipo_detectado" not in st.session_state:
    st.session_state.tipo_detectado = "NORMA"

if "cabecalho_dados" not in st.session_state:
    st.session_state.cabecalho_dados = {
        "LOGOMARCA DO HOSPITAL": True, "TÍTULO DO DOCUMENTO": True, "TIPO DE DOCUMENTO": True,
        "CÓDIGO DO DOCUMENTO": True, "VERSÃO": True, "PÁGINAS": True
    }
if "historico_dados" not in st.session_state:
    st.session_state.historico_dados = {
        "DATA DA APROVAÇÃO / VALIDADE": True, "REGISTRO HISTÓRICO DO DOCUMENTO": True
    }
if "stats_texto" not in st.session_state:
    st.session_state.stats_texto = {
        "PAPEL": "SIM", "MARGENS": "SIM", 
        "MODELO DA FONTE E TAMANHO (CORPO DO TEXTO)": "SIM", 
        "FONTE E TAMANHO (DENTRO DE TABELAS/HISTÓRICO)": "NÃO",
        "ESPAÇAMENTO ENTRE LINHAS": "SIM", "ALINHAMENTO": "SIM", "PARÁGRAFO": "SIM", 
        "FIGURAS, TABELAS E GRÁFICOS": "SIM", "PAGINAÇÃO": "SIM", "MARCA D'AGUA": "SIM", 
        "REFERÊNCIAS": "SIM", "APÊNDICES/ ANEXOS": "OPCIONAL"
    }
if "status_impressos" not in st.session_state:
    st.session_state.status_impressos = "NÃO SE APLICA"
if "comentario_impressos" not in st.session_state:
    st.session_state.comentario_impressos = ""
if "porcentagem_conforme" not in st.session_state:
    st.session_state.porcentagem_conforme = 90
if "erros_formatacao" not in st.session_state:
    st.session_state.erros_formatacao = []

# Sincronização dos controles da barra lateral
if logo_hospital_manual:
    st.session_state.cabecalho_dados["LOGOMARCA DO HOSPITAL"] = True
if codigo_manual:
    st.session_state.cabecalho_dados["CÓDIGO DO DOCUMENTO"] = True

# Dicionário de tradução dos tipos de documentos normativos
NOMES_TIPOS = {
    "NOR": "NORMA (NOR)",
    "POP": "PROCEDIMENTO OPERACIONAL PADRÃO (POP)",
    "PROT": "PROTOCOLO (PROT)",
    "MAN": "MANUAL (MAN)",
    "REG": "REGIMENTO INTERNO (REG)",
    "ROT": "ROTINA (ROT)",
    "POL": "POLÍTICA INSTITUCIONAL (POL)"
}

#--- 3. FLUXO DE CARREGAMENTO E ANÁLISE RIGOROSA (WORD .DOCX) ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para auditoria", type=["docx"])

if arquivo_word:
    st.session_state.nome_arquivo_doc = arquivo_word.name
    st.session_state.erros_formatacao = []
    has_tables_or_images = False
    corpo_texto_ok = True
    tabelas_texto_ok = True
    
    # Reinicia marcações estruturais para nova varredura dinamicamente
    st.session_state.cabecalho_dados = {k: False for k in st.session_state.cabecalho_dados.keys()}
    st.session_state.historico_dados = {k: False for k in st.session_state.historico_dados.keys()}
    
    if logo_hospital_manual: st.session_state.cabecalho_dados["LOGOMARCA DO HOSPITAL"] = True
    if codigo_manual: st.session_state.cabecalho_dados["CÓDIGO DO DOCUMENTO"] = True
    
    # Identificação do tipo de documento pelo nome do arquivo original
    st.session_state.tipo_detectado = "NOR"
    for sigla in NOMES_TIPOS.keys():
        if sigla in arquivo_word.name.upper():
            st.session_state.tipo_detectado = sigla
            break
            
    doc = docx.Document(arquivo_word)
    conteudo_linhas = []
    
    # 1. Varredura do Corpo do Texto
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

    # 2. Varredura das Tabelas
    if len(doc.tables) > 0:
        has_tables_or_images = True
        
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
                                    st.session_state.erros_formatacao.append(f"❌ **Cabeçalho**: O título '{text_clean[:15]}...' está com tamanho **{f_tam}pt**. O correto é **10pt (Negrito)**.")
                            else:
                                if f_tam and int(f_tam) != 9 and int(f_tam) != 10:
                                    tabelas_texto_ok = False
                                    st.session_state.erros_formatacao.append(f"❌ **Dados Internos**: O texto '{text_clean[:15]}...' está com tamanho **{f_tam}pt**. O correto é **9pt (Justificado)**.")

    # 3. Varredura de Cabeçalhos e Rodapés textuais das seções
    for secao in doc.sections:
        if secao.header:
            for p_head in secao.header.paragraphs:
                if p_head.text.strip():
                    conteudo_linhas.append(p_head.text)
        if secao.footer:
            for p_foot in secao.footer.paragraphs:
                if p_foot.text.strip():
                    conteudo_linhas.append(p_foot.text)
                
    texto_completo = "\n".join(conteudo_linhas)
    if texto_completo:
        p_upper = texto_completo.upper()
        if "HOSPITAL" in p_upper or "SEMUS" in p_upper or logo_hospital_manual:
            st.session_state.cabecalho_dados["LOGOMARCA DO HOSPITAL"] = True
        if any(term in p_upper for term in ["NORMA", "PROCEDIMENTO", "PROTOCOLO", "PROT", "MANUAL", "REGIMENTO"]):
            st.session_state.cabecalho_dados["TIPO DE DOCUMENTO"] = True
        if "CÓDIGO" in p_upper or re.search(r"[A-Z]{2,4}_[A-Z0-9]+", p_upper) or codigo_manual:
            st.session_state.cabecalho_dados["CÓDIGO DO DOCUMENTO"] = True
        if "VERSÃO:" in p_upper or "VERSÃO" in p_upper or "1ª" in p_upper or "2ª" in p_upper or "3ª" in p_upper:
            st.session_state.cabecalho_dados["VERSÃO"] = True
        if any(term in p_upper for term in ["PÁGINAS", "PÁG", "FL.", "FOLHA", "PAG"]):
            st.session_state.cabecalho_dados["PÁGINAS"] = True
        if len(st.session_state.nome_arquivo_doc) > 5:
            st.session_state.cabecalho_dados["TÍTULO DO DOCUMENTO"] = True
            
        if any(term in p_upper for term in ["VALIDADE", "DATA APROVAÇÃO", "DATA DE APROVAÇÃO", "APROVAÇÃO:"]):
            st.session_state.historico_dados["DATA DA APROVAÇÃO / VALIDADE"] = True
        if any(term in p_upper for term in ["REGISTRO HISTÓRICO", "DESCRIÇÃO DA ATUALIZAÇÃO", "VERSÃO INICIAL", "HISTÓRICO"]):
            st.session_state.historico_dados["REGISTRO HISTÓRICO DO DOCUMENTO"] = True

    st.session_state.stats_texto["MODELO DA FONTE E TAMANHO (CORPO DO TEXTO)"] = "SIM" if corpo_texto_ok else "NÃO"
    st.session_state.stats_texto["FONTE E TAMANHO (DENTRO DE TABELAS/HISTÓRICO)"] = "SIM" if tabelas_texto_ok else "NÃO"

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

#--- 4. INTERFACE GRÁFICA DO ESPELHO DA FICHA (RENDERIZAÇÃO DE SEÇÕES E TIPO DETECTADO) ---
st.markdown("---")
# EXIBIÇÃO DO TIPO: Mostra de forma destacada o tipo do documento normativo
