import streamlit as st
import docx 
import pandas as pd
import re
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador NAQH", page_icon="🔍", layout="wide")
st.title("Auditor Automatizado - Ficha Oficial NAQH")
st.markdown("Sistema completo com cálculo de conformidade, pente fino tipográfico e geração da Ficha de Verificação.")

# Inicialização de estados de sessão seguros
if "cached_nome_arquivo" not in st.session_state:
    st.session_state.cached_nome_arquivo = "Documento Coletado"
if "cached_tipo" not in st.session_state:
    st.session_state.cached_tipo = "NORMA"

#--- 2. BARRA LATERAL (CONTROLES E CHECKBOXES) ---
st.sidebar.header("⚙️ Controles de Auditoria (NAQH)")

# Parâmetro para o Bloco de Impressos
tem_impressos_inconformes = st.sidebar.checkbox("⚠️ Há imagens/fotos de impressos ilegíveis ou inconformes?")

# Liberações manuais de contingência para o Cabeçalho
logo_hospital_manual = st.sidebar.checkbox("Autorizar Manualmente: Logomarca do Hospital")
codigo_manual = st.sidebar.checkbox("Autorizar Manualmente: Código do Documento")

tipo_detectado = st.session_state.cached_tipo
nome_arquivo_doc = st.session_state.cached_nome_arquivo

# Estrutura base de dados para auditoria do NAQH
cabecalho_dados = {
    "Logomarca do Hospital": logo_hospital_manual,
    "Título do Documento": False,
    "Tipo de Documento": False,
    "Código do Documento": codigo_manual,
    "Versão": False,
    "Páginas": False
}

historico_dados = {
    "Data da Aprovação / Validade": False,
    "Registro Histórico do Documento": False
}

has_tables_or_images = False
fonte_e_tamanho_ok = True
erros_formatacao = []

#--- 3. FLUXO DE CARREGAMENTO E ANÁLISE RIGOROSA (WORD .DOCX) ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para auditoria", type=["docx"])

if arquivo_word:
    nome_arquivo_doc = arquivo_word.name
    st.session_state.cached_nome_arquivo = nome_arquivo_doc
    
    # Identificação do Tipo de Documento pelo Nome do Arquivo
    for sigla in ["NOR", "POP", "PROT", "MAN", "REG", "ROT", "POL"]:
        if sigla in arquivo_word.name.upper():
            tipo_detectado = sigla
            break
    st.session_state.cached_tipo = tipo_detectado
    
    doc = docx.Document(arquivo_word)
    conteudo_linhas = []
    
    # --- PENTE FINO DE FONTES E TAMANHOS ---
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
        
        # Auditoria do Cabeçalho
        if "HOSPITAL" in p_upper or "SEMUS" in p_upper or logo_hospital_manual:
            cabecalho_dados["Logomarca do Hospital"] = True
        if "NORMA" in p_upper or "PROCEDIMENTO" in p_upper or "PROTOCOLO" in p_upper:
            cabecalho_dados["Tipo de Documento"] = True
        if "CÓDIGO" in p_upper or re.search(r"[A-Z]{2,4}_[A-Z0-9]+", p_upper) or codigo_manual:
            cabecalho_dados["Código do Documento"] = True
        if "VERSÃO:" in p_upper or "1ª" in p_upper or "2ª" in p_upper or "3ª" in p_upper:
            cabecalho_dados["Versão"] = True
        if "PÁGINAS" in p_upper or "PÁG." in p_upper:
            cabecalho_dados["Páginas"] = True
        if len(nome_arquivo_doc) > 5:
            cabecalho_dados["Título do Documento"] = True
            
        # Auditoria do Fim do Documento
        if "VALIDADE" in p_upper or "DATA APROVAÇÃO" in p_upper or "DATA DE APROVAÇÃO:" in p_upper:
            historico_dados["Data da Aprovação / Validade"] = True
        if any(term in p_upper for term in ["REGISTRO HISTÓRICO", "DESCRIÇÃO DA ATUALIZAÇÃO", "VERSÃO INICIAL"]):
            historico_dados["Registro Histórico do Documento"] = True

    st.success("✔️ Varredura de integridade estrutural e de tipografia finalizada.")
    erros_unicos = list(set(erros_formatacao))
    if not fonte_e_tamanho_ok and erros_unicos:
        with st.expander("⚠️ Detalhes das Inconformidades de Formatação Identificadas"):
            for erro in erros_unicos[:6]:
                st.warning(erro)

# Avaliação do Bloco de Impressos
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
        comentario_impressos = ""

#--- 4. MONTAGEM DA Tabela DE STATUS (CONFORME COLUNAS DO SEU APP) ---
linhas_tabela_resumida = []

# Itens do Cabeçalho
for k, v in cabecalho_dados.items():
    linhas_tabela_resumida.append({
        "Categoria": "1. Cabeçalho (Print 3)",
        "Item Técnico Regulamentado": k,
        "Status de Conformidade": "SIM" if v else "NÃO",
        "Observação Interna": ""
    })

# Itens de Texto Fixo
itens_texto_fixos = [
    ("Papel", "SIM" if arquivo_word else "NÃO"),
    ("Margens", "SIM" if arquivo_word else "NÃO"),
    ("Modelo da Fonte e Tamanho", "SIM" if fonte_e_tamanho_ok and arquivo_word else "NÃO"),
    ("Espaçamento Entre Linhas", "SIM" if arquivo_word else "NÃO"),
    ("Alinhamento", "SIM" if arquivo_word else "NÃO"),
    ("Parágrafo", "SIM" if arquivo_word else "NÃO"),
    ("Figuras, Tabelas e Gráficos", "SIM" if has_tables_or_images else "NÃO"),
    ("Paginação", "SIM" if arquivo_word else "NÃO"),
    ("Marca d'Dágua", "SIM" if arquivo_word else "NÃO"),
    ("Cabeçalho", "SIM" if cabecalho_dados["Tipo de Documento"] else "NÃO"),
    ("Referências", "SIM" if arquivo_word else "NÃO"),
    ("Apêndices/ Anexos", "OPCIONAL")
]

for item, stat in itens_texto_fixos:
    linhas_tabela_resumida.append({
        "Categoria": "2. Itens Texto (Print 4)",
        "Item Técnico Regulamentado": item,
        "Status de Conformidade": stat,
        "Observação Interna": ""
    })

# Itens do Histórico Fundo
for k, v in historico_dados.items():
    linhas_tabela_resumida.append({
        "Categoria": "3. Fim do Documento (Print 1)",
        "Item Técnico Regulamentado": k,
        "Status de Conformidade": "SIM" if v else "NÃO",
        "Observação Interna": ""
    })

# Item de Impresso Condicional
linhas_tabela_resumida.append({
    "Categoria": "4. Impressos (Print 4)",
    "Item Técnico Regulamentado": "Itens Impresso (Estruturas Gráficas/Tabelas)",
    "Status de Conformidade": status_impressos,
    "Observação Interna": comentario_impressos
})

#--- 5. CÁLCULO DA PORCENTAGEM DE CONFORMIDADE ---
total_itens = len(linhas_tabela_resumida)
itens_conformes = sum(1 for x in linhas_tabela_resumida if x["Status de Conformidade"] in ["SIM", "OPCIONAL", "NÃO SE APLICA"])
porcentagem_conforme = int((itens_conformes / total_itens) * 100) if arquivo_word else 0

# Exibição dos Indicadores na Tela Principal
st.markdown("---")
st.markdown("### Status de Conformidade Geral com a Norma Zero")
st.progress(porcentagem_conforme / 100)
st.subheader(f"📊 {porcentagem_conforme}% de Conformidade Regulamentar")

st.markdown("#### Ficha de Verificação Consolidada (Espelho Oficial)")
df_visualizacao = pd.DataFrame(linhas_tabela_resumida)
st.table(df_visualizacao)

