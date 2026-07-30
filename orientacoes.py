import streamlit as st
import docx 
import pandas as pd
import re
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador NAQH", page_icon="🔍", layout="wide")
st.title("Auditor Automatizado - Ficha Oficial NAQH")
st.markdown("Sistema parametrizado com pente fino de fontes/tamanhos e regras setoriais estritas do Hospital Dr. Jackson Lago.")

# Inicialização de estados de sessão seguros
if "cached_nome_arquivo" not in st.session_state:
    st.session_state.cached_nome_arquivo = "Documento Coletado"
if "cached_tipo" not in st.session_state:
    st.session_state.cached_tipo = "NORMA"

#--- 2. BARRA LATERAL (CONTROLES ESPECÍFICOS DO SETOR) ---
st.sidebar.header("⚙️ Controles de Auditoria (NAQH)")

# Parâmetros para o Bloco de Impressos (4º Print)
tem_impressos_inconformes = st.sidebar.checkbox("⚠️ Há imagens/fotos de impressos ilegíveis ou inconformes?")

# Liberações manuais de contingência para o Cabeçalho
logo_hospital_manual = st.sidebar.checkbox("Autorizar Manualmente: Logomarca do Hospital")
codigo_manual = st.sidebar.checkbox("Autorizar Manualmente: Código do Documento")

tipo_detectado = st.session_state.cached_tipo
nome_arquivo_doc = st.session_state.cached_nome_arquivo

# Status iniciais dos itens auditados pelo NAQH
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
    
    # Identificação automática do Tipo de Documento pelo Nome do Arquivo
    for sigla in ["NOR", "POP", "PROT", "MAN", "REG", "ROT", "POL"]:
        if sigla in arquivo_word.name.upper():
            tipo_detectado = sigla
            break
    st.session_state.cached_tipo = tipo_detectado
    
    doc = docx.Document(arquivo_word)
    conteudo_linhas = []
    
    # --- PENTE FINO DE FONTES E TAMANHOS (NORMA ZERO DETALHADA) ---
    
    # 1. Varredura do Corpo do Texto Geral (Deve ser Calibri 11 ou Arial 11)
    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        conteudo_linhas.append(p.text)
        
        for r in p.runs:
            nome_fonte = r.font.name
            tamanho_fonte = r.font.size.pt if r.font.size else None
            
            if nome_fonte:
                if nome_fonte.upper() not in ["CALIBRI", "ARIAL"]:
                    fonte_e_tamanho_ok = False
                    erros_formatacao.append(f"Fonte incorreta no corpo do texto: '{nome_fonte}' (Esperado: Calibri ou Arial)")
            if tamanho_fonte:
                if int(tamanho_fonte) != 11:
                    fonte_e_tamanho_ok = False
                    erros_formatacao.append(f"Tamanho incorreto no corpo: {tamanho_fonte}pt (Esperado: 11pt)")

    # 2. Varredura das Tabelas / Registro Histórico (Títulos: 10pt Negrito | Dados: 9pt Justificado)
    if len(doc.tables) > 0:
        has_tables_or_images = True
        
    for tabela in doc.tables:
        for i_linha, linha in enumerate(tabela.rows):
            for celula in linha.cells:
                text_clean = celula.text.strip()
                if text_clean:
                    conteudo_linhas.append(text_clean)
                
                for p_celula in celula.paragraphs:
                    if not p_celula.text.strip():
                        continue
                        
                    # Verifica Alinhamento das Células de Dados das tabelas (Deve ser Justificado = 3)
                    if i_linha > 0 and p_celula.alignment and p_celula.alignment != docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY:
                        fonte_e_tamanho_ok = False
                        erros_formatacao.append(f"Alinhamento incorreto na célula de dados da tabela (Esperado: Justificado)")

                    for r_celula in p_celula.runs:
                        f_nome = r_celula.font.name
                        f_tam = r_celula.font.size.pt if r_celula.font.size else None
                        
                        # Linha de Títulos das Tabelas (Exemplo: Versão | Data | Descrição de atualização)
                        if i_linha == 0:
                            if f_tam and int(f_tam) != 10:
                                fonte_e_tamanho_ok = False
                                erros_formatacao.append(f"Título da tabela com tamanho inválido: {f_tam}pt (Esperado: 10pt)")
                            if r_celula.bold is not True and len(p_celula.text.strip()) > 2:
                                fonte_e_tamanho_ok = False
                                erros_formatacao.append(f"Título da tabela deveria estar em Negrito: '{p_celula.text[:20]}...'")
                        
                        # Linhas de preenchimento/dados internos da tabela
                        else:
                            if f_tam and int(f_tam) != 9:
                                fonte_e_tamanho_ok = False
                                erros_formatacao.append(f"Dados internos da tabela com tamanho inválido: {f_tam}pt (Esperado: 9pt)")

    # Leitura técnica das seções de cabeçalhos das páginas
    for secao in doc.sections:
        if secao.header:
            for p_head in secao.header.paragraphs:
                if p_head.text.strip():
                    conteudo_linhas.append(p_head.text)
                
    texto_completo = "\n".join(conteudo_linhas)
    
    if texto_completo:
        p_upper = texto_completo.upper()
        
        # --- AUDITORIA DO BLOCO 1: CABEÇALHO DA PÁGINA (3º Print) ---
        if "HOSPITAL" in p_upper or "SEMUS" in p_upper or logo_hospital_manual:
            cabecalho_dados["Logomarca do Hospital"] = True
        if "NORMA" in p_upper or "PROCEDIMENTO" in p_upper or "PROTOCOLO" in p_upper:
            cabecalho_dados["Tipo de Documento"] = True
        if "CÓDIGO" in p_upper or re.search(r"[A-Z]{2,4}_[A-Z0-9]+", p_upper) or codigo_manual:
            cabecalho_dados["Código do Documento"] = True
        if "VERSÃO:" in p_upper or "1ª" in p_upper or "2ª" in p_upper or "3ª" in p_upper or "4ª" in p_upper:
            cabecalho_dados["Versão"] = True
        if "PÁGINAS" in p_upper or "PÁG." in p_upper:
            cabecalho_dados["Páginas"] = True
        if len(nome_arquivo_doc) > 5:
            cabecalho_dados["Título do Documento"] = True
            
        # --- AUDITORIA DO BLOCO 4: REGISTRO HISTÓRICO (1º Print) ---
        if "VALIDADE" in p_upper or "DATA APROVAÇÃO" in p_upper or "DATA DE APROVAÇÃO:" in p_upper:
            historico_dados["Data da Aprovação / Validade"] = True
        if any(term in p_upper for term in ["REGISTRO HISTÓRICO", "DESCRIÇÃO DA ATUALIZAÇÃO", "VERSÃO INICIAL"]):
            historico_dados["Registro Histórico do Documento"] = True

    # Renderização das inconformidades encontradas na tela principal
    st.success("✔️ Varredura de integridade estrutural e de tipografia finalizada.")
    if not fonte_e_tamanho_ok and erros_formatacao:
        with st.expander("⚠️ Detalhes das Inconformidades de Formatação Identificadas"):
            # Exibe no máximo as 8 primeiras falhas para manter a legibilidade
            for erro in list(set(erros_formatacao))[:8]:
                st.warning(erro)

# Avaliação do status lógico para a Ficha de Impressos (4º Print)
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
        comentario_impressos = "Não se aplica (Não constam anexos impressos no documento)."

#--- 4. RELATÓRIO DO ESPELHO DA FICHA NA TELA DO SISTEMA ---
st.markdown("---")
st.subheader("📋 Ficha de Verificação Técnica - Visão do Auditor (NAQH)")
st.caption("Nota de Conformidade Escopo: Os blocos 2, 5 e 6 pertencem a outras divisões institucionais e foram suprimidos automaticamente.")

linhas_exibicao = []
for k, v in cabecalho_dados.items():
    linhas_exibicao.append({"Categoria": "1. Cabeçalho (Print 3)", "Item Técnico Regulamentado": k, "Status de Conformidade": "SIM" if v else "NÃO", "Observação Interna": ""})

for k, v in historico_dados.items():
    linhas_exibicao.append({"Categoria": "Fim do Documento (Print 1)", "Item Técnico Regulamentado": k, "Status de Conformidade": "SIM" if v else "NÃO", "Observação Interna": ""})

linhas_exibicao.append({
    "Categoria": "2. Impressos (Print 4)", 
    "Item Técnico Regulamentado": "Itens Impresso (Estruturas Gráficas/Tabelas)", 
    "Status de Conformidade": status_impressos, 
    "Observação Interna": comentario_impressos
})

st.table(pd.DataFrame(linhas_exibicao))

#--- 5. CONSTRUÇÃO E FORMATAÇÃO DA STRING DE TEXTO PARA EXPORTAÇÃO COMPATÍVEL WORD ---
texto_documento_word = (
    "SÃO LUÍS | SEMUS\n"
    "PREFEITURA DE SÃO LUÍS\n"
    "SECRETARIA MUNICIPAL DE SAÚDE\n"
    "HOSPITAL DA CIDADE DR. JACKSON LAGO\n"
    "FICHA DE VERIFICAÇÃO PARA APROVAÇÃO DO DOCUMENTO\n"
    "============================================================\n\n"
    f"DOCUMENTO EM ANÁLISE: {nome_arquivo_doc}\n"
    f"SETOR RESPONSÁVEL: NAQH (Núcleo de Avaliação e Qualidade Hospitalar)\n"
    "------------------------------------------------------------\n\n"
    "1. CABEÇALHO (DIRETRIZ DO PRINT 3)\n"
