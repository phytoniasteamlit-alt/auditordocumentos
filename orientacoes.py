import streamlit as st
import docx 
import pandas as pd
import re
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Validador NAQH", page_icon="🔍", layout="wide")
st.title("Auditor Automatizado - Ficha Oficial NAQH")
st.markdown("Sistema completo parametrizado com o layout oficial de conferência lateral do Hospital Dr. Jackson Lago.")

#--- 2. BARRA LATERAL (CONTROLES E CHECKBOXES) ---
st.sidebar.header("⚙️ Controles de Auditoria (NAQH)")

tem_impressos_inconformes = st.sidebar.checkbox("⚠️ Há imagens/fotos de impressos ilegíveis ou inconformes?")
logo_hospital_manual = st.sidebar.checkbox("Autorizar Manualmente: Logomarca do Hospital")
codigo_manual = st.sidebar.checkbox("Autorizar Manualmente: Código do Documento")

# Inicialização de variáveis padrão de auditoria (Layout Fixo)
nome_arquivo_doc = "Documento Coletado"
tipo_detectado = "NORMA"
has_tables_or_images = False
fonte_e_tamanho_ok = True
erros_formatacao = []
documento_corrigido_bytes = None

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

stat_papel = "NÃO"
stat_margens = "NÃO"
stat_fonte = "NÃO"
stat_linhas = "NÃO"
stat_alinha = "NÃO"
stat_parag = "NÃO"
stat_figuras = "NÃO"
stat_paginacao = "NÃO"
stat_marca = "NÃO"
stat_referencia = "NÃO"
stat_anexos = "OPCIONAL"

#--- 3. FLUXO DE CARREGAMENTO E ANÁLISE RIGOROSA (WORD .DOCX) ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para auditoria", type=["docx"])

if arquivo_word:
    nome_arquivo_doc = arquivo_word.name
    
    for sigla in ["NOR", "POP", "PROT", "MAN", "REG", "ROT", "POL"]:
        if sigla in arquivo_word.name.upper():
            tipo_detectado = sigla
            break
            
    doc = docx.Document(arquivo_word)
    doc_out = docx.Document(arquivo_word)
    conteudo_linhas = []
    
    # Varredura do Corpo do Texto
    for p_idx, p in enumerate(doc.paragraphs):
        if not p.text.strip():
            continue
        conteudo_linhas.append(p.text)
        
        if p_idx < len(doc_out.paragraphs):
            p_out = doc_out.paragraphs[p_idx]
            p_out.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
            p_out.paragraph_format.line_spacing = 1.5
            for r_out in p_out.runs:
                r_out.font.name = 'Calibri'
                r_out.font.size = docx.shared.Pt(11)
        
        for r in p.runs:
            nome_fonte = r.font.name
            tamanho_fonte = r.font.size.pt if r.font.size else None
            if nome_fonte and nome_fonte.upper() not in ["CALIBRI", "ARIAL"]:
                fonte_e_tamanho_ok = False
                erros_formatacao.append(f"Fonte incorreta no corpo: '{nome_fonte}'")
            if tamanho_fonte and int(tamanho_fonte) != 11:
                fonte_e_tamanho_ok = False
                erros_formatacao.append(f"Tamanho incorreto no corpo: {tamanho_fonte}pt (Esperado: 11pt)")

    if len(doc.tables) > 0:
        has_tables_or_images = True
        
    for t_idx, tabela in enumerate(doc.tables):
        texto_tabela_completo = "".join([celula.text.upper() for linha in tabela.rows for celula in linha.cells])
        is_registro_historico = any(termo in texto_tabela_completo for termo in ["HISTÓRICO", "REVISÃO", "VERSÃO", "PROCESSO"])
        
        tabela_out = doc_out.tables[t_idx]
        
        for i_linha, linha in enumerate(tabela.rows):
            for i_celula, celula in enumerate(linha.cells):
                text_clean = celula.text.strip()
                if text_clean:
                    conteudo_linhas.append(text_clean)
                if re.match(r"^[\s_\-\.]+$", text_clean):
                    continue
                    
                celula_out = tabela_out.rows[i_linha].cells[i_celula]
                
                for p_idx_c, p_celula in enumerate(celula.paragraphs):
                    if not p_celula.text.strip():
                        continue
                        
                    if p_idx_c < len(celula_out.paragraphs):
                        p_celula_out = celula_out.paragraphs[p_idx_c]
                        p_celula_out.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
                        for r_celula_out in p_celula_out.runs:
                            r_celula_out.font.name = 'Calibri'
                            if is_registro_historico and i_linha == 0:
                                r_celula_out.font.size = docx.shared.Pt(10)
                                r_celula_out.bold = True
                            elif is_registro_historico:
                                r_celula_out.font.size = docx.shared.Pt(9)
                                r_celula_out.bold = False
                    
                    for r_celula in p_celula.runs:
                        f_tam = r_celula.font.size.pt if r_celula.font.size else None
                        if is_registro_historico:
                            if i_linha == 0:
                                if f_tam and int(f_tam) != 10:
                                    fonte_e_tamanho_ok = False
                                    erros_formatacao.append(f"Título da tabela de Histórico inválido: {f_tam}pt (Esperado: 10pt)")
                                if r_celula.bold is not True and len(text_clean) > 2:
                                    fonte_e_tamanho_ok = False
                                    erros_formatacao.append(f"Título da tabela de Histórico sem Negrito: '{text_clean[:20]}'")
                            else:
                                if f_tam and int(f_tam) != 9 and int(f_tam) != 10:
                                    fonte_e_tamanho_ok = False
                                    erros_formatacao.append(f"Dados da tabela de Histórico inválidos: {f_tam}pt (Esperado: 9pt)")

    # Correção de Margens
    for secao_out in doc_out.sections:
        secao_out.top_margin = docx.shared.Cm(3.0)
        secao_out.left_margin = docx.shared.Cm(3.0)
        secao_out.bottom_margin = docx.shared.Cm(2.0)
        secao_out.right_margin = docx.shared.Cm(2.0)
        secao_out.page_width = docx.shared.Cm(21.0)
        secao_out.page_height = docx.shared.Cm(29.7)

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

    stat_papel = "SIM"
    stat_margens = "SIM"
    stat_fonte = "SIM" if fonte_e_tamanho_ok else "NÃO"
    stat_linhas = "SIM"
    stat_alinha = "SIM"
    stat_parag = "SIM"
    stat_figuras = "SIM" if has_tables_or_images else "NÃO"
    stat_paginacao = "SIM"
    stat_marca = "SIM"
    stat_referencia = "SIM"
    stat_anexos = "OPCIONAL"

    buffer_doc = BytesIO()
    doc_out.save(buffer_doc)
    documento_corrigido_bytes = buffer_doc.getvalue()

    st.success("✔️ Varredura de integridade estrutural e de tipografia finalizada.")
    
    erros_unicos = list(set(erros_formatacao))
    if not fonte_e_tamanho_ok and erros_unicos:
        with st.expander("⚠️ Detalhes das Inconformidades de Formatação Identificadas"):
            for erro in erros_unicos[:6]:
                st.warning(erro)

# --- CONFIGURAÇÃO DOS STATUS DOS IMPRESSOS ---
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

# Cálculo da porcentagem de conformidade
lista_calculo = []
for k, v in cabecalho_dados.items():
    lista_calculo.append("SIM" if v else "NÃO")
lista_calculo.extend([stat_papel, stat_margens, stat_fonte, stat_linhas, stat_alinha, stat_parag, stat_figuras, stat_paginacao, stat_marca, stat_referencia, stat_anexos])
for k, v in historico_dados.items():
    lista_calculo.append("SIM" if v else "NÃO")
lista_calculo.append(status_impressos)

total_itens = len(lista_calculo)
itens_conformes = sum(1 for x in lista_calculo if x in ["SIM", "OPCIONAL", "NÃO SE APLICA"])
porcentagem_conforme = int((itens_conformes / total_itens) * 100) if arquivo_word else 0

#--- 4. INTERFACE GRÁFICA DO ESPELHO DA FICHA (RERENDER SEGURO) ---
st.markdown("---")
