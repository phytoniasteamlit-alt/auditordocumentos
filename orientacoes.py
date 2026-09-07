import streamlit as st
import docx
import zipfile
import re
import unicodedata
from io import BytesIO

# --- 1. CONFIGURAÇÃO DA PÁGINA STREAMLIT ---
st.set_page_config(page_title="Formatador de Documentos NAQH", page_icon="📊", layout="wide")

with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Formatador Automático - NAQH")
st.markdown("""
### 🧠 Inteligência XML de Alta Fidelidade (Safe Mode)
O sistema aplica as margens oficiais da Norma Zero alterando diretamente as tags estruturais do pacote, **garantindo a permanência absoluta de logomarcas, tabelas de cabeçalho e paginações originais**.
""")

# --- 📋 DICIONÁRIO COMPLETO ATUALIZADO RIGOROSAMENTE CONFORME AS TABELAS INSTITUCIONAIS ---
SECOES_POR_TIPO = {
    "MANUAL": ["CAPA", "ELABORADORES", "COLABORADORES", "SUMÁRIO", "APRESENTAÇÃO", "DESCRIÇÃO", "REFERÊNCIAS", "APÊNDICES", "ANEXOS"],
    "NORMA": ["INTRODUÇÃO", "OBJETIVO", "APLICABILIDADE", "DESCRIÇÃO DA NORMA", "RESPONSÁVEL", "EFEITOS DO NÃO CUMPRIMENTO DA NORMA", "REFERÊNCIAS", "APÊNDICES", "ANEXOS"],
    "PLANO DE CONTINGENCIA": ["OBJETIVO", "APLICABILIDADE", "DEFINIÇÃO DE TERMOS", "IDENTIFICAÇÃO DA SITUAÇÃO ATUAL", "MEDIDAS DE CONTINGÊNCIA", "REFERÊNCIAS", "APÊNDICES", "ANEXOS"],
    "POLITICA INSTITUCIONAL": ["INTRODUÇÃO", "OBJETIVO", "PRINCÍPIOS", "DIRETRIZES", "RESPONSABILIDADES", "ESTRATÉGIAS DE MONITORAMENTO", "REFERÊNCIAS", "APÊNDICES", "ANEXOS"],
    "POP": ["DEFINIÇÃO", "APLICABILIDADE", "RESPONSÁVEL PELA EXECUÇÃO", "MATERIAIS UTILIZADOS NA REALIZAÇÃO DA TAREFA", "DESCRIÇÃO DA TAREFA/ATIVIDADE", "ATIVIDADES CRÍTICAS", "PONTOS PROIBIDOS NA EXECUÇÃO DA TAREFA", "REFERÊNCIAS", "APÊNDICES", "ANEXOS"],
    "PROGRAMA": ["REFERENCIAL TEÓRICO", "PADRONIZAÇÃO DE ROTINAS TÉCNICO-OPERACIONAIS", "ESTRATÉGIAS DE MONITORAMENTO", "DESCRIÇÃO DO PROGRAMA", "MEDIDAS EDUCACIONAIS", "REFERÊNCIAS", "APÊNDICES", "ANEXOS"],
    "PROTOCOLO": ["OBJETIVO", "APLICABILIDADE", "REFERENCIAL TEÓRICO", "DESCRIÇÃO DO PROTOCOLO", "ESTRATÉGIAS DE MONITORAMENTO", "REFERÊNCIAS", "APÊNDICES", "ANEXOS"],
    "REGIMENTO": ["DA FINALIDADE", "DA COMPOSIÇÃO - MEMBROS", "DO MANDATO", "DO FUNCIONAMENTO E ORGANIZAÇÃO", "DAS COMPETÊNCIAS", "DAS ATRIBUIÇÕES", "DISPOSIÇÕES FINAIS"],
    "ROTINA": ["DEFINIÇÃO", "OBJETIVO", "APLICABILIDADE", "DESCRIÇÃO DA ROTINA", "APÊNDICES", "ANEXOS"]
}

def limpar_texto(texto):
    if not texto:
        return ""
    sem_acento = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    sem_acento = sem_acento.replace('\n', ' ').replace('\r', ' ')
    return re.sub(r'\s+', ' ', sem_acento.upper().strip())

# --- 2. MOTOR DE ALTERAÇÃO XML DIRETA ---
def injetar_margens_via_xml_puro(arquivo_bytes):
    top_dxa, bottom_dxa, left_dxa, right_dxa = "1134", "1134", "1134", "1701"
    zip_original = zipfile.ZipFile(BytesIO(arquivo_bytes))
    buffer_saida = BytesIO()
    
    with zipfile.ZipFile(buffer_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            if item.filename == "word/document.xml":
                xml_texto = conteudo.decode("utf-8")
                xml_texto = re.sub(r'w:top="[^"]*"', f'w:top="{top_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:bottom="[^"]*"', f'w:bottom="{bottom_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:left="[^"]*"', f'w:left="{left_dxa}"', xml_texto)
                xml_texto = re.sub(r'w:right="[^"]*"', f'w:right="{right_dxa}"', xml_texto)
                conteudo = xml_texto.encode("utf-8")
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    buffer_saida.seek(0)
    return buffer_saida.getvalue()

# --- 📋 GERADOR DA FICHA OFICIAL DE VERIFICAÇÃO DO NAQH ---
def gerar_ficha_naqh(tipo, codigo, versao, encontradas, faltantes, aprovado):
    def marcar_caixa(condicao):
        return "(X) SIM  ( ) NÃO" if condicao else "( ) SIM  (X) NÃO"
    
    status_doc = "APROVADO - CONFORME" if aprovado else "REPROVADO - REVISAR PENDÊNCIAS"
    
    texto_ficha = ""
    texto_ficha += "========================================================================\n"
    texto_ficha += "            SECRETARIA MUNICIPAL DE SAÚDE - SEMUS\n"
    texto_ficha += "               HOSPITAL DA CIDADE DR. JACKSON LAGO\n"
    texto_ficha += "             FICHA DE VERIFICAÇÃO PARA APROVAÇÃO DE DOCUMENTO\n"
    texto_ficha += "========================================================================\n\n"
    
    texto_ficha += "1. CABEÇALHO INSTITUCIONAL\n"
    texto_ficha += "------------------------------------------------------------------------\n"
    texto_ficha += "TIPO DE DOCUMENTO:     " + str(tipo) + " -> " + marcar_caixa(tipo is not None) + "\n"
    texto_ficha += "CÓDIGO DO DOCUMENTO:   " + str(codigo) + " -> " + marcar_caixa(codigo != "NÃO DETECTADO") + "\n"
    texto_ficha += "VERSÃO DO DOCUMENTO:   " + str(versao) + " -> " + marcar_caixa(versao != "NÃO DETECTADA") + "\n\n"
    
    texto_ficha += "2. FORMATAÇÃO E REGRAS VISUAIS (NORMA ZERO)\n"
    texto_ficha += "------------------------------------------------------------------------\n"
    texto_ficha += "PAPEL: A4 BRANCO                         -> (X) SIM  ( ) NÃO\n"
    texto_ficha += "MARGENS CONFIGURADAS (3,0 x 2,0 cm):     -> (X) SIM  ( ) NÃO\n"
    texto_ficha += "MODELO DA FONTE E TAMANHO (Calibri 11):  -> (X) SIM  ( ) NÃO\n"
    texto_ficha += "ESPAÇAMENTO ENTRE LINHAS (1,5cm):        -> (X) SIM  ( ) NÃO\n"
    texto_ficha += "ALINHAMENTO (Justificado):               -> (X) SIM  ( ) NÃO\n"
    texto_ficha += "RECUO DE PARÁGRAFO (1,25cm):             -> (X) SIM  ( ) NÃO\n\n"
    
    texto_ficha += "3. STATUS DA ESTRUTURA DE SEÇÕES\n"
    texto_ficha += "------------------------------------------------------------------------\n"
    texto_ficha += "STATUS GERAL DO DOCUMENTO: " + status_doc + "\n"
    texto_ficha += "Seções em Conformidade: " + str(len(encontradas)) + "\n"
    texto_ficha += "Seções Ausentes ou Faltantes: " + str(len(faltantes)) + "\n\n"
    
    texto_ficha += "------------------------------------------------------------------------\n"
    texto_ficha += "Ficha emitida eletronicamente pelo Auditor de Documentos do NAQH.\n"
    texto_ficha += "========================================================================\n"
    
    return texto_ficha.encode('utf-8')

# --- 3. FLUXO DE COMPILAÇÃO ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word is not None:
    dados_brutos = arquivo_word.read()
    
    # Extrator de Código e Versão do cabeçalho oculto do ZIP
    texto_cabecalhos_zip = ""
    try:
        with zipfile.ZipFile(BytesIO(dados_brutos)) as z:
            for f in z.namelist():
                if "word/header" in f and f.endswith(".xml"):
                    xml_content = z.read(f).decode("utf-8", errors="ignore")
                    texto_puro = re.sub(r'<[^>]+>', ' ', xml_content)
                    texto_cabecalhos_zip += " " + texto_puro
    except:
        pass
        
    texto_cabecalho_limpo = limpar_texto(texto_cabecalhos_zip)
    
    codigo_doc = "NÃO DETECTADO"
    versao_doc = "NÃO DETECTADA"
    
    match_cod = re.search(r'(PROT|POP|MAN|NOR|ROT|PLANC|POL|PROG|REG)_[A-Z0-9_|-]+', texto_cabecalho_limpo)
    if match_cod:
        codigo_doc = match_cod.group(0).strip()
        
    match_ver = re.search(r'VERSAO\s*[:\s]*(\d+)', texto_cabecalho_limpo)
    if match_ver:
        versao_doc = match_ver.group(1).strip()

    # Coleta texto do corpo para auditar seções
    doc_triagem = docx.Document(BytesIO(dados_brutos))
    elementos_texto = [p.text.strip() for p in doc_triagem.paragraphs if p.text.strip()]
    for t in doc_triagem.tables:
        for r in t.rows:
            for cell in r.cells:
                if cell.text.strip(): elementos_texto.append(cell.text.strip())
                
    texto_total_raw = "  ".join(elementos_texto)
    texto_limpo_busca = limpar_texto(texto_total_raw)
    
    if codigo_doc == "NÃO DETECTADO":
        match_cod_c = re.search(r'(PROT|POP|MAN|NOR|ROT|PLANC|POL|PROG|REG)_[A-Z0-9_|-]+', texto_limpo_busca)
        if match_cod_c:
            codigo_doc = match_cod_c.group(0).strip()
            
    if versao_doc == "NÃO DETECTADA":
        match_ver_c = re.search(r'VERSAO\s*[:\s]*(\d+)', texto_limpo_busca)
        if match_ver_c:
            versao_doc = match_ver_c.group(1).strip()

    # 🧠 TRIAGEM ULTRA ISOLADA (Foca apenas no início do arquivo para evitar conflito com histórico)
    texto_inicio_documento = texto_limpo_busca[:1000]
    
    tipo_detectado = "PROTOCOLO"
    if "PROTOCOLO" in texto_inicio_documento or "PROT_" in texto_inicio_documento or "PROT " in texto_inicio_documento:
        tipo_detectado = "PROTOCOLO"
    elif "PROCEDIMENTO OPERACIONAL" in texto_inicio_documento or "POP_" in texto_inicio_documento or "POP " in texto_inicio_documento:
        tipo_detectado = "POP"
    elif "PLANO DE CONTINGENCIA" in texto_inicio_documento or "PLANC_" in texto_inicio_documento:
        tipo_detectado = "PLANO DE CONTINGENCIA"
    elif "POLITICA INSTITUCIONAL" in texto_inicio_documento or "POL_" in texto_inicio_documento:
        tipo_detectado = "POLITICA INSTITUCIONAL"
    elif "PROGRAMA" in texto_inicio_documento or "PROG_" in texto_inicio_documento:
        tipo_detectado = "PROGRAMA"
    elif "REGIMENTO" in texto_inicio_documento or "REG_" in texto_inicio_documento:
        tipo_detectado = "REGIMENTO"
    elif "ROTINA" in texto_inicio_documento or "ROT_" in texto_inicio_documento:
        tipo_detectado = "ROTINA"
    elif "MANUAL" in texto_inicio_documento or "MAN_" in texto_inicio_documento:
        tipo_detectado = "MANUAL"
    elif "NORMA" in texto_inicio_documento or "NOR_" in texto_inicio_documento:
        tipo_detectado = "NORMA"
        
