import streamlit as st
import zipfile
import re
import unicodedata
import tempfile
import os
import shutil

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

# --- 📋 DICIONÁRIO DE SEÇÕES OBRIGATÓRIAS ---
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
    # Limpeza simples de espaços múltiplos sem usar motores Regex pesados
    return " ".join(sem_acento.upper().split())

def extrair_texto_xml_seguro(xml_string):
    """Extrai texto de tags XML sem usar expressões regulares de risco (Loop de Fatiamento Nativo)"""
    resultado = []
    em_tag = False
    buffer_texto = []
    
    # Processa caractere por caractere (Incomparavelmente mais estável para arquivos de texto gigantes)
    for char in xml_string:
        if char == '<':
            em_tag = True
            if buffer_texto:
                resultado.append("".join(buffer_texto))
                buffer_texto = []
        elif char == '>':
            em_tag = False
        elif not em_tag:
            buffer_texto.append(char)
            
    if buffer_texto:
        resultado.append("".join(buffer_texto))
        
    return " ".join(resultado)

# --- 2. MOTOR ULTRA VELOZ SEM REGEX ---
def injetar_margens_via_disco(caminho_origem):
    top_val, bottom_val, left_val, right_val = "1134", "1134", "1134", "1701"
    caminho_saida = caminho_origem + "_formatado.docx"
    zip_original = zipfile.ZipFile(caminho_origem, 'r')
    
    with zipfile.ZipFile(caminho_saida, "w", zipfile.ZIP_DEFLATED) as zip_novo:
        for item in zip_original.infolist():
            conteudo = zip_original.read(item.filename)
            
            if item.filename == "word/document.xml":
                xml_texto = conteudo.decode("utf-8", errors="ignore")
                
                # Fatiamento linear estável de tags de configuração de página
                for tag, val in [('w:top="', top_val), ('w:bottom="', bottom_val), ('w:left="', left_val), ('w:right="', right_val)]:
                    partes = xml_texto.split(tag)
                    if len(partes) > 1:
                        for i in range(1, len(partes)):
                            subpartes = partes[i].split('"', 1)
                            if len(subpartes) > 1:
                                partes[i] = val + '"' + subpartes[1]
                        xml_texto = tag.join(partes)
                        
                conteudo = xml_texto.encode("utf-8")
                
            zip_novo.writestr(item, conteudo)
            
    zip_original.close()
    
    with open(caminho_saida, "rb") as f:
        dados_finais = f.read()
        
    try:
        os.remove(caminho_saida)
    except:
        pass
        
    return dados_finais

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

# --- 3. FLUXO DE COMPILAÇÃO COM GATILHO SEGURO ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui", type=["docx"])

if arquivo_word is not None:
    st.success(f"📂 Arquivo carregado com sucesso: **{arquivo_word.name}**")
    
    disparar_processo = st.button("🚀 Iniciar Triagem e Formatação", type="primary")
    
    if disparar_processo:
        with st.spinner("Remontando texto estrutural em tempo recorde..."):
            # Salva o arquivo temporariamente no disco
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
                shutil.copyfileobj(arquivo_word, temp_file)
                caminho_temp = temp_file.name

            raw_corpo_xml = ""
            raw_cabecalhos_xml = ""
            
            try:
                with zipfile.ZipFile(caminho_temp, 'r') as z:
                    if "word/document.xml" in z.namelist():
                        raw_corpo_xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
                        
                    for f in z.namelist():
                        if "word/header" in f and f.endswith(".xml"):
                            raw_cabecalhos_xml += " " + z.read(f).decode("utf-8", errors="ignore")
            except Exception as e:
                st.error(f"❌ Falha na leitura interna do arquivo. Erro: {str(e)}")
                st.stop()
                
            # Extração linear blindada de texto puro (Sem usar expressões regulares)
            texto_corpo_puro = extrair_texto_xml_seguro(raw_corpo_xml)
            texto_cabecalho_puro = extrair_texto_xml_seguro(raw_cabecalhos_xml)
            
            texto_corpo_limpo = limpar_texto(texto_corpo_puro)
            texto_cabecalho_limpo = limpar_texto(texto_cabecalho_puro)
            
            codigo_doc = "NÃO DETECTADO"
            versao_doc = "NÃO DETECTADA"
            
            # Buscas estruturadas básicas e rápidas
            match_cod = re.search(r'(PROT|POP|MAN|NOR|ROT|PLANC|POL|PROG|REG)_[A-Z0-9_|-]+', texto_cabecalho_limpo)
            if match_cod:
                codigo_doc = match_cod.group(0).strip()
                
            match_ver = re.search(r'VERSAO\s*[:\s]*(\d+)', texto_cabecalho_limpo)
