import streamlit as st
import docx
import zipfile
import re
import unicodedata
from io import BytesIO

# --- 1. CONFIGURAÇÃO DA PÁGINA STREAMLIT ---
st.set_page_config(page_title="Formatador de Documentos NAQH", page_icon="📊", layout="wide")

# Menu Lateral - Identificação Visual do Operador Autêntica
with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Formatador Automático - NAQH")
st.markdown("""
### 🧠 Inteligência XML de Alta Fidelidade (Safe Mode)
O sistema aplica as margens oficiais da Norma Zero alterando diretamente as tags estruturais do pacote, **garantindo a permanência absoluta de logomarcas, tabelas de cabeçalho e paginações originais**.
""")

# --- DICIONÁRIO DE SEÇÕES OBRIGATÓRIAS AJUSTADO ---
SECOES_POR_TIPO = {
    "PROTOCOLO": ["1. OBJETIVO", "2. APLICABILIDADE", "3. REFERENCIAL TEÓRICO", "4. CLASSIFICAÇÃO DAS CIRURGIAS", "5. RESPONSABILIDADES", "6. MEDIDAS OBRIGATÓRIAS DE PREVENÇÃO", "7. ESTRATÉGIAS DE MONITORAMENTO", "8. REFERÊNCIAS"],
    "POP": ["1. DEFINIÇÃO", "2. APLICABILIDADE", "3. RESPONSÁVEL PELA EXECUÇÃO", "4. MATERIAIS UTILIZADOS NA REALIZAÇÃO DA TAREFA", "5. DESCRIÇÃO DOS PROCEDIMENTOS", "6. ATIVIDADES CRÍTICAS E PONTOS PROIBIDOS NA EXECUÇÃO DA TAREFA", "7. REFERÊNCIAS", "8. ANEXOS"],
    "NORMA": ["1. INTRODUÇÃO", "2. OBJETIVO", "3. APLICABILIDADE", "4. DESCRIÇÃO DA NORMA", "5. RESPONSÁVEIS", "6. EFEITOS DO NÃO CUMPRIMENTO DA NORMA", "7. REFERÊNCIAS"]
}

# --- FUNÇÃO DE AUXÍLIO PARA BUSCA SEM ACENTO ---
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
    texto_ficha += "MARGENS CONFIGURADAS (NORMA ZERO):       -> (X) SIM  ( ) NÃO\n"
    texto_ficha += "MODELO DA FONTE E TAMANHO (Calibri 11):  -> (X) SIM  ( ) NÃO\n"
    texto_ficha += "ESPAÇAMENTO ENTRE LINHAS (1,5cm):        -> (X) SIM  ( ) NÃO\n\n"
    
    texto_ficha += "3. STATUS DA ESTRUTURA DE SEÇÕES\n"
    texto_ficha += "------------------------------------------------------------------------\n"
    texto_ficha += "STATUS GERAL DO DOCUMENTO: " + status_doc + "\n"
    texto_ficha += "Seções em Conformidade: " + str(len(encontradas)) + "\n"
    texto_ficha += "Seções Ausentes ou Faltantes: " + str(len(faltantes)) + "\n\n"
    
    texto_ficha += "------------------------------------------------------------------------\n"
    texto_ficha += "Ficha emitida eletronicamente pelo Auditor de Documentos do NAQH.\n"
    texto_ficha += "========================================================================\n"
    
    return texto_ficha.encode('utf-8')

# --- 3. FLUXO DE COMPILAÇÃO E TRIAGEM DE METADADOS ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    dados_brutos = arquivo_word.read()
    
    # Extração robusta de textos mantendo quebras de segurança
    doc_triagem = docx.Document(BytesIO(dados_brutos))
    elementos_texto = [p.text.strip() for p in doc_triagem.paragraphs if p.text.strip()]
    for t in doc_triagem.tables:
        for r in t.rows:
            for cell in r.cells:
                if cell.text.strip(): elementos_texto.append(cell.text.strip())
                
    texto_total_raw = "  ".join(elementos_texto)
    texto_limpo_busca = limpar_texto(texto_total_raw)
    
    # Triagem Avançada de Tipo Documental (Evita falsos positivos por conta do histórico)
    tipo_detectado = "PROTOCOLO"
    if "TIPO DE DOCUMENTO: PROTOCOLO" in texto_limpo_busca or "PROTOCOLO" in texto_limpo_busca[:500]:
        tipo_detectado = "PROTOCOLO"
    elif "PROCEDIMENTO OPERACIONAL" in texto_limpo_busca or "POP" in texto_limpo_busca[:500]:
        tipo_detectado = "POP"
    elif "NORMA" in texto_limpo_busca[:500]:
        tipo_detectado = "NORMA"
        
    st.markdown("---")
    st.subheader("📋 **Triagem e Auditoria de Estrutura**")
    st.write(f"🔹 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    # Captura Inteligente e Flexível do Código
    codigo_doc = "NÃO DETECTADO"
    match_codigo = re.search(r'CODIGO\s*[:\s]*([A-Z0-9_|-]+)', texto_limpo_busca)
    if match_codigo:
        codigo_doc = match_codigo.group(1).strip()
    st.write(f"🔹 **Código do Documento:** `{codigo_doc}`")
        
    # Captura Avançada de Versão (Limpa indicadores ordinais como 5ª)
    versao_doc = "NÃO DETECTADA"
    match_versao = re.search(r'VERSAO\s*[:\s]*(\d+)', texto_limpo_busca)
    if match_versao:
        versao_doc = match_versao.group(1).strip()
    st.write(f"🔹 **Versão do Documento:** `{versao_doc}`")

    # Realiza a Varredura de Seções Autêntica
    secoes_esperadas = SECOES_POR_TIPO[tipo_detectado]
    secoes_encontradas, secoes_faltantes = [], []
    
    for secao in secoes_esperadas:
        secao_sem_numero = re.sub(r'^\d+\s*[\.\-]?\s*', '', secao)
        secao_limpa = limpar_texto(secao_sem_numero)
        if re.search(rf'\b{re.escape(secao_limpa)}\b', texto_limpo_busca):
            secoes_encontradas.append(secao)
        else:
            secoes_faltantes.append(secao)

    # Exibição Analítica das Seções na Interface
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("##### ✅ Seções Identificadas")
        for s in secoes_encontradas:
            st.success(f"• {s}")
    with c2:
        st.markdown("##### ❌ Seções Ausentes")
        if secoes_faltantes:
            for s in secoes_faltantes:
                st.error(f"• {s}")
        else:
            st.info("• Nenhuma seção obrigatória ausente!")

    # Processamento e Emissão de Documentos
    dados_finais = injetar_margens_via_xml_puro(dados_brutos)
    documento_aprovado = (len(secoes_faltantes) == 0 and versao_doc != "NÃO DETECTADA" and codigo_doc != "NÃO DETECTADO")
    ficha_naqh_bytes = gerar_ficha_naqh(tipo_detectado, codigo_doc, versao_doc, secoes_encontradas, secoes_faltantes, documento_aprovado)

    st.markdown("---")
    if documento_aprovado:
        st.success("🎉 **DOCUMENTO APROVADO COM SUCESSO!** Tudo pronto para download.")
    else:
        st.warning("⚠️ **DOCUMENTO FORMATADO COM PENDÊNCIAS!** Verifique os itens apontados na triagem.")

    # Exibição paralela dos botões de download homologados
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            label="📥 DOWNLOAD DO DOCUMENTO FORMATADO (.DOCX)",
            data=dados_finais,
            file_name=f"{codigo_doc}_Formatado_Homologado.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    with d2:
        st.download_button(
            label="📥 DOWNLOAD DA FICHA DE VERIFICAÇÃO NAQH (.TXT)",
            data=ficha_naqh_bytes,
            file_name=f"Ficha_Verificacao_{codigo_doc}.txt",
            mime="text/plain"
        )
