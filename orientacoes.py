import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import unicodedata
from io import BytesIO

st.set_page_config(page_title="AUDITORIA E FORMATAÇÃO AUTOMÁTICA", page_icon="🔍", layout="wide")

st.title("🔍 AUDITORIA + FORMATAÇÃO AUTOMÁTICA (NORMA ZERO)")
st.markdown("### ✅ Validação Inteligente, Formatação Estrita Calibri e Emissão de Ficha NAQH!")

# ============================================================
# 🧹 FUNÇÃO: REMOVER ACENTOS E NORMALIZAR TEXTO PARA BUSCA
# ============================================================
def limpar_texto(texto):
    if not texto:
        return ""
    sem_acento = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    sem_acento = sem_acento.replace('\n', ' ').replace('\r', ' ')
    return re.sub(r'\s+', ' ', sem_acento.upper().strip())

# ============================================================
# 📋 SEÇÕES OBRIGATÓRIAS (ATUALIZADAS CONFORME QUADRO 1 DA INSTITUIÇÃO)
# ============================================================
SECOES_POR_TIPO = {
    "PROT": ["1. OBJETIVO", "2. APLICABILIDADE", "3. REFERENCIAL TEÓRICO", "4. CLASSIFICAÇÃO DAS CIRURGIAS", "5. RESPONSABILIDADES", "6. MEDIDAS OBRIGATÓRIAS DE PREVENÇÃO", "7. ESTRATÉGIAS DE MONITORAMENTO", "8. REFERÊNCIAS"],
    "POP": ["1. DEFINIÇÃO", "2. APLICABILIDADE", "3. RESPONSÁVEL PELA EXECUÇÃO", "4. MATERIAIS UTILIZADOS NA REALIZAÇÃO DA TAREFA", "5. DESCRIÇÃO DOS PROCEDIMENTOS", "6. ATIVIDADES CRÍTICAS E PONTOS PROIBIDOS NA EXECUÇÃO DA TAREFA", "7. REFERÊNCIAS", "8. ANEXOS"],
    "POI": ["1. INTRODUÇÃO", "2. OBJETIVO", "3. FINALIDADE", "4. ABRANGÊNCIA", "5. RESPONSABILIDADES", "6. GESTÃO DE RISCO", "7. ANEXOS", "8. REFERÊNCIAS"],
    "NORMA": ["1. INTRODUÇÃO", "2. OBJETIVO", "3. APLICABILIDADE", "4. DESCRIÇÃO DA NORMA", "5. RESPONSÁVEIS", "6. EFEITOS DO NÃO CUMPRIMENTO DA NORMA", "7. REFERÊNCIAS"],
    "NOR": ["1. OBJETIVO", "2. APLICABILIDADE", "3. DESCRIÇÃO DA NORMA", "4. RESPONSÁVEL", "5. EFETIVO NO CUMPRIMENTO", "6. NORMA DE REFERÊNCIA", "7. ANEXOS"],
    "REG": ["1. FINALIDADE", "2. ÂMBITO", "3. COMPETÊNCIA E ORGANIZAÇÃO", "4. DISPOSIÇÕES GERAIS", "5. DISPOSIÇÕES FINAIS"],
    "PROG": ["1. REFERENCIAL TEÓRICO", "2. OBJETIVOS", "3. METAS E INDICADORES", "4. DEFINIÇÃO DE METAS", "5. ACOMPANHAMENTO E MONITORAMENTO", "6. AVALIÇÃO DE RESULTADOS", "7. REFERÊNCIAS", "8. ANEXOS"],
    "PLAN": ["1. OBJETIVO", "2. APLICABILIDADE", "3. DESCRIÇÃO DO CENÁRIO DE RISCO", "4. MEDIDAS DE CONTINGÊNCIA", "5. ESTRATÉGIAS DE RESPOSTA", "6. REFERÊNCIAS", "7. ANEXOS"],
    "ROT": ["1. OBJETIVO", "2. APLICABILIDADE", "3. DESCRIÇÃO DA ROTINA", "4. RESPONSÁVEL", "5. ETAPAS DE EXECUÇÃO", "6. REFERÊNCIAS", "7. ANEXOS"]
}

# ============================================================
# 🧠 ENGINE DE AUDITORIA E LEITURA DE CABEÇALHO
# ============================================================
def auditar_documento(doc):
    elementos_texto = []
    for p in doc.paragraphs:
        if p.text.strip(): elementos_texto.append(p.text)
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                if celula.text.strip(): elementos_texto.append(celula.text)
                
    texto_bruto = "  ".join(elementos_texto)
    texto = limpar_texto(texto_bruto)
    
    tipo_detectado = "PROT"
    if "NORMA" in texto:
        tipo_detectado = "NORMA"
    else:
        for tipo in SECOES_POR_TIPO.keys():
            if re.search(rf'\b{tipo}\b', texto):
                tipo_detectado = tipo
                break
            
    codigo_detectado = None
    match_codigo = re.search(r'CODIGO[\s:]+([A-Z0-9_|-]+)', texto)
    if match_codigo:
        codigo_detectado = match_codigo.group(1).strip()
        
    versao_detectada = None
    match_versao = re.search(r'VERSAO[\s:]*(\d+)', texto)
    if match_versao:
        versao_detectada = match_versao.group(1).strip()
        
    validade_detectada = None
    match_validade = re.search(r'VALIDADE[\s:]*([\d/X]+)', texto)
    if match_validade:
        validade_detectada = match_validade.group(1).strip()
        
    secoes_esperadas = SECOES_POR_TIPO[tipo_detectado]
    secoes_encontradas, secoes_faltantes = [], []
    
    for secao in secoes_esperadas:
        secao_sem_numero = re.sub(r'^\d+\s*', '', secao)
        secao_limpa = limpar_texto(secao_sem_numero)
        if re.search(rf'\b{re.escape(secao_limpa)}\b', texto):
            secoes_encontradas.append(secao)
        else:
            secoes_faltantes.append(secao)
            
    aprovado = (len(secoes_faltantes) == 0 and codigo_detectado is not None and versao_detectada is not None)
    
    return {
        "tipo": tipo_detectado, "codigo": codigo_detectado, "versao": versao_detectada,
        "validade": validade_detectada, "secoes_encontradas": secoes_encontradas,
        "secoes_faltantes": secoes_faltantes, "aprovado": aprovado
    }

# ============================================================
# 🎨 ENGINE DE FORMATAÇÃO (ESTRITO CALIBRI CONFORME NORMA ZERO)
# ============================================================
def formatar_pelas_normas(doc):
    # 1. Configuração de Margens Atualizadas (3x3x2x2)
    for section in doc.sections:
        section.top_margin = Cm(3.0)
        section.left_margin = Cm(3.0)
        section.bottom_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
        
    # 2. Formatação do Corpo de Texto Principal (Calibri 11)
    for p in doc.paragraphs:
        texto_paragrafo = p.text.strip()
        if not texto_paragrafo:
            continue
            
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        eh_titulo = any(limpar_texto(re.sub(r'^\d+\s*', '', s)) in limpar_texto(texto_paragrafo) for lista in SECOES_POR_TIPO.values() for s in lista)
        eh_lista = p.style.name.startswith('List') or texto_paragrafo.startswith(('-', '•', '*', 'a)', 'b)', 'c)', 'd)'))

        if eh_titulo:
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(12)
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                r.font.name = 'Calibri'
                r.font.size = Pt(11)
                r.bold = True
        elif eh_lista:
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                r.font.name = 'Calibri'
                r.font.size = Pt(11)
        else:
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.first_line_indent = Cm(1.25)  # Recuo padrão fixado em 1,25 cm
            for r in p.runs:
                r.font.name = 'Calibri'
                r.font.size = Pt(11)

    # 3. Formatação das Tabelas e Cabeçalhos (Calibri 10 - Espaçamento Simples)
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for p in celula.paragraphs:
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.first_line_indent = Cm(0)
                    for r in p.runs:
                        r.font.name = 'Calibri'
                        r.font.size = Pt(10)
                        
    output = BytesIO()
    doc.save(output)
    return output.getvalue()

# ============================================================
# 📋 GERADOR DA FICHA OFICIAL DE VERIFICAÇÃO DO NAQH (MÉTODO ULTRA SEGURO)
# ============================================================
def gerar_ficha_naqh(rel):
    def sim_nao(condicao):
        if condicao:
            return "(X) SIM  ( ) NÃO"
        return "( ) SIM  (X) NÃO"
    
    tipo_str = str(rel['tipo'])
    cod_str = str(rel['codigo'] or 'NÃO ENCONTRADO')
    ver_str = str(rel['versao'] or 'NÃO ENCONTRADA')
    val_str = str(rel['validade'] or 'NÃO PREENCHIDA')
    
    status_doc = "REPROVADO - REVISAR PENDÊNCIAS"
    if rel['aprovado']:
        status_doc = "APROVADO - CONFORME"
        
    cenc = str(len(rel['secoes_encontradas']))
    cfal = str(len(rel['secoes_faltantes']))
    
    linhas = [
        "========================================================================",
        "            SECRETARIA MUNICIPAL DE SAÚDE - SEMUS",
        "               HOSPITAL DA CIDADE DR. JACKSON LAGO",
        "             FICHA DE VERIFICAÇÃO PARA APROVAÇÃO DE DOCUMENTO",
        "========================================================================",
        "",
        "1. CABEÇALHO INSTITUCIONAL",
        "------------------------------------------------------------------------",
        "TIPO DE DOCUMENTO:     " + tipo_str + " -> " + sim_nao(rel['tipo'] is not None),
        "CÓDIGO DO DOCUMENTO:   " + cod_str + " -> " + sim_nao(rel['codigo'] is not None),
        "VERSÃO DO DOCUMENTO:   " + ver_str + " -> " + sim_nao(rel['versao'] is not None),
        "VALIDADE EXIBIDA:      " + val_str,
        "",
        "2. FORMATAÇÃO E REGRAS VISUAIS (NORMA ZERO)",
        "------------------------------------------------------------------------",
        "PAPEL: A4 BRANCO                         -> (X) SIM  ( ) NÃO",
        "MARGENS CONFIGURADAS (3x3x2x2):          -> (X) SIM  ( ) NÃO",
        "MODELO DA FONTE E TAMANHO (Calibri 11):  -> (X) SIM  ( ) NÃO",
        "ESPAÇAMENTO ENTRE LINHAS (1,5cm):        -> (X) SIM  ( ) NÃO",
        "ALINHAMENTO (Justificado):               -> (X) SIM  ( ) NÃO",
        "RECUO DE PARÁGRAFO (1,25cm):             -> (X) SIM  ( ) NÃO",
        "",
        "3. STATUS DA ESTRUTURA DE SEÇÕES",
