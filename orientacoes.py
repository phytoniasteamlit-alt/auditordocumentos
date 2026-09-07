import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import unicodedata
from io import BytesIO

st.set_page_config(page_title="AUDITORIA E FORMATAÇÃO AUTOMÁTICA", page_icon="🔍", layout="wide")

st.title("🔍 AUDITORIA + FORMATAÇÃO AUTOMÁTICA (NORMA ZERO)")
st.markdown("### ✅ Detecta cabeçalhos, valida seções, formata pelas regras ABNT e libera o Download!")

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
# 📋 SEÇÕES OBRIGATÓRIAS (NORMA ZERO)
# ============================================================
SECOES_POR_TIPO = {
    "PROT": ["1. OBJETIVO", "2. APLICABILIDADE", "3. REFERENCIAL TEÓRICO", "4. CLASSIFICAÇÃO DAS CIRURGIAS", "5. RESPONSABILIDADES", "6. MEDIDAS OBRIGATORIAS DE PREVENÇÃO", "7. ESTRATÉGIAS DE MONITORAMENTO", "8. REFERÊNCIAS"],
    "POP": ["1. DEFINIÇÃO", "2. APLICABILIDADE", "3. RESPONSÁVEL", "4. DESCRIÇÃO DA EXECUÇÃO", "5. MATERIAIS UTILIZADOS", "6. TARIFA", "7. REFERÊNCIAS", "8. ANEXOS"],
    "POI": ["1. INTRODUÇÃO", "2. OBJETIVO", "3. FINALIDADE", "4. ABRANGÊNCIA", "5. RESPONSABILIDADES", "6. GESTÃO DE RISCO", "7. ANEXOS", "8. REFERÊNCIAS"],
    "NOR": ["1. OBJETIVO", "2. APLICABILIDADE", "3. DESCRIÇÃO DA NORMA", "4. RESPONSÁVEL", "5. EFETIVO NO CUMPRIMENTO", "6. NORMA DE REFERÊNCIA", "7. ANEXOS"],
    "REG": ["1. FINALIDADE", "2. ÂMBITO", "3. COMPETÊNCIA E ORGANIZAÇÃO", "4. DISPOSIÇÕES GERAIS", "5. DISPOSIÇÕES FINAIS"],
    "PROG": ["1. REFERENCIAL TEÓRICO", "2. OBJETIVOS", "3. METAS E INDICADORES", "4. DEFINIÇÃO DE METAS", "5. ACOMPANHAMENTO E MONITORAMENTO", "6. AVALIÇÃO DE RESULTADOS", "7. REFERÊNCIAS", "8. ANEXOS"],
    "PLAN": ["1. OBJETIVO", "2. APLICABILIDADE", "3. DESCRIÇÃO DO CENÁRIO DE RISCO", "4. MEDIDAS DE CONTINGÊNCIA", "5. ESTRATÉGIAS DE RESPOSTA", "6. REFERÊNCIAS", "7. ANEXOS"],
    "ROT": ["1. OBJETIVO", "2. APLICABILIDADE", "3. DESCRIÇÃO DA ROTINA", "4. RESPONSÁVEL", "5. ETAPAS DE EXECUCAO", "6. REFERÊNCIAS", "7. ANEXOS"]
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
        secao_sem_numero = re.sub(r'^\d+\.\s*', '', secao)
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
# 🎨 ENGINE DE FORMATAÇÃO E REGRAS VISUAIS (NORMA ZERO)
# ============================================================
def formatar_pelas_normas(doc):
    # 1. Configuração Estrita de Margens (ABNT / Norma Zero)
    for section in doc.sections:
        section.top_margin = Cm(3.0)
        section.left_margin = Cm(3.0)
        section.bottom_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
        
    # 2. Formatação do Corpo de Texto Principal
    for p in doc.paragraphs:
        texto_paragrafo = p.text.strip()
        if not texto_paragrafo:
            continue
            
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Verifica se o parágrafo atual é um Título/Subtítulo Principal
        eh_titulo = any(limpar_texto(re.sub(r'^\d+\.\s*', '', s)) in limpar_texto(texto_paragrafo) for lista in SECOES_POR_TIPO.values() for s in lista)
        
        # Verifica se o texto opera como item listado (bullets ou alfabéticos)
        eh_lista = p.style.name.startswith('List') or texto_paragrafo.startswith(('-', '•', '*', 'a)', 'b)', 'c)', 'd)', 'e)', 'f)', 'g)', 'h)', 'i)', 'j)', 'k)'))

        if eh_titulo:
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.first_line_indent = Cm(0)  # Títulos ficam na margem esquerda
            for r in p.runs:
                r.font.name = 'Arial'
                r.font.size = Pt(12)
                r.bold = True
        elif eh_lista:
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.first_line_indent = Cm(0)  # Listas não levam recuo de parágrafo
            for r in p.runs:
                r.font.name = 'Arial'
                r.font.size = Pt(11)
        else:
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.first_line_indent = Cm(1.25)  # Parágrafo padrão leva recuo de 1,25 cm
            for r in p.runs:
                r.font.name = 'Arial'
                r.font.size = Pt(11)

    # 3. Formatação das Tabelas e do Cabeçalho Institucional
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for p in celula.paragraphs:
                    p.paragraph_format.line_spacing = 1.0  # Espaçamento simples dentro de tabelas
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.first_line_indent = Cm(0)  # Sem recuo dentro de tabelas
                    for r in p.runs:
                        r.font.name = 'Arial'
                        r.font.size = Pt(10)
                        
    output = BytesIO()
    doc.save(output)
    return output.getvalue()

# ============================================================
# 🚀 INTERFACE GRÁFICA STREAMLIT
# ============================================================
with st.form("interface_auditoria_e_formatador"):
    arquivo_word = st.file_uploader("📂 Arraste o documento WORD (.docx) institucional AQUI", type=["docx"])
    enviado = st.form_submit_button("🔍 EXECUTAR AUDITORIA E FORMATAÇÃO", type="primary")

if enviado and arquivo_word:
    st.info(f"✅ Documento recebido: **{arquivo_word.name}**")
    
    # Inicializa variáveis para controle de escopo seguro
    rel = None
    conteudo_arquivo = arquivo_word.read()
    
    try:
        # Carrega o documento original na memória e roda a auditoria
        doc_original = docx.Document(BytesIO(conteudo_arquivo))
        rel = auditar_documento(doc_original)
        
        st.markdown("---")
        st.subheader("📋 RESULTADOS DA VALIDAÇÃO")
        
        c1, c2 = st.columns(2)
        with c1:
            st.info(f"**Tipo Identificado:** {rel['tipo']}")
            if rel['codigo']:
                st.success(f"**Código Localizado:** {rel['codigo']}")
            else:
                st.error("**Código:** ❌ NÃO ENCONTRADO NO CABEÇALHO")
        with c2:
            if rel['versao']:
                st.success(f"**Versão Localizada:** {rel['versao']}")
            else:
                st.error("**Versão:** ❌ NÃO ENCONTRADA NO CABEÇALHO")
            if rel['validade']:
                st.info(f"**Validade:** {rel['validade']}")
            else:
                st.info("**Validade:** ⚠️ Campo não preenchido")
        
        st.markdown("---")
        
        # Mostra na tela o status de cada seção avaliada
        st.subheader("👁️ STATUS DAS SEÇÕES EXIGIDAS")
        col_enc, col_fal = st.columns(2)
        
        with col_enc:
            st.markdown("#### ✅ Encontradas no texto")
            for s in rel["secoes_encontradas"]:
                st.success(f"• {s}")
                
        with col_fal:
            st.markdown("#### ❌ Faltantes ou incorretas")
            if rel["secoes_faltantes"]:
                for s in rel["secoes_faltantes"]:
                    st.error(f"• {s}")
            else:
                st.info("• Nenhuma seção ausente!")
                
    except Exception as e:
