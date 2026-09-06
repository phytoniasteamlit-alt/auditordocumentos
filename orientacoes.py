import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Triagem NAQH", page_icon="🛡️", layout="wide")
st.title("Triagem & Formatação Automatizada à Risca - NAQH")
st.markdown("""
### 🧠 Foco no Conteúdo, Automação Absoluta na Forma
As profissionais escrevem apenas o conteúdo. O sistema valida a presença das seções obrigatórias 
e aplica todas as regras tipográficas da Norma Zero de forma estrita (Corpo, Tabelas e Referências).
""")

NOMES_TIPOS = {
    "NOR": "NORMA (NOR)",
    "POP": "PROCEDIMENTO OPERACIONAL PADRÃO (POP)",
    "PROT": "PROTOCOLO (PROT)",
    "MAN": "MANUAL (MAN)",
    "REG": "REGIMENTO INTERNO (REG)",
    "ROT": "ROTINA (ROT)",
    "POL": "POLÍTICA INSTITUCIONAL (POL)"
}

# --- 2. BARRA LATERAL ---
st.sidebar.header("⚙️ Controles de Auditoria (NAQH)")
tem_impressos_inconformes = st.sidebar.checkbox("⚠️ Há imagens de impressos ilegíveis ou fora do padrão?")

# --- 3. FLUXO DE CARREGAMENTO ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para triagem e formatação automática", type=["docx"])

if arquivo_word:
    doc = docx.Document(arquivo_word)
    
    # Mapeamento de Seções Críticas do Corpo do Texto conforme a Norma Zero
    secoes_obrigatorias = {
        "INTRODUÇÃO": False,
        "OBJETIVO": False,
        "APLICABILIDADE": False,
        "DESCRIÇÃO": False,
        "RESPONSÁVEIS": False,
        "REFERÊNCIAS": False
    }
    
    # Rastreamento de Identificação Obrigatória na Triagem (Código e Versão)
    controle_obrigatorio = {
        "CÓDIGO": False,
        "VERSÃO": False
    }
    
    erros_gravissimos = []
    erros_visuais = []

    # --- ETAPA 1: TRIAGEM TEXTUAL DO CORPO ---
    for paragraph in doc.paragraphs:
        texto_paragrafo = paragraph.text.strip().upper()
        
        if "INTRODUÇÃO" in texto_paragrafo or "INTRODUCAO" in texto_paragrafo:
            secoes_obrigatorias["INTRODUÇÃO"] = True
        if "OBJETIVO" in texto_paragrafo:
            secoes_obrigatorias["OBJETIVO"] = True
        if "APLICABILIDADE" in texto_paragrafo:
            secoes_obrigatorias["APLICABILIDADE"] = True
        if "DESCRIÇÃO" in texto_paragrafo or "DESCRICAO" in texto_paragrafo:
            secoes_obrigatorias["DESCRIÇÃO"] = True
        if "RESPONSÁVEIS" in texto_paragrafo or "RESPONSAVEIS" in texto_paragrafo:
            secoes_obrigatorias["RESPONSÁVEIS"] = True
        if "REFERÊNCIAS" in texto_paragrafo or "REFERENCIA" in texto_paragrafo:
            secoes_obrigatorias["REFERÊNCIAS"] = True

    # --- ETAPA 2: TRIAGEM DE IDENTIFICAÇÃO (CABEÇALHOS/TABELAS) ---
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texto_celula = cell.text.strip().upper()
                
                if "CÓDIGO:" in texto_celula or "CODIGO:" in texto_celula:
                    if len(texto_celula) > 8 or "_" in texto_celula or "-" in texto_celula:
                        controle_obrigatorio["CÓDIGO"] = True
                        
                if "VERSÃO:" in texto_celula or "VERSAO:" in texto_celula:
                    if any(char.isdigit() for char in texto_celula):
                        controle_obrigatorio["VERSÃO"] = True

    # --- ETAPA 3: VALIDAÇÃO DOS CRITÉRIOS DE RETENÇÃO ---
    for campo, preenchido in controle_obrigatorio.items():
        if not preenchido:
            erros_gravissimos.append(f"❌ **IDENTIFICAÇÃO AUSENTE:** O campo obrigatório **'{campo}'** não foi preenchido ou está incompleto no cabeçalho.")
            
    for secao, encontrada in secoes_obrigatorias.items():
        if not encontrada:
            erros_gravissimos.append(f"❌ **OMISSÃO DE SEÇÃO CRÍTICA:** A seção obrigatória **'{secao}'** não foi localizada no corpo do documento.")

    # --- ETAPA 4: PROCESSAMENTO E FORMATAÇÃO (SE PASSAR NA TRIAGEM) ---
    if not erros_gravissimos:
        st.success("✅ **TRIAGEM INICIAL APROVADA:** Estrutura básica validada. Aplicando correções estéticas...")

        # 1. Correção Automática de Margens (Regra Atualizada: 3x3x2x2)
        for section in doc.sections:
            section.top_margin = Cm(3.0)     
            section.left_margin = Cm(3.0)    
            section.bottom_margin = Cm(2.0)  
            section.right_margin = Cm(2.0)   

        # 2. Correção Automática do Corpo do Texto (Calibri 11, Justificado, 1.5, Recuo 1,25cm)
        for paragraph in doc.paragraphs:
            if not paragraph.text.strip():
                continue
            texto_limpo = paragraph.text.upper().strip()
            
            # Pula as referências para aplicar a regra de tamanho 10 separadamente
            if "REFERÊNCIAS" in texto_limpo or "REFERENCIA" in texto_limpo:
                continue
                
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.first_line_indent = Cm(1.25)
            
            for run in paragraph.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        # 3. Correção Automática de Tabelas (Títulos: Calibri 10 Negrito | Dados: Calibri 9 Justificado)
        for table in doc.tables:
            for i, row in enumerate(table.rows):
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if i == 0:  # Títulos das Colunas
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            for run in paragraph.runs:
                                run.font.name = 'Calibri'
                                run.font.size = Pt(10)
                                run.font.bold = True
                        else:  # Dados e Histórico
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            for run in paragraph.runs:
                                run.font.name = 'Calibri'
                                run.font.size = Pt(9)
                                run.font.bold = False

        # 4. ISOLAMENTO E FORMATAÇÃO DAS REFERÊNCIAS (Calibri 10, Esquerda, Sem recuo)
        capturar_referencias = False
        for paragraph in doc.paragraphs:
            if "REFERÊNCIAS" in paragraph.text.upper().strip():
                capturar_referencias = True
                continue
            if capturar_referencias and paragraph.text.strip():
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                paragraph.paragraph_format.first_line_indent = Cm(0)  # Referência não tem recuo de parágrafo
                for run in paragraph.runs:
                    run.font.name = 'Calibri'
                    run.font.size = Pt(10)

        # 5. Organização de Imagens e Prints (Limitação de margem útil para 16cm)
        largura_maxima_util = Cm(16.0)
        if len(doc.inline_shapes) > 0:
            for shape in doc.inline_shapes:
                if shape.width > largura_maxima_util:
                    proporcao = largura_maxima_util / shape.width
                    shape.width = largura_maxima_util
                    shape.height = int(shape.height * proporcao)
            erros_visuais.append("📸 **Enquadramento de Prints:** Imagens e fluxogramas foram redimensionados para não estourarem o layout físico do papel.")

        # Geração do arquivo modificado em memória para download
        conteudo_corrigido = BytesIO()
        doc.save(conteudo_corrigido)
        conteudo_corrigido.seek(0)

        st.balloons()
        st.markdown("### ✨ O documento está perfeitamente formatado e pronto para a análise das meninas!")
        
        st.download_button(
            label="📥 Baixar Documento Formatado para Verificação (.docx)",
            data=conteudo_corrigido,
            file_name=f"TRIADO_NAQH_{arquivo_word.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        if erros_visuais or tem_impressos_inconformes:
            st.markdown("---")
            st.subheader("📋 Relatório Secundário para o Auditor")
            for aviso in erros_visuais:
                st.markdown(aviso)

    # --- ETAPA 5: PAINEL DE RETENÇÃO (DOCUMENTO REJEITADO) ---
    else:
        st.markdown("---")
        st.error("🚨 **DOCUMENTO RETIDO NA TRIAGEM INICIAL**")
        st.subheader("⛔ Erro Estrutural Grave: O arquivo não possui os requisitos mínimos de conteúdo para seguir adiante.")
        st.markdown("### 📋 Relatório de Devolução ao Setor de Origem:")
        
        for erro in erros_gravissimos:
            st.markdown(erro)
