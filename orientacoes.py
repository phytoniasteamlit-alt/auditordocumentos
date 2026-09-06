import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Triagem NAQH", page_icon="🛡️", layout="wide")
st.title("Triagem & Formatação Estrita com Ajuste de Layout - NAQH")
st.markdown("""
### 🧠 Inteligência Aplicada à Norma Zero (Versão Ajustada)
O sistema corrige desvios estéticos, remove espaços vazios fantasmas, autoajusta tabelas para não estourarem 
e garante que o texto do corpo não invada o cabeçalho.
""")

# --- 2. FLUXO DE CARREGAMENTO ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para triagem e formatação perfeita", type=["docx"])

if arquivo_word:
    doc = docx.Document(arquivo_word)
    
    # --- PROCESSAMENTO PRÉVIO: LIMPEZA DE PARÁGRAFOS FANTASMAS (ESPAÇOS EM BRANCO) ---
    # Remove linhas vazias extras que geram grandes clarões no documento ao aplicar espaçamento 1.5
    p_indices_remover = []
    for i, p in enumerate(doc.paragraphs):
        if not p.text.strip():
            p_indices_remover.append(i)
            
    # Remove de trás para frente para não quebrar os índices
    for index in sorted(p_indices_remover, reverse=True):
        p_to_remove = doc.paragraphs[index]
        p_element = p_to_remove._element
        p_element.getparent().remove(p_element)

    # Captura textual consolidada pós-limpeza
    texto_corpo_completo = ""
    for p in doc.paragraphs:
        texto_corpo_completo += " " + p.text.strip().upper()
        
    texto_tabelas_completo = ""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texto_tabelas_completo += " " + cell.text.strip().upper()

    texto_total_validacao = texto_corpo_completo + texto_tabelas_completo

    # --- ETAPA 1: IDENTIFICAÇÃO DO TIPO DE DOCUMENTO ---
    tipo_detectado = "NORMA"
    if "PROTOCOLO" in texto_total_validacao or "PROT_" in texto_total_validacao:
        tipo_detectado = "PROTOCOLO"
    elif "PROCEDIMENTO OPERACIONAL PADRÃO" in texto_total_validacao or "POP" in texto_total_validacao:
        tipo_detectado = "POP"
    elif "ROTINA" in texto_total_validacao or "ROT_" in texto_total_validacao:
        tipo_detectado = "ROTINA"

    st.write(f"📊 **Tipo de Documento Identificado:** {tipo_detectado}")

    # --- ETAPA 2: REQUISITOS SEGUNDO QUADRO 1 ---
    secoes_obrigatorias = {}
    if tipo_detectado == "PROTOCOLO":
        secoes_obrigatorias = {
            "OBJETIVO": "OBJETIVO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "REFERENCIAL TEÓRICO": "REFERENCIAL" in texto_total_validacao or "TEÓRICO" in texto_total_validacao,
            "ESTRATÉGIAS DE MONITORAMENTO": "MONITORAMENTO" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao or "REFERENCIA" in texto_total_validacao
        }
    elif tipo_detectado == "NORMA":
        secoes_obrigatorias = {
            "INTRODUÇÃO": "INTRODUÇÃO" in texto_total_validacao,
            "OBJETIVO": "OBJETIVO" in texto_total_validacao,
            "APLICABILIDADE": "APLICABILIDADE" in texto_total_validacao,
            "DESCRIÇÃO DA NORMA": "DESCRIÇÃO" in texto_total_validacao,
            "RESPONSÁVEL": "RESPONSÁVEL" in texto_total_validacao,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_total_validacao
        }

    # --- ETAPA 3: VALIDAÇÃO DO CABEÇALHO ---
    possui_codigo = "CÓDIGO:" in texto_total_validacao or "CODIGO:" in texto_total_validacao or "CÓDIGO" in texto_total_validacao or "PROT_" in texto_total_validacao
    possui_versao = ("VERSÃO" in texto_total_validacao or "VERSAO" in texto_total_validacao) and any(char.isdigit() for char in texto_total_validacao)

    erros_gravissimos = []
    if not possui_codigo:
        erros_gravissimos.append("❌ **IDENTIFICAÇÃO AUSENTE:** Campo **'CÓDIGO'** incompleto ou ausente.")
    if not possui_versao:
        erros_gravissimos.append("❌ **IDENTIFICAÇÃO AUSENTE:** Campo **'VERSÃO'** ausente ou incorreto.")
        
    for secao, encontrada in secoes_obrigatorias.items():
        if not encontrada:
            erros_gravissimos.append(f"❌ **OMISSÃO DE SEÇÃO CRÍTICA (Modelo {tipo_detectado}):** Seção **'{secao}'** não localizada.")

    # --- ETAPA 4: PROCESSAMENTO E CORREÇÃO DE LAYOUT ---
    if not erros_gravissimos:
        st.success("✅ **TRIAGEM CONCLUÍDA!** Gerando arquivo com layout otimizado...")

        # 1. Configuração de Margens Estritas (3x3x2x2)
        for section in doc.sections:
            section.top_margin = Cm(3.0)     
            section.left_margin = Cm(3.0)    
            section.bottom_margin = Cm(2.0)  
            section.right_margin = Cm(2.0)   

        # 2. Formatação do Corpo do Texto (Calibri 11, Justificado, 1.5, Recuo 1,25cm)
        for paragraph in doc.paragraphs:
            texto_limpo = paragraph.text.upper().strip()
            if "REFERÊNCIAS" in texto_limpo or "REFERENCIA" in texto_limpo:
                continue
                
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.first_line_indent = Cm(1.25)
            
            # Força o distanciamento correto para evitar que o texto invada tabelas superiores
            paragraph.paragraph_format.space_before = Pt(3)
            paragraph.paragraph_format.space_after = Pt(3)
            
            for run in paragraph.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        # 3. Ajuste Avançado de Tabelas (Previne estouros laterais e desalinhamentos)
        largura_maxima_tabela = Cm(16.0) # Largura útil total da página
        for table in doc.tables:
            # Força a tabela a se autoajustar e respeitar o limite físico da página
            table.autofit = True
            table.allow_autofit = True
            
            for i, row in enumerate(table.rows):
                for cell in row.cells:
                    # Define largura estrita por célula para evitar quebras esquisitas
                    cell.width = largura_maxima_tabela
                    
                    for paragraph in cell.paragraphs:
                        # Previne que o parágrafo dentro da tabela herde recuos do corpo do texto
                        paragraph.paragraph_format.first_line_indent = Cm(0)
                        paragraph.paragraph_format.space_before = Pt(2)
                        paragraph.paragraph_format.space_after = Pt(2)
                        
                        if i == 0:  # Títulos
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            for run in paragraph.runs:
                                run.font.name = 'Calibri'
                                run.font.size = Pt(10)
                                run.font.bold = True
                        else:  # Dados
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            for run in paragraph.runs:
                                run.font.name = 'Calibri'
                                run.font.size = Pt(9)
                                run.font.bold = False

        # 4. Formatação de Referências (Calibri 10, Esquerda)
        capturar_referencias = False
        for paragraph in doc.paragraphs:
            if "REFERÊNCIAS" in paragraph.text.upper().strip():
                capturar_referencias = True
                continue
            if capturar_referencias and paragraph.text.strip():
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                paragraph.paragraph_format.first_line_indent = Cm(0)
                for run in paragraph.runs:
                    run.font.name = 'Calibri'
                    run.font.size = Pt(10)

        # 5. Redimensionamento de Imagens
        for shape in doc.inline_shapes:
            if shape.width > largura_maxima_tabela:
                proporcao = largura_maxima_tabela / shape.width
                shape.width = largura_maxima_tabela
                shape.height = int(shape.height * proporcao)

        # Buffer para download
        conteudo_corrigido = BytesIO()
        doc.save(conteudo_corrigido)
        conteudo_corrigido.seek(0)

        st.balloons()
        st.download_button(
            label="📥 Baixar Documento com Layout Corrigido (.docx)",
            data=conteudo_corrigido,
            file_name=f"PERFEITO_NAQH_{arquivo_word.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("🚨 **DOCUMENTO RETIDO NA TRIAGEM INICIAL**")
        for erro in erros_gravissimos:
            st.markdown(erro)
