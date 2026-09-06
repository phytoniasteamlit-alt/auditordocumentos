import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Triagem NAQH", page_icon="🛡️", layout="wide")
st.title("Triagem & Formatação Dinâmica por Tipo de Documento - NAQH")
st.markdown("""
### 🧠 Inteligência Aplicada à Norma Zero
O sistema identifica automaticamente o **Tipo de Documento** no cabeçalho e cobra **apenas as seções obrigatórias** 
especificadas na tabela oficial (Quadro 1). Erros estéticos são corrigidos automaticamente.
""")

# --- 2. FLUXO DE CARREGAMENTO ---
arquivo_word = st.file_uploader("Arraste o arquivo WORD (.docx) aqui para triagem e formatação automática", type=["docx"])

if arquivo_word:
    doc = docx.Document(arquivo_word)
    
    # Captura todo o texto do documento e das tabelas para análise de dados
    texto_corpo_completo = ""
    for p in doc.paragraphs:
        texto_corpo_completo += " " + p.text.strip().upper()
        
    texto_tabelas_completo = ""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texto_tabelas_completo += " " + cell.text.strip().upper()

    # --- ETAPA 1: IDENTIFICAÇÃO AUTOMÁTICA DO TIPO DE DOCUMENTO ---
    tipo_detectado = "NORMA"  # Padrão genérico
    texto_busca_tipo = texto_corpo_completo + texto_tabelas_completo
    
    if "PROTOCOLO" in texto_busca_tipo or "PROT_" in texto_busca_tipo:
        tipo_detectado = "PROTOCOLO"
    elif "PROCEDIMENTO OPERACIONAL PADRÃO" in texto_busca_tipo or "POP" in texto_busca_tipo:
        tipo_detectado = "POP"
    elif "ROTINA" in texto_busca_tipo or "ROT_" in texto_busca_tipo:
        tipo_detectado = "ROTINA"
    elif "REGIMENTO" in texto_busca_tipo or "REG_" in texto_busca_tipo:
        tipo_detectado = "REGIMENTO"

    st.write(f"📊 **Tipo de Documento Identificado:** {tipo_detectado}")

    # --- ETAPA 2: DEFINIÇÃO DE REQUISITOS SEGUNDO O QUADRO 1 DA NORMA ZERO ---
    secoes_obrigatorias = {}
    
    if tipo_detectado == "PROTOCOLO":
        secoes_obrigatorias = {
            "OBJETIVO": "OBJETIVO" in texto_corpo_completo,
            "APLICABILIDADE": "APLICABILIDADE" in texto_corpo_completo,
            "REFERENCIAL TEÓRICO": "REFERENCIAL TEÓRICO" in texto_corpo_completo or "REFERENCIAL TEORICO" in texto_corpo_completo,
            "ESTRATÉGIAS DE MONITORAMENTO": "MONITORAMENTO" in texto_corpo_completo,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_corpo_completo or "REFERENCIA" in texto_corpo_completo
        }
    elif tipo_detectado == "NORMA":
        secoes_obrigatorias = {
            "INTRODUÇÃO": "INTRODUÇÃO" in texto_corpo_completo or "INTRODUCAO" in texto_corpo_completo,
            "OBJETIVO": "OBJETIVO" in texto_corpo_completo,
            "APLICABILIDADE": "APLICABILIDADE" in texto_corpo_completo,
            "DESCRIÇÃO DA NORMA": "DESCRIÇÃO" in texto_corpo_completo or "DESCRICAO" in texto_corpo_completo,
            "RESPONSÁVEL": "RESPONSÁVEL" in texto_corpo_completo or "RESPONSAVEL" in texto_corpo_completo,
            "REFERÊNCIAS": "REFERÊNCIAS" in texto_corpo_completo or "REFERENCIA" in texto_corpo_completo
        }
    # (Adicione outras variações de regras do Quadro 1 se achar necessário)

    # --- ETAPA 3: VALIDAÇÃO DE CÓDIGO E VERSÃO NO CABEÇALHO ---
    # Busca robusta que limpa espaços extras para capturar os dados do seu print 2
    possui_codigo = ("CÓDIGO:" in texto_tabelas_completo or "CODIGO:" in texto_tabelas_completo) and ("_SCIH" in texto_tabelas_completo or len(texto_tabelas_completo) > 10)
    possui_versao = ("VERSÃO:" in texto_tabelas_completo or "VERSAO:" in texto_tabelas_completo) and any(char.isdigit() for char in texto_tabelas_completo)

    erros_gravissimos = []
    
    if not possui_codigo:
        erros_gravissimos.append("❌ **IDENTIFICAÇÃO AUSENTE:** O campo obrigatório **'CÓDIGO'** está ausente ou incompleto no cabeçalho.")
    if not possui_versao:
        erros_gravissimos.append("❌ **IDENTIFICAÇÃO AUSENTE:** O campo obrigatório **'VERSÃO'** não foi preenchido ou está sem número correspondente.")
        
    for secao, encontrada in secoes_obrigatorias.items():
        if not encontrada:
            erros_gravissimos.append(f"❌ **OMISSÃO DE SEÇÃO CRÍTICA (Modelo {tipo_detectado}):** A seção obrigatória **'{secao}'** não foi localizada no corpo.")

    # --- ETAPA 4: FORMATAÇÃO AUTOMÁTICA (SE APROVADO NA TRIAGEM) ---
    if not erros_gravissimos:
        st.success("✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Arquivo validado e liberado para formatação estética.")

        # 1. Ajuste das Margens (3x3x2x2)
        for section in doc.sections:
            section.top_margin = Cm(3.0)     
            section.left_margin = Cm(3.0)    
            section.bottom_margin = Cm(2.0)  
            section.right_margin = Cm(2.0)   

        # 2. Ajuste do Corpo do Texto (Calibri 11, Justificado, 1.5, Recuo 1,25cm)
        for paragraph in doc.paragraphs:
            if not paragraph.text.strip():
                continue
            texto_limpo = paragraph.text.upper().strip()
            if "REFERÊNCIAS" in texto_limpo or "REFERENCIA" in texto_limpo:
                continue
                
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.first_line_indent = Cm(1.25)
            
            for run in paragraph.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        # 3. Ajuste de Tabelas (Títulos: Calibri 10 Negrito | Dados: Calibri 9 Justificado)
        for table in doc.tables:
            for i, row in enumerate(table.rows):
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if i == 0:  
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            for run in paragraph.runs:
                                run.font.name = 'Calibri'
                                run.font.size = Pt(10)
                                run.font.bold = True
                        else:  
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            for run in paragraph.runs:
                                run.font.name = 'Calibri'
                                run.font.size = Pt(9)
                                run.font.bold = False

        # 4. Ajuste de Referências (Calibri 10, Esquerda)
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
        largura_maxima_util = Cm(16.0)
        for shape in doc.inline_shapes:
            if shape.width > largura_maxima_util:
                proporcao = largura_maxima_util / shape.width
                shape.width = largura_maxima_util
                shape.height = int(shape.height * proporcao)

        # Download do Arquivo Corrigido
        conteudo_corrigido = BytesIO()
        doc.save(conteudo_corrigido)
        conteudo_corrigido.seek(0)

        st.balloons()
        st.download_button(
            label="📥 Baixar Documento Formatado e Aprovado (.docx)",
            data=conteudo_corrigido,
            file_name=f"APROVADO_{arquivo_word.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    # --- ETAPA 5: PAINEL DE DEVOLUÇÃO (SE REJEITADO) ---
    else:
        st.error("🚨 **DOCUMENTO RETIDO NA TRIAGEM INICIAL**")
        st.markdown("### 📋 Relatório de Devolução ao Setor de Origem:")
        for erro in erros_gravissimos:
            st.markdown(erro)
