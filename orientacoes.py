import streamlit as st
import docx
from docx.shared import Cm
from io import BytesIO
import re

# --- 1. CONFIGURAÇÃO DA PÁGINA STREAMLIT ---
st.set_page_config(page_title="Formatador de Documentos NAQH", page_icon="📊", layout="wide")

# Menu Lateral - Identificação Visual do Operador (Ezequias Santos Agt Administrativo)
with st.sidebar:
    st.markdown("### 🧑‍💻 Operador")
    st.markdown("**Ezequias Santos**\n*Agt Administrativo*")
    st.divider()

st.title("Triagem Avançada & Formatador Automático - NAQH")
st.markdown("""
### 🧠 Inteligência Multidocumento e Controle de Atualizações
O sistema realiza a triagem dinâmica identificando o tipo de documento de forma segura, **preservando 100% o alinhamento nativo de listas, cabeçalhos e subseções do arquivo original**.
""")

# --- 2. MOTOR DE PRESERVAÇÃO INTEGRAL (PRESERVA TEXTOS E ESTILOS NATIVOS) ---
def aplicar_formatacao_protegida(arquivo_bytes):
    # Abre o documento mantendo todas as suas estruturas originais intactas
    doc = docx.Document(arquivo_bytes)
    
    # ATUALIZAÇÃO GEOMÉTRICA DE MARGENS (Norma Zero: Sup 2 / Inf 2 / Esq 2 / Dir 3)
    # Altera os limites da página diretamente no XML estrutural sem reformatar o texto corrido
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(3.0)

    # Nota: Removemos qualquer alteração forçada de parágrafos ou loops cegos sobre doc.paragraphs.
    # Isso garante que as tabelas de cabeçalho flutuantes, numerações e marcadores nativos
    # permaneçam exatamente na mesma posição em que foram criados no arquivo original.

    conteudo_saida = BytesIO()
    doc.save(conteudo_saida)
    conteudo_saida.seek(0)
    return conteudo_saida.getvalue()

# --- 3. FLUXO DE PROCESSAMENTO ---
arquivo_word = st.file_uploader("Arraste o documento WORD (.docx) aqui para Triagem e Formatação", type=["docx"])

if arquivo_word:
    doc_triagem = docx.Document(arquivo_word)
    
    # Varredura veloz e superficial para triagem de metadados
    texto_corpo_raw = " ".join([p.text.strip() for p in doc_triagem.paragraphs[:30]])
    texto_tabelas_raw = " ".join([cell.text.strip() for t in doc_triagem.tables[:1] for r in t.rows for cell in r.cells])
    texto_total_raw = texto_corpo_raw + " " + texto_tabelas_raw
    texto_total_validacao = texto_total_raw.upper()
    
    tipo_detectado = "PROTOCOLO"
    if "PROCEDIMENTO OPERACIONAL" in texto_total_validacao or "POP" in texto_total_validacao:
        tipo_detectado = "POP"
        
    st.write(f"📋 **Tipo de Documento Identificado:** `{tipo_detectado}`")
    
    codigo_doc = "PROT_SCIH005"
    match_codigo = re.search(r'\b(PROT|POP|MAN|NOR|ROT)_[A-Z0-9_\s-]+\b', texto_total_raw, re.IGNORECASE)
    if match_codigo:
        codigo_doc = match_codigo.group(0).strip().upper().replace(" ", "")

    # Gera a saída preservando a integridade original de cabeçalhos e textos
    dados_finais = aplicar_formatacao_protegida(arquivo_word)

    st.download_button(
        label="📥 CLIQUE AQUI PARA BAIXAR O DOCUMENTO FORMATADO",
        data=dados_finais,
        file_name=f"{codigo_doc}_Formatado_Homologado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.success(f"✅ **TRIAGEM CONCLUÍDA COM SUCESSO!** Margens ajustadas para o padrão institucional com preservação total de textos, logos e cabeçalhos.")
