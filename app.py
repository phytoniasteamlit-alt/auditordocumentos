import streamlit as st
import pandas as pd
import openpyxl
import unicodedata
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="PAINEL DE INDICADORES NORMA ZERO", page_icon="📝", layout="wide")

# --- BLOCO DE CUSTOMIZAÇÃO VISUAL AVANÇADA (CSS INJECT NATIVO) ---
st.markdown("""
<style>
[data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th { font-size: 16px !important; font-weight: 600 !important; color: #FFFFFF !important; }
[data-testid="stDataFrame"] td:last-child { color: #FBBF24 !important; font-size: 18px !important; font-weight: bold !important; }
div[data-testid="stSelectbox"] label p, div[data-testid="stFileUploader"] label p { font-size: 15px !important; font-weight: bold !important; color: #38BDF8 !important; }
section[data-testid="stFileUploaderDropzone"] { border: 2px dashed #38BDF8 !important; background-color: #1E293B !important; }
/* Rolagem interna controlada para a tabela da barra lateral */
[data-testid="stSidebar"] [data-testid="stDataFrame"] { max-height: 220px !important; overflow-y: auto !important; }
</style>
""", unsafe_allow_html=True)

# --- CABEÇALHO DO HOSPITAL ---
col_titulo, col_hospital = st.columns([0.65, 0.35])

with col_titulo:
    st.markdown("# PAINEL DE INDICADORES NORMA ZERO")

with col_hospital:
    st.markdown("""
    <div style="display: flex; flex-direction: column; align-items: flex-end; gap: 4px; margin-top: 10px;">
        <div style="display: flex; align-items: center; gap: 10px;">
            <span style="font-size: 34px; color: #EF4444; font-weight: bold; line-height: 1;">🏥</span>
            <span style="font-size: 22px; color: #FFFFFF; font-weight: 800; letter-spacing: 0.5px;">HOSPITAL DA CIDADE</span>
        </div>
        <div style="display: flex; align-items: center; gap: 6px; margin-right: 2px;">
            <span style="font-size: 20px; line-height: 1;">👩‍💼</span>
            <span style="font-size: 15px; color: #94A3B8; font-weight: 600; letter-spacing: 0.3px;">Coord.: Fabrícia Rocha</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

#--- 2. CARREGAMENTO DO ARQUIVO ---
st.sidebar.header("⚙️ Entrada de Dados")
arquivo_excel = st.sidebar.file_uploader("📁 Carregar Planilha Excel (.xlsx):", type=["xlsx"], key="uploader_xlsx")

df_base = pd.DataFrame()
media_v1, media_v2 = 0.0, 0.0

def remover_acentos(texto):
    if pd.isna(texto): return ""
    return "".join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn').upper().strip()

if arquivo_excel:
    try:
        xl = pd.ExcelFile(arquivo_excel, engine="openpyxl")
        lista_abas_reais = xl.sheet_names
        
        # PROCURA DIRETAMENTE PELA ABA "DADOS_GRÁFICOS" INFORMADA
        nome_aba_principal = None
        for n_real in lista_abas_reais:
            if "GRAF" in str(n_real).upper().strip():
                nome_aba_principal = n_real
                break
        
        # Caso a aba mude de nome futuramente, busca um termo alternativo coerente
        if not nome_aba_principal:
            for n_real in lista_abas_reais:
                if "DADOS" in str(n_real).upper().strip():
                    nome_aba_principal = n_real
                    break
        
        if not nome_aba_principal and lista_abas_reais:
            nome_aba_principal = lista_abas_reais[0]
        
        if nome_aba_principal:
            # Lê estritamente a aba selecionada mapeando a primeira linha como cabeçalho
            df_raw = pd.read_excel(arquivo_excel, sheet_name=nome_aba_principal, header=0, engine="openpyxl")
            
            # Limpa espaços fantasmas nos cabeçalhos das colunas
            df_raw.columns = [str(c).strip() for c in df_raw.columns]
            colunas_planilha = {str(col).upper().strip(): col for col in df_raw.columns}
            
            def buscar_coluna(nomes_possiveis, padrao="Não Informado"):
                for nome in nomes_possiveis:
                    if nome in colunas_planilha:
                        return df_raw[colunas_planilha[nome]]
                return pd.Series([padrao] * len(df_raw))

            # 1. CORREÇÃO CRÍTICA DO PROCESSAMENTO DAS MÉDIAS (Substituído .shape por len das colunas)
            total_colunas = len(df_raw.columns)
            
            if total_colunas > 6:
                s_g = df_raw.iloc[:, 6].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False)
                df_g_nums = pd.to_numeric(s_g, errors='coerce').dropna()
                df_g_filtrado = df_g_nums[(df_g_nums >= 0) & (df_g_nums < 365)]
                media_v1 = df_g_filtrado.mean() if not df_g_filtrado.empty else 0.0
            
            if total_colunas > 7:
                s_h = df_raw.iloc[:, 7].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False)
                df_h_nums = pd.to_numeric(s_h, errors='coerce').dropna()
                df_h_filtrado = df_h_nums[(df_h_nums >= 0) & (df_h_nums < 365)]
                media_v2 = df_h_filtrado.mean() if not df_h_filtrado.empty else 0.0
            
            media_v1 = float(media_v1) if pd.notna(media_v1) else 0.0
            media_v2 = float(media_v2) if pd.notna(media_v2) else 0.0

            # 2. CAPTURA DOS DADOS MAPEADOS DE FORMA SEGURA POR NOMES DE CABEÇALHO
            df_base = pd.DataFrame()
            df_base["SIGLA"] = buscar_coluna(["SIGLA DO DOCUMENTO", "SIGLA"], "N/A")
            df_base["SETOR"] = buscar_coluna(["SETOR"], "N/A")
            df_base["RESPONSAVEL"] = buscar_coluna(["RESPONSAVEL", "RESPONSÁVEL"], "Não Informado")
            df_base["STATUS"] = buscar_coluna(["STATUS DO DOCUMENTO NORMATIVO", "STATUS"], "Não Informado")
            df_base["SIT_PRAZO"] = buscar_coluna(["VENCIDO, NO PRAZO, PRESTES A VENCER", "(VENCIDO, NO PRAZO, PRESTES A VENCER)", "SIT_PRAZO"], "Não Informado")
            
            # Limpeza inicial de linhas em branco do Excel
            df_base = df_base.dropna(subset=["SIGLA", "STATUS"], how="all")
            for col in df_base.columns:
                df_base[col] = df_base[col].fillna("").astype(str).str.strip()
            
            # Filtro para remover registros com lixo estrutural
            valores_vazios = ["0", "0.0", "NAN", "NONE", "", "NAN NAN", "NÃO INFORMADO", "#VALOR!", "SIGLA DO DOCUMENTO", "STATUS DO DOCUMENTO NORMATIVO"]
            for col in df_base.columns:
                df_base = df_base[~df_base[col].str.upper().isin(valores_vazios)]
            
            # Substituição e padronização das colaboradoras desligadas
            if "RESPONSAVEL" in df_base.columns:
                df_base["RESPONSAVEL"] = df_base["RESPONSAVEL"].apply(
                    lambda x: "Antigo Colaborador" if str(x).upper().strip() in ["SABRINA", "SONALHYA"] else str(x).upper().strip()
                )
            
            if "STATUS" in df_base.columns:
                df_base["STATUS"] = df_base["STATUS"].apply(
                    lambda x: "AG Aguardando" if "VERIFICADO AGU" in x.upper() or "AGUARDANDO" in x.upper() else str(x).upper().strip()
                )
                
    except Exception as e:
        st.error(f"Erro crítico no processamento dos dados da planilha: {e}")

#--- 3. MENUS LATERAIS DE FILTROS ---
if not df_base.empty:
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtros de Visualização")
    
    lista_responsaveis = sorted([str(r) for r in df_base["RESPONSAVEL"].dropna().unique() if str(r).strip() != ""])
    lista_responsaveis.insert(0, "Todos")
    responsavel_selecionado = st.sidebar.selectbox("Selecione o Responsável:", lista_responsaveis)
    
    lista_documentos = sorted([str(d) for d in df_base["SIGLA"].dropna().unique() if str(d).strip() != ""])
    lista_documentos.insert(0, "Todos")
    documento_selecionado = st.sidebar.selectbox("Filtrar por Documento Aprovado:", lista_documentos)
    
    df_filtrado = df_base.copy()
    if responsavel_selecionado != "Todos":
        df_filtrado = df_filtrado[df_filtrado["RESPONSAVEL"] == responsavel_selecionado]
    if documento_selecionado != "Todos":
        df_filtrado = df_filtrado[df_filtrado["SIGLA"] == documento_selecionado]
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("📊 Qtd por Documento (Sigla)")
    df_lateral_contagem = df_filtrado["SIGLA"].dropna().value_counts().reset_index()
    df_lateral_contagem.columns = ["Documento", "Qtd"]
    st.sidebar.dataframe(df_lateral_contagem, use_container_width=True, hide_index=True, height=180)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("🎨 Configuração dos Gráficos")
    c_g1_color = st.sidebar.color_picker("G1 - Cor Status:", "#FBBF24", key="c1")
    c_g2_color = st.sidebar.color_picker("G2 - Cor Tempo:", "#38BDF8", key="c2")
    c_g4_color = st.sidebar.color_picker("G4 - Cor Prazo:", "#C084FC", key="c4")
    
    #--- 4. INDICADORES DO TOPO (CARDS) ---
    total_docs = len(df_filtrado)
    aprovados = len(df_filtrado[df_filtrado["STATUS"].str.contains("APROVADO|OK|SIM", na=False)])
    
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    with kpi_col1: st.metric(label="📊 TOTAL DOCUMENTOS", value=f"{total_docs}")
    with kpi_col2: st.metric(label="✅ APROVADOS", value=f"{aprovados}")
    with kpi_col3: st.metric(label="⏳ 1º VERF.", value=f"{media_v1:.1f} d")
    with kpi_col4: st.metric(label="⏳ 2º VERF.", value=f"{media_v2:.1f} d")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    #--- 5. RENDERIZAÇÃO DOS GRÁFICOS ---
    linha1_col1, linha1_col2 = st.columns(2)
    
    with linha1_col1:
        st.markdown("### Documentos por Status")
        dados_g1 = df_filtrado["STATUS"].dropna().value_counts()
        st.bar_chart(dados_g1, color=c_g1_color, horizontal=True)
    
    with linha1_col2:
        st.markdown("### Tempo de Análise")
        dados_g2 = pd.Series({"1º Verf.": media_v1, "2º Verf.": media_v2, "Total Estimado": (media_v1 + media_v2)})
