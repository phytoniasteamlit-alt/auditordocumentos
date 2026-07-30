import streamlit as st
import pandas as pd
import openpyxl
import unicodedata
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="PAINEL DE INDICADORES NORMA ZERO", page_icon="📋", layout="wide")

# --- BLOCO DE CUSTOMIZAÇÃO VISUAL AVANÇADA (CSS INJECT NATIVO) ---
# ALTERADO: Adicionada regra para travar o tamanho máximo da tabela lateral e não quebrar o painel
st.markdown("""
<style>
[data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th { font-size: 16px !important; font-weight: 600 !important; color: #FFFFFF !important; }
[data-testid="stDataFrame"] td:last-child { color: #FBBF24 !important; font-size: 18px !important; font-weight: bold !important; }
div[data-testid="stSelectbox"] label p, div[data-testid="stFileUploader"] label p { font-size: 15px !important; font-weight: bold !important; color: #38BDF8 !important; }
section[data-testid="stFileUploaderDropzone"] { border: 2px dashed #38BDF8 !important; background-color: #1E293B !important; }
/* Força a tabela da barra lateral a ter tamanho controlado */
[data-testid="stSidebar"] [data-testid="stDataFrame"] { max-height: 250px !important; overflow-y: auto !important; }
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
        
        nome_aba_principal = None
        for n_real in lista_abas_reais:
            n_up = str(n_real).upper().strip()
            if "VERF" in n_up or "ACOMP" in n_up:
                nome_aba_principal = n_real
                break
        
        if not nome_aba_principal and lista_abas_reais:
            nome_aba_principal = lista_abas_reais
            
        if nome_aba_principal:
            df_raw = pd.read_excel(arquivo_excel, sheet_name=nome_aba_principal, header=1, engine="openpyxl")
            df_raw.columns = df_raw.columns.astype(str).str.strip().str.upper()
            
            # 1. PROCESSAMENTO DAS MÉDIAS
            col_g, col_h = None, None
            for c in df_raw.columns:
                if "1º" in c or "1O" in c or "V 1" in c or "V1" in c: col_g = c
                if "2º" in c or "2O" in c or "V 2" in c or "V2" in c: col_h = c
            
            if col_g:
                s_g = df_raw[col_g].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False)
                df_g_nums = pd.to_numeric(s_g, errors='coerce').dropna()
                df_g_filtrado = df_g_nums[(df_g_nums >= 0) & (df_g_nums < 365)]
                media_v1 = df_g_filtrado.mean() if not df_g_filtrado.empty else 0.0
                
            if col_h:
                s_h = df_raw[col_h].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False)
                df_h_nums = pd.to_numeric(s_h, errors='coerce').dropna()
                df_h_filtrado = df_h_nums[(df_h_nums >= 0) & (df_h_nums < 365)]
                media_v2 = df_h_filtrado.mean() if not df_h_filtrado.empty else 0.0
            
            media_v1 = float(media_v1) if pd.notna(media_v1) else 0.0
            media_v2 = float(media_v2) if pd.notna(media_v2) else 0.0

            # 2. CAPTURA DOS DADOS DOS GRÁFICOS
            df_base = pd.DataFrame()
            for col_real in df_raw.columns:
                if "SIGLA" in col_real: df_base["SIGLA"] = df_raw[col_real]
                if "SETOR" in col_real: df_base["SETOR"] = df_raw[col_real]
                if "RESPONS" in col_real: df_base["RESPONSAVEL"] = df_raw[col_real]
                if "STATUS" in col_real: df_base["STATUS"] = df_raw[col_real]
                if "PRAZO" in col_real or "SIT" in col_real or "VENCIDO" in col_real: df_base["SIT_PRAZO"] = df_raw[col_real]

            num_colunas = len(df_raw.columns)
            if "SIGLA" not in df_base.columns and num_colunas > 0: df_base["SIGLA"] = df_raw.iloc[:, 0]
            if "SETOR" not in df_base.columns and num_colunas > 1: df_base["SETOR"] = df_raw.iloc[:, 1]
            if "RESPONSAVEL" not in df_base.columns and num_colunas > 3: df_base["RESPONSAVEL"] = df_raw.iloc[:, 3]
            if "STATUS" not in df_base.columns and num_colunas > 4: df_base["STATUS"] = df_raw.iloc[:, 4]
            if "SIT_PRAZO" not in df_base.columns and num_colunas > 5: df_base["SIT_PRAZO"] = df_raw.iloc[:, 5]

            df_base = df_base.dropna(subset=["SIGLA", "STATUS"], how="all")

            for col in df_base.columns:
                df_base[col] = df_base[col].fillna("").astype(str).str.strip()
            
            valores_vazios = ["0", "0.0", "NAN", "NONE", "", "NAN NAN", "NÃO INFORMADO", "A", "#VALOR!"]
            for col in df_base.columns:
                df_base.loc[df_base[col].str.upper().isin(valores_vazios), col] = None

            df_base = df_base.dropna(subset=["SIGLA", "STATUS"], how="any")

            if "RESPONSAVEL" in df_base.columns:
                df_base["RESPONSAVEL"] = df_base["RESPONSAVEL"].apply(
                    lambda x: "Antigo Colaborador" if str(x).upper().strip() in ["SABRINA", "SONALHYA"] else str(x).upper().strip()
                )
            
            if "STATUS" in df_base.columns:
                df_base["STATUS"] = df_base["STATUS"].apply(
                    lambda x: "AG Aguardando" if "VERIFICADO AGU" in x.upper() or "AGUARDANDO" in x.upper() else x
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
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("📊 Qtd por Documento (Sigla)")
    df_lateral_contagem = df_base["SIGLA"].dropna().value_counts().reset_index()
    df_lateral_contagem.columns = ["Documento", "Qtd"]
    # ALTERADO: Travado a altura padrão para gerar barra de rolagem na tabela interna lateral
    st.sidebar.dataframe(df_lateral_contagem, use_container_width=True, hide_index=True, height=200)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("🎨 Configuração dos Gráficos")
    c_g1_color = st.sidebar.color_picker("G1 - Cor Status:", "#FBBF24", key="c1")
    c_g2_color = st.sidebar.color_picker("G2 - Cor Tempo:", "#38BDF8", key="c2")
    c_g4_color = st.sidebar.color_picker("G4 - Cor Prazo:", "#C084FC", key="c4")
    
    # Executa a filtragem real reativa dos dados na tela principal
    df_filtrado = df_base.copy()
    if responsavel_selecionado != "Todos":
        df_filtrado = df_filtrado[df_filtrado["RESPONSAVEL"] == responsavel_selecionado]
    if documento_selecionado != "Todos":
        df_filtrado = df_filtrado[df_filtrado["SIGLA"] == documento_selecionado]

    #--- 4. INDICADORES DO TOPO (CARDS) ---
    total_docs = len(df_filtrado)
    aprovados = len(df_filtrado[df_filtrado["STATUS"].astype(str).str.upper().str.contains("APROVADO|OK|SIM", na=False)])
    
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    with kpi_col1: st.metric(label="📄 TOTAL DOCUMENTOS", value=f"{total_docs}")
    with kpi_col2: st.metric(label="✅ APROVADOS", value=f"{aprovados}")
    with kpi_col3: st.metric(label="⏳ 1º VERF.", value=f"{media_v1:.1f} d")
    with kpi_col4: st.metric(label="⏳ 2º VERF.", value=f"{media_v2:.1f} d")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    #--- 5. RENDERIZAÇÃO DOS GRÁFICOS INICIAIS ---
    linha1_col1, linha1_col2 = st.columns(2)
    
    with linha1_col1:
        st.markdown("### 📊 Documentos por Status")
        dados_g1 = df_filtrado["STATUS"].dropna().value_counts()
        st.bar_chart(dados_g1, color=c_g1_color, horizontal=True)
    
    with linha1_col2:
        st.markdown("### ⏱️ Tempo de Análise")
        dados_g2 = pd.Series({"1º Verf.": media_v1, "2º Verf.": media_v2, "Total Estimado": (media_v1 + media_v2)})
        st.bar_chart(dados_g2, color=c_g2_color, horizontal=True)
    
    st.markdown("---")
    
    linha2_col1, linha2_col2 = st.columns(2)
    
