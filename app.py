import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="PAINEL DE INDICADORES-NORMA ZERO", page_icon="📊", layout="wide")

# --- BLOCO DE CUSTOMIZAÇÃO VISUAL AVANÇADA (CSS INJECT NATIVO) ---
st.markdown("""
    <style>
    /* 1. Aumenta e destaca os números e textos de TODAS as tabelas/dataframes */
    [data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th {
        font-size: 16px !important;
        font-weight: 600 !important;
        color: #FFFFFF !important;
    }
    
    /* 2. Destaca a coluna de contagem/números com uma cor amarela viva para leitura rápida */
    [data-testid="stDataFrame"] td:last-child {
        color: #FBBF24 !important;
        font-size: 18px !important;
        font-weight: bold !important;
    }

    /* 3. Aumenta o tamanho dos textos dos botões e caixas de seleção da lateral */
    div[data-testid="stSelectbox"] label p, div[data-testid="stFileUploader"] label p {
        font-size: 15px !important;
        font-weight: bold !important;
        color: #38BDF8 !important;
    }

    /* 4. Altera o visual do botão de carregar arquivo para ficar bem visível */
    section[data-testid="stFileUploaderDropzone"] {
        border: 2px dashed #38BDF8 !important;
        background-color: #1E293B !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("# DASHBOARD EXECUTIVO DE INDICADORES — NAQH")
st.markdown("---")

#--- 2. CARREGAMENTO DO ARQUIVO ---
st.sidebar.header("⚙️ Entrada de Dados")
arquivo_excel = st.sidebar.file_uploader("📂 Carregar Planilha Excel (.xlsx):", type=["xlsx"], key="uploader_xlsx")

df_base = pd.DataFrame()
media_v1, media_v2, media_aaa = 0.0, 0.0, 0.0

if arquivo_excel:
    try:
        xl = pd.ExcelFile(arquivo_excel, engine="openpyxl")
        lista_abas_reais = xl.sheet_names
        
        # 1. PROCESSAMENTO DAS MÉDIAS CRONOLÓGICAS (Aba de Acompanhamento)
        nome_aba_original = None
        for n_real in lista_abas_reais:
            n_up = str(n_real).upper().strip()
            if "VERF" in n_up or "ACOMP" in n_up:
                nome_aba_original = n_real
                break
                
        if nome_aba_original:
            df_orig_cols = pd.read_excel(arquivo_excel, sheet_name=nome_aba_original, header=2, engine="openpyxl")
            df_orig_cols.columns = df_orig_cols.columns.astype(str).str.strip().str.upper()
            
            col_g, col_h, col_i = None, None, None
            for c in df_orig_cols.columns:
                if any(term in c for term in ["1º", "1O", "V 1", "I.A.V.1", "V1"]): col_g = c
                if any(term in c for term in ["2º", "2O", "V 2", "I.A.V.2", "V2"]): col_h = c
                if any(term in c for term in ["I.A.A.A", "AAA"]): col_i = c
                    
            if col_g:
                media_v1 = pd.to_numeric(df_orig_cols[col_g].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False), errors='coerce').dropna().mean()
            if col_h:
                media_v2 = pd.to_numeric(df_orig_cols[col_h].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False), errors='coerce').dropna().mean()
            if col_i:
                media_aaa = pd.to_numeric(df_orig_cols[col_i].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False), errors='coerce').dropna().mean()
                
            media_v1 = float(media_v1) if pd.notna(media_v1) else 0.0
            media_v2 = float(media_v2) if pd.notna(media_v2) else 0.0
            media_aaa = float(media_aaa) if pd.notna(media_aaa) else 0.0

        # 2. CARREGAMENTO DOS GRÁFICOS
        nome_aba_graficos = None
        for n_real in lista_abas_reais:
            n_up = str(n_real).upper().strip()
            if "GRAFIC" in n_up or "DADOS" in n_up:
                nome_aba_graficos = n_real
                break
                
        if not nome_aba_graficos:
            nome_aba_graficos = lista_abas_reais

        # Carregamento seguro direto do Pandas para evitar erro de conflito de colunas (599 colunas)
        df_raw = pd.read_excel(arquivo_excel, sheet_name=nome_aba_graficos, header=0, engine="openpyxl")
        df_raw.columns = df_raw.columns.astype(str).str.strip().str.upper()
        
        df_base = pd.DataFrame()
        df_base["SIGLA"] = df_raw.iloc[:, 0]
        df_base["SETOR"] = df_raw.iloc[:, 1]
        df_base["RESPONSAVEL"] = df_raw.iloc[:, 3]
        df_base["STATUS"] = df_raw.iloc[:, 4]
        df_base["SIT_PRAZO"] = df_raw.iloc[:, 5]

        for col in ["SIGLA", "STATUS", "SIT_PRAZO", "SETOR", "RESPONSAVEL"]:
            df_base[col] = df_base[col].astype(str).str.strip().replace(["0", "0.0", "nan", "None", "NAN", "", "nan nan"], "Não Informado")
            
        df_base = df_base[~df_base["SIT_PRAZO"].str.contains("#", na=False)]
        df_base = df_base[~df_base["STATUS"].str.contains("#", na=False)]
        df_base = df_base[(df_base["SIGLA"] != "Não Informado") | (df_base["RESPONSAVEL"] != "Não Informado")]
        df_base["SIT_PRAZO"] = df_base["SIT_PRAZO"].replace(["A", "a"], "Prestes a Vencer")

    except Exception as e:
        st.error(f"❌ Erro crítico no processamento dos dados: {e}")

#--- 3. MENUS LATERAIS DE CUSTOMIZAÇÃO E ANÁLISE ---
if not df_base.empty:
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtros de Visualização")
    lista_responsaveis = sorted([str(r).strip() for r in df_base["RESPONSAVEL"].unique() if str(r).strip() != "Não Informado"])
    lista_responsaveis.insert(0, "Todos")
    responsavel_selecionado = st.sidebar.selectbox("Selecione o Responsável:", lista_responsaveis)
    
    df_filtrado = df_base.copy()
    if responsavel_selecionado != "Todos":
        df_filtrado = df_base[df_base["RESPONSAVEL"] == responsavel_selecionado]

    # --- TABELA DE QUANTIDADE POR SIGLA NA BARRA LATERAL ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("📋 Qtd por Documento (Sigla)")
    df_lateral_contagem = df_filtrado["SIGLA"].value_counts().reset_index()
    df_lateral_contagem.columns = ["Documento", "Qtd"]
    st.sidebar.dataframe(df_lateral_contagem, use_container_width=True, hide_index=True)

    st.sidebar.markdown("---")
    st.sidebar.subheader("🎨 Configuração dos Gráficos")
    
    st.sidebar.markdown("**1. Gráfico de Situação (Prazos)**")
    t_g1 = st.sidebar.selectbox("Visualização:", ["Horizontal", "Vertical", "Linha"], key="t1")
    c_g1_color = st.sidebar.color_picker("Cor do Indicador:", "#FBBF24", key="c1")
    
    st.sidebar.markdown("**2. Gráfico de Status Normativo**")
    t_g2 = st.sidebar.selectbox("Visualização:", ["Horizontal", "Vertical", "Linha"], key="t2")
    c_g2_color = st.sidebar.color_picker("Cor do Indicador:", "#38BDF8", key="c2")
    
    st.sidebar.markdown("**3. Gráfico de Setores**")
    t_g3 = st.sidebar.selectbox("Visualização:", ["Horizontal", "Vertical", "Linha"], key="t3")
    c_g3_color = st.sidebar.color_picker("Cor do Indicador:", "#34D399", key="c3")
    
    st.sidebar.markdown("**4. Detalhamento por Documento**")
    t_g4 = st.sidebar.selectbox("Visualização:", ["Horizontal", "Vertical", "Linha"], key="t4")
    c_g4_color = st.sidebar.color_picker("Cor do Indicador:", "#C084FC", key="c4")
    
    st.sidebar.markdown("**5. Produtividade por Colaborador**")
    t_g5 = st.sidebar.selectbox("Visualização:", ["Horizontal", "Vertical", "Linha"], key="t5")
    c_g5_color = st.sidebar.color_picker("Cor do Indicador:", "#FB923C", key="c5")

    #--- 4. INDICADORES DO TOPO (CARDS KPIs) ---
    total_docs = len(df_filtrado)
    aprovados = len(df_filtrado[df_filtrado["STATUS"].str.upper().str.contains("APROVADO|OK|SIM", regex=True)])
    
    kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)
    with kpi_col1: st.metric(label="📊 TOTAL DOCUMENTOS", value=f"{total_docs}")
    with kpi_col2: st.metric(label="🏆 APROVADOS", value=f"{aprovados}")
    with kpi_col3: st.metric(label="⏳ MÉDIA I.A.V 1º", value=f"{media_v1:.1f} d")
    with kpi_col4: st.metric(label="⏳ MÉDIA I.A.V 2º", value=f"{media_v2:.1f} d")
    with kpi_col5: st.metric(label="⏳ MÉDIA I.A.A.A", value=f"{media_aaa:.1f} d")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📊 Painel Executivo NAQH — Resultados Finais")
    
    #--- 5. CONTAGENS DOS DADOS ---
    dados_g1 = df_filtrado["SIT_PRAZO"].value_counts()
    dados_g2 = df_filtrado["STATUS"].value_counts()
    dados_g3 = df_filtrado["SETOR"].value_counts()
    dados_g4 = df_filtrado["SIGLA"].value_counts()
    dados_g5 = df_filtrado["RESPONSAVEL"].value_counts()

    #--- FUNÇÃO INTERNA SEGURA DE PLOTAGEM NATIVA ---
    def renderizar_grafico_seguro(dados, tipo, cor):
        if tipo == "Horizontal":
            st.bar_chart(dados, color=cor, horizontal=True)
        elif tipo == "Vertical":
            st.bar_chart(dados, color=cor)
        else:
            st.line_chart(dados, color=cor)

    #--- 6. RENDERIZAÇÃO REAL DOS GRÁFICOS (LADO A LADO) ---
    linha1_col1, linha1_col2 = st.columns(2)
    with linha1_col1:
        st.markdown("**1. Situação de Prazos**")
        renderizar_grafico_seguro(dados_g1, t_g1, c_g1_color)
            
    with linha1_col2:
        st.markdown("**2. Status Normativo**")
        renderizar_grafico_seguro(dados_g2, t_g2, c_g2_color)

    st.markdown("---")
    
    linha2_col1, linha2_col2 = st.columns(2)
    with linha2_col1:
        st.markdown("**3. Distribuição por Setores**")
        renderizar_grafico_seguro(dados_g3, t_g3, c_g3_color)
            
    with linha2_col2:
        st.markdown("**4. Detalhamento por Documento**")
        renderizar_grafico_seguro(dados_g4, t_g4, c_g4_color)

    st.markdown("---")
    st.markdown("**5. Produtividade Geral por Colaborador**")
    renderizar_grafico_seguro(dados_g5, t_g5, c_g5_color)

    #--- 7. TABELA DETALHADA ---
    st.markdown("---")
