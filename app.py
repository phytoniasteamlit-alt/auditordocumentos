import streamlit as st
import pandas as pd
import openpyxl
import unicodedata
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="PAINEL DE INDICADORES-NORMA ZERO", page_icon="📝", layout="wide")

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
    st.markdown("# PAINEL DE INDICADORES NORMATIVOS NAQH")
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
arquivo_excel = st.sidebar.file_uploader(" Carregar Planilha Excel (.xlsx):", type=["xlsx"], key="uploader_xlsx")

df_base = pd.DataFrame()
media_v1, media_v2, media_aaa = 0.0, 0.0, 0.0

def remover_acentos(texto):
    if pd.isna(texto): return ""
    return "".join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn').upper().strip()

if arquivo_excel:
    try:
        xl = pd.ExcelFile(arquivo_excel, engine="openpyxl")
        lista_abas_reais = xl.sheet_names
        
        # 1. PROCESSAMENTO DAS MÉDIAS CRONOLÓGICAS
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
                if any(term in c for term in ["I.A.A.A", "AAA", "3º", "3O"]): col_i = c
            
            if col_g:
                media_v1 = pd.to_numeric(df_orig_cols[col_g].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False), errors='coerce').dropna().mean()
            if col_h:
                media_v2 = pd.to_numeric(df_orig_cols[col_h].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False), errors='coerce').dropna().mean()
            if col_i:
                media_aaa = pd.to_numeric(df_orig_cols[col_i].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False), errors='coerce').dropna().mean()
            
            media_v1 = float(media_v1) if pd.notna(media_v1) else 0.0
            media_v2 = float(media_v2) if pd.notna(media_v2) else 0.0
            media_aaa = float(media_aaa) if pd.notna(media_aaa) else 0.0
        
        # 2. CARREGAMENTO DOS GRÁFICOS (Aba DADOS_GRÁFICOS)
        nome_aba_graficos = None
        for n_real in lista_abas_reais:
            n_up = str(n_real).upper().strip()
            if "GRAFIC" in n_up or "DADOS" in n_up:
                nome_aba_graficos = n_real
                break
        
        if not nome_aba_graficos:
            nome_aba_graficos = lista_abas_reais if lista_abas_reais else None
            
        if nome_aba_graficos:
            df_raw = pd.read_excel(arquivo_excel, sheet_name=nome_aba_graficos, header=0, engine="openpyxl")
            df_raw.columns = df_raw.columns.astype(str).str.strip().str.upper()
            
            num_colunas = len(df_raw.columns)
            df_base = pd.DataFrame()
            df_base["SIGLA"] = df_raw.iloc[:, 0] if num_colunas > 0 else None
            df_base["SETOR"] = df_raw.iloc[:, 1] if num_colunas > 1 else None
            df_base["RESPONSAVEL"] = df_raw.iloc[:, 3] if num_colunas > 3 else "Não Informado"
            df_base["STATUS"] = df_raw.iloc[:, 4] if num_colunas > 4 else "Não Informado"
            df_base["SIT_PRAZO"] = df_raw.iloc[:, 5] if num_colunas > 5 else "Não Informado"
            
            for col in df_base.columns:
                df_base[col] = df_base[col].fillna("").astype(str).str.strip()
            
            if "STATUS" in df_base.columns:
                df_base["STATUS"] = df_base["STATUS"].apply(
                    lambda x: "AG Aguardando" if "VERIFICADO AGU" in x.upper() or "AGUARDANDO" in x.upper() else x
                )
            
            valores_vazios = ["0", "0.0", "NAN", "NONE", "", "NAN NAN", "NÃO INFORMADO", "A"]
            for col in df_base.columns:
                df_base.loc[df_base[col].str.upper().isin(valores_vazios), col] = None
            
            df_base = df_base.dropna(subset=["STATUS", "SIGLA"], how="all")
            df_base = df_base[~df_base["STATUS"].astype(str).str.contains("#", na=False)]
            df_base = df_base[~df_base["SIT_PRAZO"].astype(str).str.contains("#", na=False)]
            
    except Exception as e:
        st.error(f" Erro crítico no processamento dos dados: {e}")

#--- 3. MENUS LATERAIS DE CUSTOMIZAÇÃO E ANÁLISE ---
if not df_base.empty:
    st.sidebar.markdown("---")
    st.sidebar.subheader(" Filtros de Visualização")
    
    responsaveis_validos = df_base["RESPONSAVEL"].dropna().unique()
    lista_responsaveis = sorted([str(r) for r in responsaveis_validos])
    lista_responsaveis.insert(0, "Todos")
    responsavel_selecionado = st.sidebar.selectbox("Selecione o Responsável:", lista_responsaveis)
    
    df_filtrado = df_base.copy()
    if responsavel_selecionado != "Todos":
        df_filtrado = df_base[df_base["RESPONSAVEL"] == responsavel_selecionado]
    
    st.sidebar.markdown("---")
    st.sidebar.subheader(" Qtd por Documento (Sigla)")
    df_lateral_contagem = df_filtrado["SIGLA"].dropna().value_counts().reset_index()
    df_lateral_contagem.columns = ["Documento", "Qtd"]
    st.sidebar.dataframe(df_lateral_contagem, use_container_width=True, hide_index=True)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader(" Configuração dos Gráficos")
    c_g1_color = st.sidebar.color_picker("G1 - Cor Status / Tipo:", "#FBBF24", key="c1")
    c_g4_color = st.sidebar.color_picker("G4 - Cor Prazo:", "#C084FC", key="c4")
    c_prof_color = st.sidebar.color_picker("G5 - Cor Profissional:", "#38BDF8", key="c5")

    #--- 4. INDICADORES DO TOPO (CARDS KPIs) ---
    total_docs = len(df_filtrado)
    aprovados = len(df_filtrado[df_filtrado["STATUS"].astype(str).str.upper().str.contains("APROVADO|OK|SIM", na=False)])
    
    kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)
    with kpi_col1: st.metric(label=" TOTAL DOCUMENTOS", value=f"{total_docs}")
    with kpi_col2: st.metric(label=" APROVADOS", value=f"{aprovados}")
    with kpi_col3: st.metric(label=" 1º VERF.", value=f"{media_v1:.1f} d")
    with kpi_col4: st.metric(label=" 2º VERF.", value=f"{media_v2:.1f} d")
    with kpi_col5: st.metric(label=" 3º VERF.", value=f"{media_aaa:.1f} d")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader(" Painel Executivo NAQH — Resultados Finais")
    
    #--- 5. RENDERIZAÇÃO REAL DOS GRÁFICOS ---
    
    # LINHA 1: Status Geral e Prazos
    linha1_col1, linha1_col2 = st.columns(2)
    with linha1_col1:
        st.markdown("### Documentos por Status")
        dados_g1 = df_filtrado["STATUS"].dropna().value_counts()
        dados_g1 = dados_g1[~dados_g1.index.astype(str).str.upper().str.contains("CANCELADO", na=False)]
        st.bar_chart(dados_g1, color=c_g1_color, horizontal=True)
    
    with linha1_col2:
        st.markdown("### Situação de Prazos")
        s_prazos_limpos = df_filtrado["SIT_PRAZO"].apply(remover_acentos)
        contagem_prazos = s_prazos_limpos.value_counts()
        
        dados_prazo = pd.Series(0, index=["No Prazo", "Prestes a Vencer", "Vencido"])
        dados_prazo["No Prazo"] = int(contagem_prazos.get("VALIDO", 0) + contagem_prazos.get("NO PRAZO", 0))
        dados_prazo["Prestes a Vencer"] = int(contagem_prazos.get("PRESTES A VENCER", 0))
        dados_prazo["Vencido"] = int(contagem_prazos.get("VENCIDO", 0))
        st.bar_chart(dados_prazo, color=c_g4_color, horizontal=True)
    
    # LINHA DIVISÓRIA INVISÍVEL CORRIGIDA (Comandos nativos sequenciais de espaçamento limpo)
    st.write("")
    st.write("")
    st.write("")
    st.write("")
    
