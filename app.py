import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="PAINEL DE INDICADORES NORMA ZERO", page_icon="📋", layout="wide")

# --- BLOCO DE CUSTOMIZAÇÃO VISUAL AVANÇADA (CSS INJECT NATIVO) ---
st.markdown("""
<style>
[data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th { font-size: 16px !important; font-weight: 600 !important; color: #FFFFFF !important; }
[data-testid="stDataFrame"] td:last-child { color: #FBBF24 !important; font-size: 18px !important; font-weight: bold !important; }
div[data-testid="stSelectbox"] label p, div[data-testid="stFileUploader"] label p { font-size: 15px !important; font-weight: bold !important; color: #38BDF8 !important; }
section[data-testid="stFileUploaderDropzone"] { border: 2px dashed #38BDF8 !important; background-color: #1E293B !important; }
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
            
            col_g, col_h = None, None
            for c in df_orig_cols.columns:
                if any(term in c for term in ["1º", "1O", "V 1", "I.A.V.1", "V1"]): col_g = c
                if any(term in c for term in ["2º", "2O", "V 2", "I.A.V.2", "V2"]): col_h = c
            
            if col_g:
                s_g = df_orig_cols[col_g].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False)
                s_g = s_g.replace(to_replace=[r'#.*', r'VALOR.*'], value='NaN', regex=True)
                media_v1 = pd.to_numeric(s_g, errors='coerce').dropna().mean()
            if col_h:
                s_h = df_orig_cols[col_h].astype(str).str.replace(" dias", "", regex=False).str.replace(",", ".", regex=False)
                s_h = s_h.replace(to_replace=[r'#.*', r'VALOR.*'], value='NaN', regex=True)
                media_v2 = pd.to_numeric(s_h, errors='coerce').dropna().mean()
            
            media_v1 = float(media_v1) if pd.notna(media_v1) else 0.0
            media_v2 = float(media_v2) if pd.notna(media_v2) else 0.0
        
        # 2. CARREGAMENTO DOS DADOS PARA OS GRÁFICOS
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
            
            num_colunas = len(df_raw.columns)
            df_base = pd.DataFrame()
            
            df_base["SIGLA"] = df_raw.iloc[:, 0] if num_colunas > 0 else None
            df_base["SETOR"] = df_raw.iloc[:, 1] if num_colunas > 1 else None
            df_base["RESPONSAVEL"] = df_raw.iloc[:, 3] if num_colunas > 3 else None
            df_base["STATUS"] = df_raw.iloc[:, 4] if num_colunas > 4 else None
            df_base["SIT_PRAZO"] = df_raw.iloc[:, 5] if num_colunas > 5 else None

            # CRÍTICO: Descarta linhas completamente vazias vindas do Excel para eliminar o "0" e os 598 docs fantasmas
            df_base = df_base.dropna(subset=["SIGLA", "STATUS"], how="all")

            for col in df_base.columns:
                df_base[col] = df_base[col].fillna("").astype(str).str.strip()
            
            # Limpeza e remoção do "0" textual dos dados reais
            valores_vazios = ["0", "0.0", "NAN", "NONE", "", "NAN NAN", "NÃO INFORMADO", "A", "#VALOR!"]
            for col in df_base.columns:
                df_base.loc[df_base[col].str.upper().isin(valores_vazios), col] = None

            # Segunda filtragem de segurança pós-limpeza de texto
            df_base = df_base.dropna(subset=["SIGLA", "STATUS"], how="any")

            # Substituição das colaboradoras antigas
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
    
    lista_responsaveis = sorted([str(r) for r in df_base["RESPONSAVEL"].dropna().unique() if str(r).strip() not in ["", "None"]])
    lista_responsaveis.insert(0, "Todos")
    responsavel_selecionado = st.sidebar.selectbox("Selecione o Responsável:", lista_responsaveis)
    
    lista_documentos = sorted([str(d) for d in df_base["SIGLA"].dropna().unique() if str(d).strip() not in ["", "None"]])
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
    st.sidebar.dataframe(df_lateral_contagem, use_container_width=True, hide_index=True)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("🎨 Configuração dos Gráficos")
    c_g1_color = st.sidebar.color_picker("G1 - Cor Status:", "#FBBF24", key="c1")
    c_g2_color = st.sidebar.color_picker("G2 - Cor Tempo:", "#38BDF8", key="c2")
    c_g4_color = st.sidebar.color_picker("G4 - Cor Prazo:", "#C084FC", key="c4")
    
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
    
    with linha2_col1:
        st.markdown("### 📅 Situação de Prazos")
        contagem_prazos = df_filtrado["SIT_PRAZO"].dropna().value_counts()
        dados_prazo = pd.Series({
            "No Prazo": contagem_prazos.get("No Prazo", 0) if "No Prazo" in contagem_prazos else contagem_prazos.get("Válido", 0),
            "Prestes a Vencer": contagem_prazos.get("Prestes a Vencer", 0),
            "Vencido": contagem_prazos.get("Vencido", 0)
        })
        st.bar_chart(dados_prazo, color=c_g4_color, horizontal=True)

    with linha2_col2:
        st.markdown("### 🏆 Documentos Aprovados por Tipo")
