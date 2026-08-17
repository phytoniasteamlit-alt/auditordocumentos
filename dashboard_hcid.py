import streamlit as st
import pandas as pd
import plotly.express as px

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (Interface Dashboard Executivo)
# ==============================================================================
st.set_page_config(
    page_title="Painel de Controle de Estágios - HCID & ANEXOS",
    layout="wide",
    initial_sidebar_state="expanded"
)

def extrair_numero(valor):
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan":
        return 0
    v_str = "".join(filter(str.isdigit, str(valor)))
    return int(v_str) if v_str != "" else 0

# --- CABEÇALHO SUPERIOR INSTITUCIONAL ---
header_left, header_right = st.columns(2)
header_left.markdown("<h1 style='margin: 0; padding: 0; font-size: 2.2rem;'>📊 Painel Unificado de Indicadores de Estágio</h1>", unsafe_allow_html=True)
header_right.markdown(
    """
    <div style="text-align: right; line-height: 1.3; padding-bottom: 10px;">
        <span style="font-size: 16px; font-weight: bold;">🏥 Hospital da Cidade Dr. Jackson Lago</span><br>
        <span style="font-size: 14px; color: #ccc;">👩‍💼 Coordenação: Verônica Azevedo</span><br>
        <span style="font-size: 13px; color: #aaa; font-weight: 500;">📌 Setor Nep / Nepex</span>
    </div>
    """, 
    unsafe_allow_html=True
)
st.markdown("---")

# ==============================================================================
# 2. PAINEL DE CONTROLE E CUSTOMIZAÇÃO VISUAL (SIDEBAR)
# ==============================================================================
st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha de Estágios (.xlsx):", type=["xlsx"])

st.sidebar.markdown("---")
st.sidebar.subheader("🎨 Customização Interativa")

paleta_selecionada = st.sidebar.selectbox(
    "Tema de Cores dos Gráficos:",
    options=["Padrão Hospitalar", "Tons Pastéis", "Vibrante", "Esmeralda"],
    index=0
)

estilo_grafico = st.sidebar.radio(
    "Orientação das Barras:",
    options=["Barras Verticais", "Barras Horizontais"],
    index=1
)

if paleta_selecionada == "Tons Pastéis":
    seq_cores = px.colors.qualitative.Pastel
elif paleta_selecionada == "Vibrante":
    seq_cores = px.colors.qualitative.Prism
elif paleta_selecionada == "Esmeralda":
    seq_cores = px.colors.sequential.Mint
else:
    seq_cores = ["#4682B4", "#008080", "#20B2AA", "#5F9EA0", "#B0C4DE"]

is_vert = estilo_grafico == "Barras Verticais"

# ==============================================================================
# 3. MOTOR DE PROCESSAMENTO LINEAR ISOLADO POR ÍNDICE FÍSICO DE ABA
# ==============================================================================
def extrair_dados_hcid_estatico(uploaded_file):
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None, skiprows=7)
        if df_raw.empty:
            return pd.DataFrame()
            
        df_processado = pd.DataFrame()
        df_processado["SETOR_RAW"] = df_raw.iloc[:, 0].astype(str).str.strip().ffill()
        df_processado["SUB_SETOR"] = df_raw.iloc[:, 1].fillna("").astype(str).str.strip()
        df_processado["CATEGORIA"] = df_raw.iloc[:, 2].fillna("").astype(str).str.strip()
        
        df_processado["MANHÃ"] = df_raw.iloc[:, 3].apply(extrair_numero)
        df_processado["TARDE"] = df_raw.iloc[:, 4].apply(extrair_numero)
        df_processado["TOTAL_VAGAS"] = df_processado["MANHÃ"] + df_processado["TARDE"]
        
        linhas_validas = []
        for _, row in df_processado.iterrows():
            txt_s = str(row["SETOR_RAW"]).upper()
            txt_c = str(row["CATEGORIA"]).upper()
            if "TOTAL" in txt_s or "TOTAL" in txt_c or txt_s == "NAN" or (txt_s == "" and txt_c == ""):
                linhas_validas.append(False)
            else:
                linhas_validas.append(True)
                
        df_final = df_processado[linhas_validas].copy()
        df_final["CATEGORIA"] = df_final["CATEGORIA"].apply(lambda x: "NÃO ESPECIFICADO" if x == "" else x)
        df_final["SUB_SETOR"] = df_final["SUB_SETOR"].apply(lambda x: "GERAL" if x == "" else x)
        
        return df_final[df_final["TOTAL_VAGAS"] > 0]
    except:
        return pd.DataFrame()

def extrair_dados_anexos_calendario(uploaded_file):
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=1, header=None)
        if df_raw.empty or len(df_raw) <= 7:
            return pd.DataFrame()
            
        linha_turnos = pd.Series(df_raw.iloc[5, :]).ffill().fillna("").astype(str).tolist()
        df_corpo = df_raw.iloc[7:].copy().reset_index(drop=True)
        
        setores_col = df_corpo.iloc[:, 0].astype(str).str.strip().replace(["nan", "NAN", ""], None).ffill().fillna("GERAL")
        sub_setores_col = df_corpo.iloc[:, 1].astype(str).str.strip().replace(["nan", "NAN", ""], None).ffill().fillna("GERAL")
        categorias_col = df_corpo.iloc[:, 2].astype(str).str.strip().replace(["nan", "NAN", ""], None).ffill().fillna("NÃO ESPECIFICADO")
        
        registros_vagas = []
        num_colunas_total = len(df_raw.columns)
        
        for idx_row in range(len(df_corpo)):
            setor_a = str(setores_col.iloc[idx_row]).upper()
            cat_a = str(categorias_col.iloc[idx_row]).upper()
            
            if "TOTAL" in setor_a or "TOTAL" in cat_a or cat_a == "" or setor_a == "SETOR":
                continue
                
            for col_idx in range(8, num_colunas_total):
                vaga_bruta = df_corpo.iloc[idx_row, col_idx]
                qtd_vagas = extrair_numero(vaga_bruta)
                
                if qtd_vagas > 0:
                    turno_atual = str(linha_turnos[col_idx]).strip().upper()
                    final_turno = "MANHÃ" if "MANH" in turno_atual else "TARDE"
                    
                    registros_vagas.append({
                        "SETOR": setor_a,
                        "SUB_SETOR": str(sub_setores_col.iloc[idx_row]).upper(),
                        "CATEGORIA": cat_a,
                        "TURNO": final_turno,
                        "TOTAL_VAGAS": qtd_vagas
                    })
                    
        return pd.DataFrame(registros_vagas)
    except:
        return pd.DataFrame()

# ==============================================================================
# 4. AMBIENTE DE VISUALIZAÇÃO SEPARADO POR ABAS NAVEGACIONAIS
# ==============================================================================
if uploaded_file is not None:
    excel_file = pd.ExcelFile(uploaded_file)
    abas_planilha = excel_file.sheet_names
    
    df_hcid = extrair_dados_hcid_estatico(uploaded_file)
    df_anexos = extrair_dados_anexos_calendario(uploaded_file) if len(abas_planilha) > 1 else pd.DataFrame()

    tab_hcid, tab_anexos = st.tabs(["🏥 Hospital Geral (HCID)", "🏢 Unidades Anexas"])

    # --------------------------------==========================================
    # GUIA 1: CONTEÚDO REORGANIZADO DO HCID (SEQUÊNCIA 1 A 6)
    # --------------------------------==========================================
    with tab_hcid:
        st.markdown("<div style='background-color: #1a2a3a; padding: 12px; border-radius: 5px; margin-bottom: 20px;'><h2 style='margin:0; font-size:1.4rem; color:#fff;'>📊 QUADRO DE INDICADORES - SOMENTE HCID</h2></div>", unsafe_allow_html=True)
        
        if not df_hcid.empty:
            t_vagas_h = df_hcid["TOTAL_VAGAS"].sum()
            t_setores_h = df_hcid["SETOR_RAW"].nunique()
            t_m_h = df_hcid["MANHÃ"].sum()
            t_t_h = df_hcid["TARDE"].sum()
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("📑 Total de vagas de estágio geral HCID", f"{t_vagas_h} Vagas")
            c2.metric("🏥 Total de setores disponibilizados no HCID", f"{t_setores_h} Setores")
            c3.metric("☀️ Total de vagas de estágio do HCID por turno manhã", f"{t_m_h} M")
            c4.metric("🌙 Total de vagas de estágio do HCID tarde", f"{t_t_h} T")
            
            st.markdown("<br>", unsafe_allow_html=True)
            col1_h, col2_h = st.columns(2)
            
            with col1_h:
                st.markdown("##### 1️⃣ Total de vagas e de estágio no HCID")
                df_g1_h = df_hcid.groupby("SETOR_RAW")["TOTAL_VAGAS"].sum().reset_index()
                f1_h = px.bar(df_g1_h, x="SETOR_RAW" if is_vert else "TOTAL_VAGAS", y="TOTAL_VAGAS" if is_vert else "SETOR_RAW", orientation="v" if is_vert else "h", text_auto=True, color_discrete_sequence=seq_cores)
                f1_h.update_layout(height=320, margin=dict(l=10,r=15,t=10,b=10), xaxis_title="Setor" if is_vert else "Vagas", yaxis_title="Vagas" if is_vert else "Setor")
                f1_h.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(f1_h, use_container_width=True)
                
                st.markdown("##### 3️⃣ Total de vagas de estágio disponibilizados por setor no HCID")
                df_g3_h = df_hcid.groupby("SETOR_RAW")["TOTAL_VAGAS"].sum().reset_index().sort_values(by="TOTAL_VAGAS", ascending=True)
                f3_h = px.bar(df_g3_h, x="SETOR_RAW" if is_vert else "TOTAL_VAGAS", y="TOTAL_VAGAS" if is_vert else "SETOR_RAW", orientation="v" if is_vert else "h", text_auto=True, color_discrete_sequence=seq_cores)
                f3_h.update_layout(height=320, margin=dict(l=10,r=15,t=10,b=10), xaxis_title="Setor" if is_vert else "Vagas", yaxis_title="Vagas" if is_vert else "Setor")
                f3_h.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(f3_h, use_container_width=True)

                st.markdown("##### 5️⃣ Total de estagiários por turno por dia no HCID")
                df_melt_h = df_hcid.groupby("SETOR_RAW")[["MANHÃ", "TARDE"]].sum().reset_index().melt(id_vars="SETOR_RAW", var_name="TURNO", value_name="VAGAS")
