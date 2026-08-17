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
def extrair_dados_aba_especifica(uploaded_file, numero_posicao_aba):
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=numero_posicao_aba, header=None, skiprows=7)
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

# ==============================================================================
# 4. CRIAÇÃO DAS ABAS DE NAVEGAÇÃO ISOLADAS (HCID vs ANEXOS)
# ==============================================================================
if uploaded_file is not None:
    excel_file = pd.ExcelFile(uploaded_file)
    abas_planilha = excel_file.sheet_names
    
    df_hcid = extrair_dados_aba_especifica(uploaded_file, 0)
    df_anexos = extrair_dados_aba_especifica(uploaded_file, 1) if len(abas_planilha) > 1 else pd.DataFrame()

    tab_hcid, tab_anexos = st.tabs(["🏥 Hospital Geral (HCID)", "🏢 Unidades Anexas"])

    # --------------------------------==========================================
    # CONTEÚDO EXCLUSIVO DA ABA 1: HCID
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
            c3.metric("☀️ Total de vagas do turno manhã", f"{t_m_h} M")
            c4.metric("🌙 Total de vagas do turno tarde", f"{t_t_h} T")
            
            st.markdown("<br>", unsafe_allow_html=True)
            col1_h, col2_h = st.columns(2)
            
            with col1_h:
                st.markdown("##### 1️⃣ Total de vagas e de estágio no HCID")
                df_g1_h = df_hcid.groupby("SETOR_RAW")["TOTAL_VAGAS"].sum().reset_index()
                f1_h = px.bar(df_g1_h, x="SETOR_RAW" if is_vert else "TOTAL_VAGAS", y="TOTAL_VAGAS" if is_vert else "SETOR_RAW", orientation="v" if is_vert else "h", text_auto=True, color_discrete_sequence=seq_cores)
                f1_h.update_layout(height=320, margin=dict(l=10,r=15,t=10,b=10), xaxis_title="Setor" if is_vert else "Vagas", yaxis_title="Vagas" if is_vert else "Setor")
                f1_h.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(f1_h, use_container_width=True)
                
                st.markdown("##### 3️⃣ Setores disponibilizados para realização de estágio no HCID")
                df_g3_h = df_hcid.groupby("SETOR_RAW")["TOTAL_VAGAS"].sum().reset_index().sort_values(by="TOTAL_VAGAS", ascending=True)
                f3_h = px.bar(df_g3_h, x="SETOR_RAW" if is_vert else "TOTAL_VAGAS", y="TOTAL_VAGAS" if is_vert else "SETOR_RAW", orientation="v" if is_vert else "h", text_auto=True, color="TOTAL_VAGAS", color_continuous_scale=px.colors.sequential.Tealgrn)
                f3_h.update_layout(height=320, coloraxis_showscale=False, margin=dict(l=10,r=15,t=10,b=10), xaxis_title="Setor" if is_vert else "Vagas", yaxis_title="Vagas" if is_vert else "Setor")
                f3_h.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(f3_h, use_container_width=True)

                st.markdown("##### 5️⃣ Total de vagas de estágio disponibilizados por setor no HCID")
                f5_h = px.bar(df_g3_h, x="SETOR_RAW" if is_vert else "TOTAL_VAGAS", y="TOTAL_VAGAS" if is_vert else "SETOR_RAW", orientation="v" if is_vert else "h", text_auto=True, color_discrete_sequence=seq_cores)
                f5_h.update_layout(height=320, margin=dict(l=10,r=15,t=10,b=10), xaxis_title="Setor" if is_vert else "Vagas", yaxis_title="Vagas" if is_vert else "Setor")
                f5_h.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(f5_h, use_container_width=True)

                st.markdown("##### 7️⃣ Total de estagiários por turno por dia no HCID")
                df_melt_h = df_hcid.groupby("SETOR_RAW")[["MANHÃ", "TARDE"]].sum().reset_index().melt(id_vars="SETOR_RAW", var_name="TURNO", value_name="VAGAS")
                f7_h = px.bar(df_melt_h, x="SETOR_RAW" if is_vert else "VAGAS", y="VAGAS" if is_vert else "SETOR_RAW", color="TURNO", barmode="group", orientation="v" if is_vert else "h", text_auto=True, color_discrete_map={"MANHÃ": "#008080", "TARDE": "#FF7F50"})
                f7_h.update_layout(height=320, margin=dict(l=10,r=15,t=10,b=10), xaxis_title="Setor" if is_vert else "Vagas", yaxis_title="Vagas" if is_vert else "Setor")
                f7_h.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(f7_h, use_container_width=True)

            with col2_h:
                # CORREÇÃO CRÍTICA DO GRÁFICO 2: Corrigido o DataFrame de plotagem para reaparecer de forma estática sem erros
                st.markdown("##### 2️⃣ Total de setores disponibilizados p/ campo de estágio no HCID")
                df_g2_h = pd.DataFrame([{"Mapeamento": "Setores Ativos", "Quantidade": t_setores_h}])
                f2_h = px.bar(df_g2_h, x="Mapeamento" if is_vert else "Quantidade", y="Quantidade" if is_vert else "Mapeamento", orientation="v" if is_vert else "h", text_auto=True, color_discrete_sequence=["#2E8B57"])
                f2_h.update_layout(height=320, margin=dict(l=10,r=15,t=10,b=10), xaxis_title="Mapeamento" if is_vert else "Quantidade", yaxis_title="Quantidade" if is_vert else "Mapeamento")
