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
    # Remove textos como "por turno" e captura exclusivamente os dígitos numéricos
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
# 2. PAINEL DE CONTROLE (SIDEBAR)
# ==============================================================================
st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha de Estágios (.xlsx):", type=["xlsx"])

# ==============================================================================
# 3. MOTOR DE PROCESSAMENTO LINEAR FOCADO NAS COLUNAS FIXAS D E E
# ==============================================================================
def processar_vagas_estaticas(uploaded_file, padrao_procurado, sheet_fallback):
    excel_file = pd.ExcelFile(uploaded_file)
    abas = excel_file.sheet_names
    
    aba_real = next((op for op in padrao_procurado if op in abas), None)
    if not aba_real:
        aba_real = sheet_fallback if sheet_fallback in abas else None
        
    if not aba_real:
        return pd.DataFrame()
        
    # Lê a tabela pulando as linhas iniciais de título (Começa estritamente na linha indexada 7)
    df_raw = pd.read_excel(uploaded_file, sheet_name=aba_real, header=None, skiprows=7)
    if df_raw.empty:
        return pd.DataFrame()
        
    # Mapeamento estrito por posições de colunas físicas (A=0, B=1, C=2, D=3, E=4)
    df_processado = pd.DataFrame()
    df_processado["SETOR_RAW"] = df_raw.iloc[:, 0].astype(str).str.strip().ffill()
    df_processado["SUB_SETOR"] = df_raw.iloc[:, 1].fillna("").astype(str).str.strip()
    df_processado["CATEGORIA"] = df_raw.iloc[:, 2].fillna("").astype(str).str.strip()
    
    # Extração direta das colunas fixas D e E de vagas por turno
    df_processado["MANHÃ"] = df_raw.iloc[:, 3].apply(extrair_numero)
    df_processado["TARDE"] = df_raw.iloc[:, 4].apply(extrair_numero)
    df_processado["TOTAL_VAGAS"] = df_processado["MANHÃ"] + df_processado["TARDE"]
    
    # Filtro rígido para descartar ruídos e linhas de totalizadores nativos da planilha
    linhas_validas = []
    for _, row in df_processado.iterrows():
        txt_s = str(row["SETOR_RAW"]).upper()
        txt_c = str(row["CATEGORIA"]).upper()
        
        # Permite strings padrões como "GERAL", limpando apenas lixos e campos de TOTAL vazios
        if "TOTAL" in txt_s or "TOTAL" in txt_c or txt_s == "NAN" or (txt_s == "" and txt_c == ""):
            linhas_validas.append(False)
        else:
            linhas_validas.append(True)
            
    df_final = df_processado[linhas_validas].copy()
    
    # Padroniza visualmente os sub-setores vazios sem apagar o registro
    df_final["SUB_SETOR"] = df_final["SUB_SETOR"].apply(lambda x: "GERAL" if x == "" else x)
    
    return df_final[df_final["TOTAL_VAGAS"] > 0]

# ==============================================================================
# 4. EXECUÇÃO DO PROCESSAMENTO EM QUADROS TOTALMENTE ISOLADOS
# ==============================================================================
if uploaded_file is not None:
    # Processamento isolado do HCID usando a aba limpa configurada
    df_hcid = processar_vagas_estaticas(uploaded_file, ["HCID_BDD", "HCID", "HCID1"], "HCID_BDD")
    
    # Processamento isolado dos Anexos
    df_anexos = processar_vagas_estaticas(uploaded_file, ["ANEXO", "ANEXO2", "ANEXOS"], "ANEXO")

    # ==========================================================================
    # QUADRO 1: CONJUNTO EXCLUSIVO HOSPITAL GERAL (HCID)
    # ==========================================================================
    st.markdown("<div style='background-color: #1a2a3a; padding: 12px; border-radius: 5px; margin-bottom: 20px;'><h2 style='margin:0; font-size:1.6rem; color:#fff;'>🏥 QUADRO DE INDICADORES - SOMENTE HCID</h2></div>", unsafe_allow_html=True)
    
    if not df_hcid.empty:
        # Caixas de Texto com os valores cravados do Excel (158 Vagas / 21 Setores)
        t_vagas_h = df_hcid["TOTAL_VAGAS"].sum()
        t_setores_h = df_hcid["SETOR_RAW"].nunique()
        t_m_h = df_hcid["MANHÃ"].sum()
        t_t_h = df_hcid["TARDE"].sum()
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total de vagas de estágio geral HCID", f"{t_vagas_h} Vagas")
        c2.metric("Total de setores disponibilizados no HCID", f"{t_setores_h} Setores")
        c3.metric("Total de vagas de estágio do HCID por turno manhã", f"{t_m_h} M")
        c4.metric("Total de vagas de estágio do HCID tarde", f"{t_t_h} T")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        col1_h, col2_h = st.columns(2)
        
        with col1_h:
            st.markdown("##### 1️⃣ Total de vagas e de estágio no HCID")
            df_g1_h = df_hcid.groupby("SETOR_RAW")["TOTAL_VAGAS"].sum().reset_index()
            f1_h = px.bar(df_g1_h, x="SETOR_RAW", y="TOTAL_VAGAS", text_auto=True, color_discrete_sequence=["#4682B4"])
            f1_h.update_layout(height=280, margin=dict(l=10,r=10,t=10,b=10), xaxis_title="Setor", yaxis_title="Vagas")
            st.plotly_chart(f1_h, use_container_width=True)
            
            st.markdown("##### 3️⃣ Setores disponibilizados para realização de estágio no HCID")
            df_g3_h = df_hcid.groupby("SETOR_RAW")["TOTAL_VAGAS"].sum().reset_index().sort_values(by="TOTAL_VAGAS", ascending=True)
            f3_h = px.bar(df_g3_h, x="TOTAL_VAGAS", y="SETOR_RAW", orientation="h", text_auto=True, color="TOTAL_VAGAS", color_continuous_scale=px.colors.sequential.Tealgrn)
            f3_h.update_layout(height=280, coloraxis_showscale=False, margin=dict(l=10,r=10,t=10,b=10), xaxis_title="Vagas", yaxis_title="Setor")
            st.plotly_chart(f3_h, use_container_width=True)

            # CORREÇÃO CRÍTICA DO GRÁFICO 5: Vincula o DataFrame correto ordenado de forma crescente para as barras variarem
            st.markdown("##### 5️⃣ Total de vagas de estágio disponibilizados por setor no HCID")
            f5_h = px.bar(df_g3_h, x="TOTAL_VAGAS", y="SETOR_RAW", orientation="h", text_auto=True, color_discrete_sequence=["#5F9EA0"])
            f5_h.update_layout(height=280, margin=dict(l=10,r=10,t=10,b=10), xaxis_title="Vagas", yaxis_title="Setor")
            st.plotly_chart(f5_h, use_container_width=True)

            st.markdown("##### 7️⃣ Total de estagiários por turno por dia no HCID")
            df_melt_h = df_hcid.groupby("SETOR_RAW")[["MANHÃ", "TARDE"]].sum().reset_index().melt(id_vars="SETOR_RAW", var_name="TURNO", value_name="VAGAS")
            f7_h = px.bar(df_melt_h, x="SETOR_RAW", y="VAGAS", color="TURNO", barmode="group", text_auto=True, color_discrete_map={"MANHÃ": "#008080", "TARDE": "#FF7F50"})
            f7_h.update_layout(height=280, margin=dict(l=10,r=10,t=10,b=10), xaxis_title="Setor", yaxis_title="Vagas")
            st.plotly_chart(f7_h, use_container_width=True)

        with col2_h:
            st.markdown("##### 2️⃣ Total de setores disponibilizados p/ campo de estágio no HCID")
            df_g2_h = pd.DataFrame([{"Mapeamento": "Setores Ativos", "Quantidade": t_setores_h}])
            f2_h = px.bar(df_g2_h, x="Mapeamento", y="Quantidade", text_auto=True, color_discrete_sequence=["#2E8B57"])
            f2_h.update_layout(height=280, margin=dict(l=10,r=10,t=10,b=10), yaxis_title="Quantidade")
            st.plotly_chart(f2_h, use_container_width=True)

            st.markdown("##### 4️⃣ Categorias profissionais contemplados no estágio por setor no HCID")
            sel_s_h = st.selectbox("Escolha o Setor do HCID para Filtrar:", sorted(df_hcid["SETOR_RAW"].unique()), key="sel_g4_h")
            df_g4_h = df_hcid[df_hcid["SETOR_RAW"] == sel_s_h].groupby("CATEGORIA")["TOTAL_VAGAS"].sum().reset_index().sort_values(by="TOTAL_VAGAS", ascending=True)
            f4_h = px.bar(df_g4_h, x="TOTAL_VAGAS", y="CATEGORIA", orientation="h", text_auto=True, color_discrete_sequence=["#20B2AA"])
            f4_h.update_layout(height=215, margin=dict(l=10,r=10,t=10,b=10), xaxis_title="Vagas", yaxis_title="Profissão")
            st.plotly_chart(f4_h, use_container_width=True)

            st.markdown("##### 6️⃣ Total de vagas de estágio do HCID por turno no HCID")
            df_g6_h = pd.DataFrame([{"TURNO": "MANHÃ", "VAGAS": t_m_h}, {"TURNO": "TARDE", "VAGAS": t_t_h}])
            f6_h = px.pie(df_g6_h, values="VAGAS", names="TURNO", color="TURNO", color_discrete_map={"MANHÃ": "#008080", "TARDE": "#FF7F50"}, hole=0.4)
            f6_h.update_layout(height=280, margin=dict(l=10,r=10,t=10,b=10))
            st.plotly_chart(f6_h, use_container_width=True)
    else:
