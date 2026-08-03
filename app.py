import streamlit as st
import pandas as pd
import plotly.express as px

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (Modo Amplo)
# ==============================================================================
st.set_page_config(
    page_title="Painel de Indicadores Norma Zero",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CABEÇALHO SUPERIOR (Alinhado na mesma linha) ---
header_left, header_right = st.columns(2)

with header_left:
    st.markdown(
        "<h1 style='margin: 0; padding: 0; font-size: 2.2rem;'>📊 Painel de Indicadores Norma Zero</h1>", 
        unsafe_allow_html=True
    )

with header_right:
    st.markdown(
        """
        <div style="text-align: right; line-height: 1.2; padding-bottom: 10px;">
            <span style="font-size: 16px; font-weight: bold;">🏥 Hospital da Cidade</span><br>
            <span style="font-size: 14px; color: #888;">👩‍💼 Coord: Fabrícia Rocha</span>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

# ==============================================================================
# 2. PAINEL DE CONTROLE (SIDEBAR) & TRATAMENTO DE DADOS
# ==============================================================================
st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader(
    "Carregar Planilha Excel (.xlsx):",
    type=["xlsx"]
)

# --- PAINEL DE CUSTOMIZAÇÃO VISUAL (SIDEBAR) ---
st.sidebar.markdown("---")
st.sidebar.subheader("🎨 Customização Visual")
paleta_selecionada = st.sidebar.selectbox(
    "Tema de Cores Geral (Gráficos 1, 2 e 5):",
    options=["Padrão Hospitalar", "Tons Pastéis", "Vibrante", "Esmeralda"],
    index=2
)

if paleta_selecionada == "Tons Pastéis":
    cor_sequencia = px.colors.qualitative.Pastel
elif paleta_selecionada == "Vibrante":
    cor_sequencia = px.colors.qualitative.Prism
elif paleta_selecionada == "Esmeralda":
    cor_sequencia = px.colors.sequential.Mint
else:
    cor_sequencia = px.colors.qualitative.Safe

tipo_grafico_5 = st.sidebar.radio(
    "Estilo do Gráfico 5:",
    options=["Barras Verticais", "Barras Horizontais"],
    index=1
)

# Dicionário de cores padrão para manter a identidade dos status nos gráficos
mapa_cores_status = {
    "APROVADO": "#2ca02c",
    "AGUARD_DEV_DO_SETOR": "#d62728",
    "EM VERF INTERNA": "#bcbd22",
    "CANCELADO": "#7f7f7f"
}

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="DADOS_GRÁFICOS")
        
        # Limpeza de espaços invisíveis nas colunas de texto
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()
                
        # Substituir strings de erro do Excel por valores nulos limpos
        df = df.replace(["#VALOR!", "0", "0.0", "NONE", "NAN"], None)
        
        # Remover linhas totalmente vazias para evitar inflar contagens
        df = df.dropna(subset=["SIGLA DO DOCUMENTO", "NOME DO DOCUMENTO", "RESPONSÁVEL"], how="all")
        
        # Padronização de termos conforme solicitado na descrição da regra de negócio
        if "STATUS DO DOCUMENTO NORMATIVO" in df.columns:
            df["STATUS DO DOCUMENTO NORMATIVO"] = df["STATUS DO DOCUMENTO NORMATIVO"].replace({
                "VERIFICADO AGUARDA DEVOLUÇÃO SETOR": "AGUARD_DEV_DO_SETOR",
                "VERIFICADO AGUARDA DEVOLUÇÃO DO SETOR": "AGUARD_DEV_DO_SETOR",
                "VERF AG DEV - SETOR": "AGUARD_DEV_DO_SETOR",
                "EM VERIFICAÇÃO": "EM VERF INTERNA"
            })
            
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        st.stop()
else:
    st.info("💡 Por favor, use o menu lateral para carregar a sua planilha Excel e ativar os gráficos interativos.")
    st.stop()

# ==============================================================================
# 3. PROCESSAMENTO DOS INDICADORES CRÍTICOS (METRICS)
# ==============================================================================
status_documento = df["STATUS DO DOCUMENTO NORMATIVO"].fillna("NÃO INFORMADO")
total_docs = len(df)
aprovados = len(df[status_documento == "APROVADO"])
verf_1 = len(df[status_documento == "AGUARD_DEV_DO_SETOR"])
verf_2 = len(df[status_documento == "EM VERF INTERNA"])

# Cálculo dinâmico da nova caixa de Média ("Temp Total até Aprov")
col_media_dias = "média I.A.A.A" if "média I.A.A.A" in df.columns else df.columns[-1]
try:
    media_valores = pd.to_numeric(df[col_media_dias], errors='coerce').dropna()
    media_dias_total = round(media_valores.mean(), 1) if len(media_valores) > 0 else 0.0
except:
    media_dias_total = 0.0

# --- EXIBIÇÃO DAS CAIXAS DE MÉTRICAS INDEPENDENTES ---
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric(label="📋 Total de Documentos", value=total_docs)
m2.metric(label="✅ Aprovados", value=aprovados)
m3.metric(label="⏰ T - 1º Verf", value=verf_1)
m4.metric(label="🔍 T - 2º Verf", value=verf_2)
m5.metric(label="📅 Temp Total até Aprov", value=f"{media_dias_total} dias" if media_dias_total > 0 else "N/A")

st.markdown("---")

# ==============================================================================
# 4. FILAS DE GRÁFICOS (ROWS 1 & 2)
# ==============================================================================
row1_col1, row1_col2 = st.columns(2)

# --- GRÁFICO 1: Doc. por Status ---
with row1_col1:
    st.subheader("1 Doc. por Status")
    status_alvo = ["APROVADO", "EM VERF INTERNA", "AGUARD_DEV_DO_SETOR", "CANCELADO"]
    df_g1_filtrado = df[df["STATUS DO DOCUMENTO NORMATIVO"].isin(status_alvo)]
    df_g1 = df_g1_filtrado["STATUS DO DOCUMENTO NORMATIVO"].value_counts().reset_index()
    df_g1.columns = ["STATUS DO DOCUMENTO NORMATIVO", "Quantidade"]
    
    if not df_g1.empty:
        fig1 = px.pie(df_g1, names="STATUS DO DOCUMENTO NORMATIVO", values="Quantidade", hole=0.4,
                      color="STATUS DO DOCUMENTO NORMATIVO", color_discrete_map=mapa_cores_status)
        fig1.update_traces(textinfo='value+label', textposition='inside')
        fig1.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
        st.plotly_chart(fig1, use_container_width=True)

# --- GRÁFICO 2: Doc. Aprovados por Tipo ---
with row1_col2:
    st.subheader("2 Doc. Aprovados por Tipo")
    df_g2_filtrado = df[df["STATUS DO DOCUMENTO NORMATIVO"] == "APROVADO"]
    df_g2 = df_g2_filtrado["SIGLA DO DOCUMENTO"].value_counts().reset_index()
    df_g2.columns = ["Tipo de Documento", "Quantidade Aprovada"]
    
    if not df_g2.empty:
        fig2 = px.bar(df_g2, x="Quantidade Aprovada", y="Tipo de Documento", text="Quantidade Aprovada",
                      orientation="h", color="Tipo de Documento", color_discrete_sequence=cor_sequencia)
        fig2.update_traces(textposition="outside")
        fig2.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")

# --- GRÁFICO 3: Validade por tipo de Doc. Aprov ---
st.subheader("3 Validade por tipo de Doc. Aprov")
col_vencido = "(Vencido, No Prazo, Prestes a Vencer)"

if col_vencido in df.columns:
    df_g3_base = df[df["STATUS DO DOCUMENTO NORMATIVO"] == "APROVADO"].copy()
    df_g3_base[col_vencido] = df_g3_base[col_vencido].replace({
        "Vencido": "Vencidos",
        "Válido": "Válidos",
        "No Prazo": "No Prazo",
        "A": "Prestes a Vencer"
    })
    
    df_g3_counts = df_g3_base.groupby(["SIGLA DO DOCUMENTO", col_vencido]).size().reset_index(name="Quantidade")
    fig3 = px.bar(df_g3_counts, x="SIGLA DO DOCUMENTO", y="Quantidade", color=col_vencido,
                  text="Quantidade", barmode="group",
                  labels={"SIGLA DO DOCUMENTO": "Tipo de Documento", col_vencido: "Status de Validade"})
    fig3.update_traces(textposition="outside")
    st.plotly_chart(fig3, use_container_width=True)

st.markdown("---")

# Tratamento para profissionais desativadas
df_prof = df.copy()
if "RESPONSÁVEL" in df_prof.columns:
    df_prof["RESPONSÁVEL"] = df_prof["RESPONSÁVEL"].replace({
        "SONALIA": "OUTROS",
        "SABRINA": "OUTROS",
        "SONALHYA": "OUTROS",
        "Sonalia": "OUTROS",
        "Sabrina": "OUTROS"
    })

# --- GRÁFICO 4: Documentos por profissional ---
st.subheader("4 Documentos por profissional")
if "RESPONSÁVEL" in df_prof.columns:
    profissionais_lista = sorted([p for p in df_prof["RESPONSÁVEL"].dropna().unique().tolist() if p != "OUTROS"])
    prof_selecionado_g4 = st.selectbox("Filtrar por profissional para o Gráfico 4:", options=["Todos"] + profissionais_lista)

    if prof_selecionado_g4 == "Todos":
        df_g4 = df_prof.copy()
    else:
        df_g4 = df_prof[df_prof["RESPONSÁVEL"] == prof_selecionado_g4]

    df_g4_counts = df_g4.groupby(["RESPONSÁVEL", "STATUS DO DOCUMENTO NORMATIVO"]).size().reset_index(name="Quantidade")
    fig4 = px.bar(
        df_g4_counts, x="Quantidade", y="RESPONSÁVEL", color="STATUS DO DOCUMENTO NORMATIVO",
        barmode="group", orientation="h", height=450, text="Quantidade",
        labels={"RESPONSÁVEL": "Profissional", "Quantidade": "N° de Documentos"},
        color_discrete_map=mapa_cores_status
    )
    fig4.update_traces(textposition="outside")
    st.plotly_chart(fig4, use_container_width=True)

st.markdown("---")

# --- GRÁFICO 5: Documentos por Tipo / Profissional ---
st.subheader("5 Documentos por Tipo / Profissional")
if "RESPONSÁVEL" in df_prof.columns:
    prof_selecionado_g5 = st.selectbox("Selecione o Responsável para Filtrar a Análise Cruzada:", options=profissionais_lista)

    df_g5_filtrado = df_prof[df_prof["RESPONSÁVEL"] == prof_selecionado_g5]
    df_g5_counts = df_g5_filtrado.groupby(["SIGLA DO DOCUMENTO", "STATUS DO DOCUMENTO NORMATIVO"]).size().reset_index(name="Quantidade")

    if not df_g5_counts.empty:
        ori_5 = "h" if tipo_grafico_5 == "Barras Horizontais" else "v"
