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

# --- CABEÇALHO SUPERIOR ---
header_left, header_right = st.columns(2)

with header_right:
    st.markdown(
        """
        <div style="text-align: right; line-height: 1.2; padding-bottom: 10px;">
            <span style="font-size: 16px; font-weight: bold;">🏥 Hospital da Cidade</span><br>
            <span style="font-size: 14px; color: #888;">Coord: Fabrícia Rocha</span>
        </div>
        """,
        unsafe_allow_html=True
    )

with header_left:
    st.title("📊 Painel de Indicadores Norma Zero")

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
    index=0
)

if paleta_selecionada == "Tons Pastéis":
    cor_sequencia = px.colors.qualitative.Pastel
elif paleta_selecionada == "Vibrante":
    cor_sequencia = px.colors.qualitative.Prism
elif paleta_selecionada == "Esmeralda":
    cor_sequencia = px.colors.sequential.Mint
else:
    cor_sequencia = px.colors.qualitative.Safe

tipo_grafico_5 = st.sidebar.radio("Estilo do Gráfico 5:", options=["Barras Verticais", "Barras Horizontais"], index=0)
tipo_grafico_6 = st.sidebar.radio("Estilo do Gráfico 6:", options=["Barras Verticais", "Barras Horizontais"], index=0)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="DADOS_GRÁFICOS")
        
        # Limpeza de espaços invisíveis e conversão forçada para maiúsculas
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip().str.upper()
                
        # Substituir strings de erro do Excel por valores nulos limpos
        df = df.replace(["#VALOR!", "0", "0.0", "NONE", "NAN", "NAN"], None)
        
        # Remover linhas totalmente vazias para evitar inflar contagens
        df = df.dropna(subset=["SIGLA DO DOCUMENTO", "NOME DO DOCUMENTO", "RESPONSÁVEL"], how="all")

        # Altera a legenda "A" para "AGD DEV SETOR" em memória na coluna temporal
        col_vencido = "(Vencido, No Prazo, Prestes a Vencer)"
        if col_vencido in df.columns:
            df[col_vencido] = df[col_vencido].replace({"A": "AGD DEV SETOR"})
            
        # Altera o status longo para "AG. DEV - SETOR" diretamente na memória
        if "STATUS DO DOCUMENTO NORMATIVO" in df.columns:
            df["STATUS DO DOCUMENTO NORMATIVO"] = df["STATUS DO DOCUMENTO NORMATIVO"].replace({
                "VERIFICADO AGUARDA DEVOLUÇÃO SETOR": "AG. DEV - SETOR"
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

# Filtros diretos por igualdade para evitar quebras por linhas vazias
verf_1 = len(df[status_documento == "AG. DEV - SETOR"])
verf_2 = len(df[status_documento == "EM VERIFICAÇÃO"])

# --- EXIBIÇÃO DAS CAIXAS DE MÉTRICAS INDEPENDENTES ---
m1, m2, m3, m4 = st.columns(4)
m1.metric(label="📄 Total de Documentos", value=total_docs)
m2.metric(label="✅ Aprovados", value=aprovados)
m3.metric(label="⏰ T - 1º Verf", value=verf_1)
m4.metric(label="⏰ T - 2º Verf", value=verf_2)

st.markdown("---")

# ==============================================================================
# 4. FILAS DE GRÁFICOS (ROWS 1 & 2) - EXIBIÇÃO DE QUANTIDADES ABSOLUTAS
# ==============================================================================
row1_col1, row1_col2 = st.columns(2)

# GRÁFICO 1: Status Temporal
with row1_col1:
    st.subheader("1 Válidos, Vencidos, no Prazo")
    df_g1 = df[col_vencido].value_counts().reset_index()
    df_g1.columns = [col_vencido, "Quantidade"]
    
    if not df_g1.empty:
        fig1 = px.pie(df_g1, names=col_vencido, values="Quantidade", hole=0.4, color_discrete_sequence=cor_sequencia)
        fig1.update_traces(textinfo='value+label', textposition='inside')
        fig1.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
        st.plotly_chart(fig1, use_container_width=True)

# GRÁFICO 2: Status por Documentos
with row1_col2:
    st.subheader("2 Status por Documentos")
    df_g2 = df["STATUS DO DOCUMENTO NORMATIVO"].value_counts().reset_index()
    df_g2.columns = ["Status", "Quantidade"]
    
    if not df_g2.empty:
        fig2 = px.pie(df_g2, names="Status", values="Quantidade", hole=0.4, color_discrete_sequence=cor_sequencia)
        fig2.update_traces(textinfo='value+label', textposition='inside')
        fig2.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")

# ==============================================================================
# 5. GRÁFICOS INTERATIVOS COM MAPA DE CORES CUSTOMIZADO
# ==============================================================================

# Gráfico 3
st.subheader("3 Nº Documentos por Status")
status_disponiveis = df["STATUS DO DOCUMENTO NORMATIVO"].dropna().unique().tolist()

status_selecionados = st.multiselect(
    "Filtrar por Status do Documento:",
    options=status_disponiveis,
    default=status_disponiveis
)

df_g3 = df[df["STATUS DO DOCUMENTO NORMATIVO"].isin(status_selecionados)]
df_g3_counts = df_g3["STATUS DO DOCUMENTO NORMATIVO"].value_counts().reset_index()
df_g3_counts.columns = ["Status", "Total"]

fig3 = px.bar(
    df_g3_counts, x="Status", y="Total", text="Total", color="Status",
    labels={"Total": "N° de Documentos"},
    color_discrete_map={"APROVADO": "#2ca02c", "AG. DEV - SETOR": "#d62728", "EM VERIFICAÇÃO": "#bcbd22", "CANCELADO": "#7f7f7f"}
)
fig3.update_traces(textposition="outside")
st.plotly_chart(fig3, use_container_width=True)

st.markdown("---")

# ==============================================================================
# 6. GRÁFICO 4: DOCUMENTOS POR PROFISSIONAL
# ==============================================================================
st.subheader("4 Documentos por Profissional")
profissionais_totais = df["RESPONSÁVEL"].dropna().unique().tolist()
profissionais_ativos = [p for p in profissionais_totais if p not in ["SABRINA", "SONALHYA", "SONALIA"]]

prof_selecionado = st.selectbox("Selecionar Profissional para Análise:", options=["Todos"] + profissionais_ativos)

if prof_selecionado == "Todos":
    df_g4 = df[df["RESPONSÁVEL"].isin(profissionais_ativos)]
else:
    df_g4 = df[df["RESPONSÁVEL"] == prof_selecionado]
    
df_g4_counts = df_g4.groupby(["RESPONSÁVEL", "STATUS DO DOCUMENTO NORMATIVO"]).size().reset_index(name="Quantidade")

fig4 = px.bar(
    df_g4_counts, x="Quantidade", y="RESPONSÁVEL", color="STATUS DO DOCUMENTO NORMATIVO",
    barmode="group", orientation="h", height=450, text="Quantidade",
    labels={"RESPONSÁVEL": "Profissional", "Quantidade": "Nº de Documentos"},
    category_orders={"STATUS DO DOCUMENTO NORMATIVO": ["APROVADO", "AG. DEV - SETOR", "EM VERIFICAÇÃO", "CANCELADO"]},
    color_discrete_map={"APROVADO": "#2ca02c", "AG. DEV - SETOR": "#d62728", "EM VERIFICAÇÃO": "#bcbd22", "CANCELADO": "#7f7f7f"}
)
fig4.update_traces(textposition="outside")
fig4.update_yaxes(categoryorder="total ascending")
st.plotly_chart(fig4, use_container_width=True)

st.markdown("---")

# ==============================================================================
# 7. GRÁFICO 5: DOCUMENTOS APROVADOS POR TIPO
# ==============================================================================
st.subheader("5 Documentos Aprovados por Tipo")
tipos_disponiveis = df["SIGLA DO DOCUMENTO"].dropna().unique().tolist()

tipos_selecionados = st.multiselect(
    "Filtrar por Tipo de Documento (Sigla):", 
    options=tipos_disponiveis, 
    default=tipos_disponiveis
)

df_g5 = df[(df["STATUS DO DOCUMENTO NORMATIVO"] == "APROVADO") & (df["SIGLA DO DOCUMENTO"].isin(tipos_selecionados))]
df_g5_counts = df_g5["SIGLA DO DOCUMENTO"].value_counts().reset_index()
df_g5_counts.columns = ["Tipo de Documento", "Quantidade Aprovada"]

g5_horizontal = (tipo_grafico_5 == "Barras Horizontais")
eixo_x_5 = "Quantidade Aprovada" if g5_horizontal else "Tipo de Documento"
eixo_y_5 = "Tipo de Documento" if g5_horizontal else "Quantidade Aprovada"
ori_5 = "h" if g5_horizontal else "v"

fig5 = px.bar(df_g5_counts, x=eixo_x_5, y=eixo_y_5, text="Quantidade Aprovada", color="Tipo de Documento", orientation=ori_5, color_discrete_sequence=cor_sequencia)
fig5.update_traces(textposition="outside")
st.plotly_chart(fig5, use_container_width=True)

st.markdown("---")

# ==============================================================================
# 8. GRÁFICO 6: DETALHAMENTO CRUZADO DE TIPOS DE DOCUMENTO POR STATUS
