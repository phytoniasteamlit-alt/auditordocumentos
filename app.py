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

# --- CABEÇALHO SUPERIOR (LADO DIREITO / ESQUERDA) ---
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
# 2. PAINEL DE CONTROLE (SIDEBAR) & TRATAMENTO PROFILÁTICO DE DADOS
# ==============================================================================
st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader(
    "Carregar Planilha Excel (.xlsx):", 
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="DADOS_GRÁFICOS")
        
        # 1. Remover espaços em branco invisíveis do início e fim dos textos das colunas
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()
                
        # 2. Substituir strings de erro do Excel por valores nulos limpos
        df = df.replace(["#VALOR!", "0", "0.0", "None", "nan", "NaN"], None)
        
        # 3. Remover linhas completamente vazias para não inflar a contagem de documentos
        df = df.dropna(subset=["SIGLA DO DOCUMENTO", "NOME DO DOCUMENTO", "RESPONSÁVEL"], how="all")

        # Altera a legenda "A" para "Agd Dev Setor" em memória na coluna temporal
        col_vencido = "(Vencido, No Prazo, Prestes a Vencer)"
        if col_vencido in df.columns:
            df[col_vencido] = df[col_vencido].replace({"A": "Agd Dev Setor"})
            
        # Altera o status longo para "AG. DEV - SETOR" diretamente na memória
        if "STATUS DO DOCUMENTO NORMATIVO" in df.columns:
            df["STATUS DO DOCUMENTO NORMATIVO"] = df["STATUS DO DOCUMENTO NORMATIVO"].replace({
                "VERIFICADO AGUARDA DEVOLUÇÃO SETOR": "AG. DEV - SETOR"
            })
            
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        st.info("Certifique-se de que a aba carregada chama-se exatamente 'DADOS_GRÁFICOS'.")
        st.stop()
else:
    st.info("💡 Por favor, use o menu lateral para carregar a sua planilha Excel e ativar os gráficos interativos.")
    st.stop()

# ==============================================================================
# 3. PROCESSAMENTO DOS INDICADORES CRÍTICOS (METRICS)
# ==============================================================================
status_documento = df["STATUS DO DOCUMENTO NORMATIVO"].fillna("Não Informado")
total_docs = len(df)
aprovados = len(df[status_documento.str.upper() == "APROVADO"])

# Contagem baseada nos textos limpos e abreviados da coluna
verf_1 = len(df[status_documento.str.contains("AG. DEV - SETOR", case=False, na=False)])
verf_2 = len(df[status_documento.str.contains("EM VERIFICAÇÃO", case=False, na=False)])

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
        fig1 = px.pie(
            df_g1, names=col_vencido, values="Quantidade", hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig1.update_traces(textinfo='value+label', textposition='inside')
        fig1.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
        st.plotly_chart(fig1, use_container_width=True)
    else:
        st.warning("Sem dados suficientes para gerar o gráfico 1.")

# GRÁFICO 2: Status por Documentos
with row1_col2:
    st.subheader("2 Status por Documentos")
    df_g2 = df["STATUS DO DOCUMENTO NORMATIVO"].value_counts().reset_index()
    df_g2.columns = ["Status", "Quantidade"]
    
    if not df_g2.empty:
        fig2 = px.pie(
            df_g2, names="Status", values="Quantidade", hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        fig2.update_traces(textinfo='value+label', textposition='inside')
        fig2.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.warning("Sem dados suficientes para gerar o gráfico 2.")

st.markdown("---")

# ==============================================================================
# 5. GRÁFICOS INTERATIVOS COM MAPA DE CORES CUSTOMIZADO
# ==============================================================================

# Gráfico 3
st.subheader("3 Nº Documentos por Status")

status_disponiveis = df["STATUS DO DOCUMENTO NORMATIVO"].dropna().unique().tolist()

if status_disponiveis:
    status_selecionados = st.multiselect(
        "Filtrar por Status do Documento:",
        options=status_disponiveis,
        default=status_disponiveis
    )
    
    df_g3 = df[df["STATUS DO DOCUMENTO NORMATIVO"].isin(status_selecionados)]
    
    if not df_g3.empty:
        df_g3_counts = df_g3["STATUS DO DOCUMENTO NORMATIVO"].value_counts().reset_index()
        df_g3_counts.columns = ["Status", "Total"]
        
        fig3 = px.bar(
            df_g3_counts, x="Status", y="Total", text="Total",
            color="Status",
            labels={"Total": "N° de Documentos"},
            color_discrete_map={
                "APROVADO": "#2ca02c",
                "AG. DEV - SETOR": "#d62728",
                "EM VERIFICAÇÃO": "#bcbd22",
                "CANCELADO": "#7f7f7f"
            }
        )
        fig3.update_traces(textposition="outside")
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("Selecione pelo menos um status para renderizar o gráfico.")
else:
    st.warning("Coluna de status indisponível ou vazia.")

st.markdown("---")

# ==============================================================================
# 6. ANÁLISE POR PROFISSIONAL (EIXOS LIMPOS E EMPILHAMENTO ORGANIZADO)
# ==============================================================================
row2_col1, row2_col2 = st.columns(2)

# GRÁFICO 4: Documentos por Profissional
with row2_col1:
    st.subheader("4 Documentos por Profissional")
    
    profissionais_totais = df["RESPONSÁVEL"].dropna().unique().tolist()
    
    # Filtra removendo Sabrina e Sonalhya do selectbox de filtros ativos
    profissionais_ativos = [p for p in profissionais_totais if p.upper() not in ["SABRINA", "SONALHYA"]]
    
    if profissionais_ativos:
        prof_selecionado = st.selectbox(
            "Selecionar Profissional:",
            options=["Todos"] + profissionais_ativos
        )
        
        if prof_selecionado == "Todos":
            df_g4 = df
        else:
            df_g4 = df[df["RESPONSÁVEL"] == prof_selecionado]
            
        if not df_g4.empty:
            df_g4_counts = df_g4.groupby(["RESPONSÁVEL", "STATUS DO DOCUMENTO NORMATIVO"]).size().reset_index(name="Quantidade")
            
            # Barras horizontais agrupadas por Profissional no eixo Y
            fig4 = px.bar(
                df_g4_counts, 
                x="Quantidade", 
                y="RESPONSÁVEL", 
                color="STATUS DO DOCUMENTO NORMATIVO",
                barmode="group",
                orientation="h",
                height=500,
                text="Quantidade",
                labels={"RESPONSÁVEL": "Profissional", "Quantidade": "Nº de Documentos"},
                color_discrete_map={
                    "APROVADO": "#2ca02c",
                    "AG. DEV - SETOR": "#d62728",
                    "EM VERIFICAÇÃO": "#bcbd22",
                    "CANCELADO": "#7f7f7f"
                }
            )
            fig4.update_traces(textposition="outside")
            fig4.update_yaxes(categoryorder="total ascending")
            st.plotly_chart(fig4, use_container_width=True)
        else:
            st.info("Nenhum dado encontrado para o profissional selecionado.")
    else:
        st.warning("Coluna de profissionais indisponível.")

# GRÁFICO 5: Documentos Aprovados por Tipo (BLOCO REINDENTADO E CORRIGIDO)
with row2_col2:
    st.subheader("5 Documentos Aprovados por Tipo")
    tipos_disponiveis = df["SIGLA DO DOCUMENTO"].dropna().unique().tolist()
    
    if tipos_disponiveis:
        tipos_selecionados = st.multiselect(
            "Filtrar por Tipo de Documento (Sigla):",
            options=tipos_disponiveis, 
            default=tipos_disponiveis
        )
        
        df_g5 = df[(df["STATUS DO DOCUMENTO NORMATIVO"].str.upper() == "APROVADO") & (df["SIGLA DO DOCUMENTO"].isin(tipos_selecionados))]
        
        if not df_g5.empty:
