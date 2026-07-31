import streamlit as st
import pandas as pd
import plotly.express as px

# Configuração da página da Streamlit em modo amplo
st.set_page_config(
    page_title="Painel de Indicadores Norma Zero",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CABEÇALHO SUPERIOR (LADO DIREITO) ---
header_left, header_right = st.columns([3, 1])

with header_right:
    st.markdown(
        """
        <div style="text-align: right; line-height: 1.2; padding-bottom: 10px;">
            <span style="font-size: 16px; font-weight: bold;">🏥 Hospital da Cidade</span><br>
            <span style="font-size: 14px; color: #888;">👩‍⚕️ Coord: Fabrícia Rocha 🏆</span>
        </div>
        """, 
        unsafe_allow_html=True
    )

with header_left:
    st.title("📊 Painel de Indicadores Norma Zero")

st.markdown("---")

# --- PAINEL DE CONTROLE (SIDEBAR) ---
st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha Excel (.xlsx):", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Tratamento Profilático contra erros de dados
        df = pd.read_excel(uploaded_file, sheet_name="DADOS_GRÁFICOS")
        
        # 1. Remover espaços em branco invisíveis do início e fim dos textos das colunas
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()
                
        # 2. Substituir strings de erro do Excel (#VALOR!, 0 ou textos nulos) por valores limpos
        df = df.replace(["#VALOR!", "0", "0.0", "None", "nan", "NaN"], None)
        
        # 3. Remover linhas completamente vazias para não inflar a contagem de documentos
        df = df.dropna(subset=["SIGLA DO DOCUMENTO", "NOME DO DOCUMENTO", "RESPONSÁVEL"], how="all")

        # --- PROCESSAMENTO DOS INDICADORES (METRICS) ---
        status_documento = df["STATUS DO DOCUMENTO NORMATIVO"].fillna("Não Informado")
        
        total_docs = len(df)
        aprovados = len(df[status_documento.str.upper() == "APROVADO"])
        
        # Contagem para o 1º Verf e 2º Verf baseado na coluna de Status do Documento Normativo
        verf_1 = len(df[status_documento.str.contains("VERIFICADO AGUARDA", case=False, na=False)])
        verf_2 = len(df[status_documento.str.contains("EM VERIFICAÇÃO", case=False, na=False)])

        # --- EXIBIÇÃO DAS CAIXAS DE MÉTRICAS INDEPENDENTES ---
        m1, m2, m3, m4 = st.columns(4)
        m1.metric(label="📄 Total de Documentos", value=total_docs)
        m2.metric(label="✅ Aprovados", value=aprovados)
        m3.metric(label="⏰ T - 1º Verf", value=verf_1)
        m4.metric(label="⏰ T - 2º Verf", value=verf_2)
        
        st.markdown("---")

        # --- FILAS DE GRÁFICOS (ROWS) ---
        
        # --- BLOC0 1: GRÁFICOS DE VISÃO GERAL (Pizzas/Roscas) ---
        row1_col1, row1_col2 = st.columns(2)
        
        with row1_col1:
            st.subheader("1️⃣ Válidos, Vencidos, no Prazo")
            df_g1 = df["(Vencido, No Prazo, Prestes a Vencer)"].value_counts().reset_index()
            df_g1.columns = ["Status Temporal", "Quantidade"]
            
            if not df_g1.empty:
                fig1 = px.pie(df_g1, names="Status Temporal", values="Quantidade", hole=0.4,
                              color_discrete_sequence=px.colors.qualitative.Pastel)
                fig1.update_layout(margin=dict(l=20, r=20, t=30, b=20))
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.warning("Sem dados suficientes para gerar o gráfico 1.")

        with row1_col2:
            st.subheader("2️⃣ Status por Documentos")
            df_g2 = df["STATUS DO DOCUMENTO NORMATIVO"].value_counts().reset_index()
            df_g2.columns = ["Status", "Quantidade"]
            
            if not df_g2.empty:
                fig2 = px.pie(df_g2, names="Status", values="Quantidade", hole=0.4,
                              color_discrete_sequence=px.colors.qualitative.Safe)
                fig2.update_layout(margin=dict(l=20, r=20, t=30, b=20))
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.warning("Sem dados suficientes para gerar o gráfico 2.")

        st.markdown("---")

        # --- BLOCO 2: GRÁFICOS INTERATIVOS COM FILTROS MULTIPLOS ---
        
        # Gráfico 3: Nº de documentos por status com filtro dinâmico
        st.subheader("3️⃣ Número de Documentos por Status")
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
                fig3 = px.bar(df_g3_counts, x="Status", y="Total", text="Total", color="Status",
                              labels={"Total": "Nº de Documentos"}, color_discrete_sequence=px.colors.qualitative.Set2)
                fig3.update_traces(textposition="outside")
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.info("Selecione pelo menos um status para renderizar o gráfico.")
        else:
            st.warning("Coluna de status indisponível ou vazia.")

        st.markdown("---")
        
        row2_col1, row2_col2 = st.columns(2)

        # Gráfico 4: Documentos por Profissional
        with row2_col1:
            st.subheader("4️⃣ Documentos por Profissional")
            profissionais = df["RESPONSÁVEL"].dropna().unique().tolist()
            
            if profissionais:
                prof_selecionado = st.selectbox("Selecionar Profissional:", options=["Todos"] + profissionais)
                
                if prof_selecionado == "Todos":
                    df_g4 = df
                else:
                    df_g4 = df[df["RESPONSÁVEL"] == prof_selecionado]
                
                if not df_g4.empty:
                    df_g4_counts = df_g4.groupby(["RESPONSÁVEL", "STATUS DO DOCUMENTO NORMATIVO"]).size().reset_index(name="Quantidade")
                    fig4 = px.bar(df_g4_counts, x="STATUS DO DOCUMENTO NORMATIVO", y="Quantidade", color="STATUS DO DOCUMENTO NORMATIVO",
                                  facet_col="RESPONSÁVEL", facet_col_wrap=2, barmode="group")
                    st.plotly_chart(fig4, use_container_width=True)
                else:
                    st.info("Nenhum dado encontrado para o profissional selecionado.")
            else:
                st.warning("Coluna de profissionais indisponível.")

        # Gráfico 5: Documentos Aprovados por Tipo
        with row2_col2:
            st.subheader("5️⃣ Documentos Aprovados por Tipo")
            tipos_disponiveis = df["SIGLA DO DOCUMENTO"].dropna().unique().tolist()
            
            if tipos_disponiveis:
                tipos_selecionados = st.multiselect(
                    "Filtrar por Tipo de Documento (Sigla):", 
                    options=tipos_disponiveis, 
                    default=tipos_disponiveis
                )
                
                df_g5 = df[(df["STATUS DO DOCUMENTO NORMATIVO"].str.upper() == "APROVADO") & 
                           (df["SIGLA DO DOCUMENTO"].isin(tipos_selecionados))]
                
                if not df_g5.empty:
                    df_g5_counts = df_g5["SIGLA DO DOCUMENTO"].value_counts().reset_index()
                    df_g5_counts.columns = ["Tipo de Documento", "Quantidade Aprovada"]
                    fig5 = px.bar(df_g5_counts, x="Tipo de Documento", y="Quantidade Aprovada", text="Quantidade Aprovada",
                                  color="Tipo de Documento", color_discrete_sequence=px.colors.qualitative.Bold)
                    fig5.update_traces(textposition="outside")
                    st.plotly_chart(fig5, use_container_width=True)
                else:
                    st.info("Nenhum documento 'Aprovado' encontrado para os filtros selecionados.")
            else:
                st.warning("Coluna de siglas indisponível.")

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        st.info("Certifique-se de que a aba carregada chama-se exatamente 'DADOS_GRÁFICOS'.")
else:
    st.info("💡 Por favor, use o menu lateral para carregar a sua planilha Excel e ativar os gráficos interativos.")
