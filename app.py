import streamlit st
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
# 2. PAINEL DE CONTROLE FIXO NO TOPO (NÃO TRAVA A ROLAGEM DA TELA)
# ==============================================================================
uploaded_file = st.file_uploader(
    "📂 Passo 1: Carregue sua Planilha Excel aqui (.xlsx):", 
    type=["xlsx"]
)

st.markdown("---")

# Campo de customização visual fixo e visível no topo da página principal
st.subheader("🎨 Passo 2: Customização Visual e Estilo")
c1, c2, c3 = st.columns(3)
with c1:
    paleta_selecionada = st.selectbox(
        "Escolha o Tema de Cores Geral:",
        options=["Padrão Hospitalar", "Tons Pastéis", "Vibrante", "Esmeralda"],
        index=0
    )
with c2:
    tipo_grafico_5 = st.radio("Orientação do Gráfico 5:", options=["Vertical", "Horizontal"], index=0, horizontal=True)
with c3:
    tipo_grafico_6 = st.radio("Orientação do Gráfico 6:", options=["Vertical", "Horizontal"], index=0, horizontal=True)

if paleta_selecionada == "Tons Pastéis":
    cor_sequencia = px.colors.qualitative.Pastel
elif paleta_selecionada == "Vibrante":
    cor_sequencia = px.colors.qualitative.Prism
elif paleta_selecionada == "Esmeralda":
    cor_sequencia = px.colors.sequential.Mint
else:
    cor_sequencia = px.colors.qualitative.Safe

st.markdown("---")

# ==============================================================================
# 3. LEITURA E PADRONIZAÇÃO COMPLETA DA PLANILHA
# ==============================================================================
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="DADOS_GRÁFICOS")
        
        # Limpa espaços invisíveis e garante que textos fiquem legíveis
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip()
                
        df = df.replace(["#VALOR!", "0", "0.0", "None", "nan", "NaN"], None)
        df = df.dropna(subset=["SIGLA DO DOCUMENTO", "NOME DO DOCUMENTO", "RESPONSÁVEL"], how="all")

        # Tratamento das legendas solicitadas
        col_vencido = "(Vencido, No Prazo, Prestes a Vencer)"
        if col_vencido in df.columns:
            df[col_vencido] = df[col_vencido].replace({"A": "Agd Dev Setor"})
            
        if "STATUS DO DOCUMENTO NORMATIVO" in df.columns:
            df["STATUS DO DOCUMENTO NORMATIVO"] = df["STATUS DO DOCUMENTO NORMATIVO"].replace({
                "VERIFICADO AGUARDA DEVOLUÇÃO SETOR": "AG. DEV - SETOR"
            })
            
    except Exception as e:
        st.error(f"Erro crítico ao processar o arquivo: {e}")
        st.stop()
else:
    st.info("💡 Aguardando o carregamento da planilha Excel para ativar todos os 7 gráficos do painel.")
    st.stop()

# ==============================================================================
# 4. EXIBIÇÃO DAS MÉTRICAS DO TOPO
# ==============================================================================
status_documento = df["STATUS DO DOCUMENTO NORMATIVO"].fillna("Não Informado")
total_docs = len(df)
aprovados = len(df[status_documento.str.upper() == "APROVADO"])
verf_1 = len(df[status_documento.str.upper() == "AG. DEV - SETOR"])
verf_2 = len(df[status_documento.str.upper() == "EM VERIFICAÇÃO"])

m1, m2, m3, m4 = st.columns(4)
m1.metric(label="📄 Total de Documentos", value=total_docs)
m2.metric(label="✅ Aprovados", value=aprovados)
m3.metric(label="⏰ T - 1º Verf", value=verf_1)
m4.metric(label="⏰ T - 2º Verf", value=verf_2)

st.markdown("---")

# ==============================================================================
# 5. FILA DE PIZZAS (GRÁFICOS 1 & 2)
# ==============================================================================
row1_col1, row1_col2 = st.columns(2)

with row1_col1:
    st.subheader("1 Válidos, Vencidos, no Prazo")
    df_g1 = df[col_vencido].value_counts().reset_index()
    df_g1.columns = [col_vencido, "Quantidade"]
    fig1 = px.pie(df_g1, names=col_vencido, values="Quantidade", hole=0.4, color_discrete_sequence=cor_sequencia)
    fig1.update_traces(textinfo='value+label', textposition='inside')
    fig1.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
    st.plotly_chart(fig1, use_container_width=True)

with row1_col2:
    st.subheader("2 Status por Documentos")
    df_g2 = df["STATUS DO DOCUMENTO NORMATIVO"].value_counts().reset_index()
    df_g2.columns = ["Status", "Quantidade"]
    fig2 = px.pie(df_g2, names="Status", values="Quantidade", hole=0.4, color_discrete_sequence=cor_sequencia)
    fig2.update_traces(textinfo='value+label', textposition='inside')
    fig2.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
    st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")

# ==============================================================================
# 6. GRÁFICO 3: N° DOCUMENTOS POR STATUS
# ==============================================================================
st.subheader("3 Nº Documentos por Status")
status_disponiveis = df["STATUS DO DOCUMENTO NORMATIVO"].dropna().unique().tolist()
status_selecionados = st.multiselect("Filtrar por Status do Documento:", options=status_disponiveis, default=status_disponiveis)

df_g3 = df[df["STATUS DO DOCUMENTO NORMATIVO"].isin(status_selecionados)]
df_g3_counts = df_g3["STATUS DO DOCUMENTO NORMATIVO"].value_counts().reset_index()
df_g3_counts.columns = ["Status", "Total"]

fig3 = px.bar(
    df_g3_counts, x="Status", y="Total", text="Total", color="Status", labels={"Total": "N° de Documentos"},
    color_discrete_map={"APROVADO": "#2ca02c", "AG. DEV - SETOR": "#d62728", "EM VERIFICAÇÃO": "#bcbd22", "CANCELADO": "#7f7f7f"}
)
fig3.update_traces(textposition="outside")
st.plotly_chart(fig3, use_container_width=True)

st.markdown("---")

# ==============================================================================
# 7. GRÁFICO 4: DOCUMENTOS POR PROFISSIONAL
# ==============================================================================
st.subheader("4 Documentos por Profissional")
profissionais_totais = df["RESPONSÁVEL"].dropna().unique().tolist()
profissionais_ativos = [p for p in profissionais_totais if p.upper() not in ["SABRINA", "SONALHYA", "SONALIA"]]

prof_selecionado = st.selectbox("Selecionar Profissional para Análise:", options=["Todos"] + profissionais_ativos)

if str(prof_selecionado).upper() == "TODOS":
    df_g4 = df[df["RESPONSÁVEL"].str.upper().isin([p.upper() for p in profissionais_ativos])]
else:
    df_g4 = df[df["RESPONSÁVEL"].str.upper() == str(prof_selecionado).upper()]
    
df_g4_counts = df_g4.groupby(["RESPONSÁVEL", "STATUS DO DOCUMENTO NORMATIVO"]).size().reset_index(name="Quantidade")

fig4 = px.bar(
    df_g4_counts, x="Quantidade", y="RESPONSÁVEL", color="STATUS DO DOCUMENTO NORMATIVO", barmode="group", orientation="h", height=450, text="Quantidade",
    labels={"RESPONSÁVEL": "Profissional", "Quantidade": "Nº de Documentos"},
    category_orders={"STATUS DO DOCUMENTO NORMATIVO": ["APROVADO", "AG. DEV - SETOR", "EM VERIFICAÇÃO", "CANCELADO"]},
    color_discrete_map={"APROVADO": "#2ca02c", "AG. DEV - SETOR": "#d62728", "EM VERIFICAÇÃO": "#bcbd22", "CANCELADO": "#7f7f7f"}
)
fig4.update_traces(textposition="outside")
fig4.update_yaxes(categoryorder="total ascending")
st.plotly_chart(fig4, use_container_width=True)

st.markdown("---")

# ==============================================================================
# 8. GRÁFICO 5: DOCUMENTOS APROVADOS POR TIPO
# ==============================================================================
st.subheader("5 Documentos Aprovados por Tipo")
tipos_disponiveis = df["SIGLA DO DOCUMENTO"].dropna().unique().tolist()
tipos_selecionados = st.multiselect("Filtrar por Tipo de Documento (Sigla):", options=tipos_disponiveis, default=tipos_disponiveis)

df_g5 = df[(df["STATUS DO DOCUMENTO NORMATIVO"].str.upper() == "APROVADO") & (df["SIGLA DO DOCUMENTO"].isin(tipos_selecionados))]
df_g5_counts = df_g5["SIGLA DO DOCUMENTO"].value_counts().reset_index()
df_g5_counts.columns = ["Tipo de Documento", "Quantidade Aprovada"]

ori_5 = "h" if tipo_grafico_5 == "Horizontal" else "v"
x_5 = "Quantidade Aprovada" if tipo_grafico_5 == "Horizontal" else "Tipo de Documento"
y_5 = "Tipo de Documento" if tipo_grafico_5 == "Horizontal" else "Quantidade Aprovada"

# [CORREÇÃO VISUAL] Váriavel corrigida para 'cor_sequencia' eliminando o travamento do rodapé
fig5 = px.bar(df_g5_counts, x=x_5, y=y_5, text="Quantidade Aprovada", color="Tipo de Documento", orientation=ori_5, color_discrete_sequence=cor_sequencia)
fig5.update_traces(textposition="outside")
st.plotly_chart(fig5, use_container_width=True)

st.markdown("---")

# ==============================================================================
# 9. GRÁFICO 6: DETALHAMENTO CRUZADO POR PROFISSIONAL
# ==============================================================================
st.subheader("6 Detalhamento de Tipos de Documento por Status e Profissional")
