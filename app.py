import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata

st.set_page_config(
    page_title="Painel de Indicadores Norma Zero",
    layout="wide",
    initial_sidebar_state="expanded"
)

def normalizar_texto(texto):
    if pd.isna(texto) or not isinstance(texto, str):
        return ""
    texto = texto.strip().upper()
    texto = "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto

header_left, header_right = st.columns(2)

with header_left:
    st.markdown("<h1 style='margin: 0; padding: 0; font-size: 2.2rem;'>📊 Painel de Indicadores Norma Zero</h1>", unsafe_allow_html=True)

with header_right:
    st.markdown('<div style="text-align: right; line-height: 1.2; padding-bottom: 10px;"><span style="font-size: 16px; font-weight: bold;">🏥 Hospital da Cidade</span><br><span style="font-size: 14px; color: #888;">👩‍💼 Coord: Fabrícia Rocha</span></div>', unsafe_allow_html=True)

st.markdown("---")

st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha Excel (.xlsx):", type=["xlsx"])

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

tipo_grafico_5 = st.sidebar.radio("Estilo do Gráfico 5:", options=["Barras Verticais", "Barras Horizontais"], index=1)

mapa_cores_status = {
    "APROVADO": "#2ca02c",
    "AGUARD_DEV_DO_SETOR": "#d62728",
    "EM VERF INTERNA": "#bcbd22",
    "CANCELADO": "#7f7f7f"
}

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="DADOS_GRÁFICOS")
        for col in df.columns:
            df = df.rename(columns={col: col.strip().upper()})
        
        # Mapeamento Inteligente de Colunas por Aproximação
        col_status = None
        col_sigla = None
        col_responsavel = None
        
        for col in df.columns:
            if "STATUS" in col:
                col_status = col
            elif "SIGLA" in col or "TIPO" in col:
                col_sigla = col
            elif "RESPONS" in col or "COORD" in col:
                col_responsavel = col
                
        col_status = col_status if col_status else df.columns[0]
        col_sigla = col_sigla if col_sigla else df.columns[0]
        col_responsavel = col_responsavel if col_responsavel else df.columns[0]
        
        df[col_status] = df[col_status].fillna("NÃO INFORMADO")
        df[col_sigla] = df[col_sigla].fillna("NÃO INFORMADO")
        df[col_responsavel] = df[col_responsavel].fillna("NÃO INFORMADO")
        
        df["STATUS_NORM"] = df[col_status].apply(normalizar_texto)
        df["RESP_NORM"] = df[col_responsavel].apply(normalizar_texto)
        
        df["STATUS_FINAL"] = df["STATUS_NORM"].replace({
            "VERIFICADO AGUARDA DEVOLUCO AMBOS": "AGUARD_DEV_DO_SETOR",
            "VERIFICADO AGUARDA DEVOLUCAO SETOR": "AGUARD_DEV_DO_SETOR",
            "VERIFICADO AGUARDA DEVOLUCAO DO SETOR": "AGUARD_DEV_DO_SETOR",
            "VERF AG DEV - SETOR": "AGUARD_DEV_DO_SETOR",
            "EM VERIFICACOES": "EM VERF INTERNA",
            "EM VERIFIKACAO": "EM VERF INTERNA",
            "EM VERIFICACAO": "EM VERF INTERNA",
            "EM VERIFICAÇÃO": "EM VERF INTERNA",
            "APROVADO": "APROVADO",
            "CANCELADO": "CANCELADO"
        })
        
        df["STATUS_LIMPO_GRAFICO"] = df["STATUS_FINAL"].apply(lambda x: x if x in mapa_cores_status else "OUTROS")
        
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
        st.stop()
else:
    st.info("💡 Por favor, use o menu lateral para carregar a sua planilha Excel e ativar os gráficos interativos.")
    st.stop()

total_docs = len(df)
aprovados = len(df[df["STATUS_LIMPO_GRAFICO"] == "APROVADO"])
verf_1 = len(df[df["STATUS_LIMPO_GRAFICO"] == "AGUARD_DEV_DO_SETOR"])
verf_2 = len(df[df["STATUS_LIMPO_GRAFICO"] == "EM VERF INTERNA"])

col_media_dias = None
for col in df.columns:
    if "MÉDIA" in col or "MEDIA" in col or "TEMPO" in col or "DIAS" in col or "I.A.A.A" in col:
        col_media_dias = col
        break

media_dias_total = 0.0
if col_media_dias:
    try:
        df_aprovados_apenas = df[df["STATUS_LIMPO_GRAFICO"] == "APROVADO"]
        media_valores = pd.to_numeric(df_aprovados_apenas[col_media_dias], errors='coerce').dropna()
        if len(media_valores) > 0:
            media_dias_total = round(media_valores.mean(), 1)
    except:
        pass

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric(label="📋 Total de Documentos", value=total_docs)
m2.metric(label="✅ Aprovados", value=aprovados)
m3.metric(label="⏰ T - 1º Verf", value=verf_1)
m4.metric(label="🔍 T - 2º Verf", value=verf_2)
m5.metric(label="📅 Temp Total até Aprov", value=f"{media_dias_total} dias" if media_dias_total > 0 else "N/A")

st.markdown("---")

row1_col1, row1_col2 = st.columns(2)

with row1_col1:
    st.subheader("1 Doc. por Status")
    df_g1 = df[df["STATUS_LIMPO_GRAFICO"].isin(mapa_cores_status.keys())]
    df_g1_counts = df_g1["STATUS_LIMPO_GRAFICO"].value_counts().reset_index()
    df_g1_counts.columns = ["STATUS DO DOCUMENTO NORMATIVO", "Quantidade"]
    if not df_g1_counts.empty:
        fig1 = px.pie(df_g1_counts, names="STATUS DO DOCUMENTO NORMATIVO", values="Quantidade", hole=0.4, color="STATUS DO DOCUMENTO NORMATIVO", color_discrete_map=mapa_cores_status)
        fig1.update_traces(textinfo='value+label', textposition='inside', insidetextfont=dict(size=14))
        fig1.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
        st.plotly_chart(fig1, use_container_width=True)

with row1_col2:
    st.subheader("2 Doc. Aprovados por Tipo")
    df_g2_filtrado = df[df["STATUS_LIMPO_GRAFICO"] == "APROVADO"]
    df_g2 = df_g2_filtrado[col_sigla].value_counts().reset_index()
    df_g2.columns = ["Tipo de Documento", "Quantidade Aprovada"]
    if not df_g2.empty:
        fig2 = px.bar(df_g2, x="Quantidade Aprovada", y="Tipo de Documento", text="Quantidade Aprovada", orientation="h", color="Tipo de Documento", color_discrete_sequence=cor_sequencia)
        fig2.update_traces(textposition="outside", textfont=dict(size=15))
        fig2.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False, xaxis=dict(title_font=dict(size=14), tickfont=dict(size=13)), yaxis=dict(title_font=dict(size=14), tickfont=dict(size=13)))
        st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")

st.subheader("3 Validade por Tipo de Documentos")
col_real_g3 = None
for col in df.columns:
    if "VENCIDO" in col or "VALIDADE" in col:
        col_real_g3 = col
        break

if col_real_g3:
    df_g3_base = df[df["STATUS_LIMPO_GRAFICO"] == "APROVADO"].copy()
    df_g3_base["VAL_NORM"] = df_g3_base[col_real_g3].apply(normalizar_texto)
    df_g3_base["STATUS_VAL"] = df_g3_base["VAL_NORM"].replace({
        "VENCIDO": "Vencidos", "Vencido": "Vencidos", "vencido": "Vencidos",
        "VALIDO": "Válidos", "Válido": "Válidos", "valido": "Válidos", "VÁLIDOS": "Válidos",
        "NO PRAZO": "No Prazo", "No Prazo": "No Prazo", "no prazo": "No Prazo",
        "A": "Prestes a Vencer", "PRESTES A VENCER": "Prestes a Vencer", "Prestes a Vencer": "Prestes a Vencer"
    })
    status_validade_disponiveis = ["Válidos", "Vencidos", "No Prazo", "Prestes a Vencer"]
    validade_selecionada = st.multiselect("Filtrar Status de Validade:", options=status_validade_disponiveis, default=status_validade_disponiveis)
    df_g3_filtrado_val = df_g3_base[df_g3_base["STATUS_VAL"].isin(validade_selecionada)]
    if not df_g3_filtrado_val.empty:
        df_g3_counts = df_g3_filtrado_val.groupby([col_sigla, "STATUS_VAL"]).size().reset_index(name="Quantidade")
        fig3 = px.bar(df_g3_counts, x=col_sigla, y="Quantidade", color="STATUS_VAL", text="Quantidade", barmode="group", labels={col_sigla: "Tipo de Documento", "STATUS_VAL": "Status de Validade"})
        fig3.update_traces(textposition="outside", textfont=dict(size=15))
        fig3.update_layout(legend=dict(font=dict(size=13), title_font=dict(size=14)), xaxis=dict(title_font=dict(size=14), tickfont=dict(size=13)), yaxis=dict(title_font=dict(size=14), tickfont=dict(size=13)))
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("Nenhum dado encontrado para os status de validade selecionados.")

st.markdown("---")

df["RESPONSAVEL_FINAL"] = df.apply(
    lambda row: "OUTROS" if row["RESP_NORM"] in ["SONALIA", "SABRINA", "SONALHYA"] else (row[col_responsavel] if pd.notna(row[col_responsavel]) else "NÃO INFORMADO"),
    axis=1
)
df["RESP_FINAL_NORM"] = df["RESPONSAVEL_FINAL"].apply(normalizar_texto)
profissionais_lista = sorted(list(set([str(p) for p in df["RESPONSAVEL_FINAL"].unique() if p and str(p).upper() != "OUTROS" and str(p).upper() != "NÃO INFORMADO"])))

st.subheader("4 Documentos por profissional")
prof_selecionado_g4 = st.selectbox("Filtrar por profissional para o Gráfico 4:", options=["Todos"] + profissionais_lista)
df_g4 = df.copy() if prof_selecionado_g4 == "Todos" else df[df["RESP_FINAL_NORM"] == normalizar_texto(prof_selecionado_g4)]
df_g4_counts = df_g4.groupby(["RESPONSAVEL_FINAL", "STATUS_LIMPO_GRAFICO"]).size().reset_index(name="Quantidade")

fig4 = px.bar(df_g4_counts, x="Quantidade", y="RESPONSAVEL_FINAL", color="STATUS_LIMPO_GRAFICO", barmode="group", orientation="h", height=500, text="Quantidade", labels={"RESPONSAVEL_FINAL": "Profissional", "STATUS_LIMPO_GRAFICO": "Status"}, color_discrete_map=mapa_cores_status)
