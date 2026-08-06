import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata

# Configuração da página mantida idêntica
st.set_page_config(
    page_title="Painel de Indicadores Norma Zero",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Função auxiliar para remover acentos e espaços invisíveis
def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = texto.strip().upper()
    return "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

header_left, header_right = st.columns(2)

# Título principal mantido com o emoji original
with header_left:
    st.markdown("<h1 style='margin: 0; padding: 0; font-size: 2.2rem;'>📊 Painel de Indicadores Norma Zero</h1>", unsafe_allow_html=True)

# Identificação do hospital e da coordenadora mantidos intactos
with header_right:
    st.markdown('<div style="text-align: right; line-height: 1.2; padding-bottom: 10px;"><span style="font-size: 16px; font-weight: bold;">🏥 Hospital da Cidade</span><br><span style="font-size: 14px; color: #888;">👩‍💼 Coord: Fabrícia Rocha</span></div>', unsafe_allow_html=True)

st.markdown("---")

st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha Excel (.xlsx):", type=["xlsx"])
st.sidebar.markdown("---")

# Customização visual da barra lateral mantida idêntica
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
        # Lendo de volta estritamente a aba DADOS_GRÁFICOS
        df = pd.read_excel(uploaded_file, sheet_name="DADOS_GRÁFICOS")
        
        for col in df.columns:
            df = df.rename(columns={col: col.strip().upper()})
        
        for col in df.columns:
            if df[col].dtype == "object":
                df[col] = df[col].astype(str).str.strip().str.upper()
                
        df = df.replace(["#VALOR!", "0", "0.0", "NONE", "NAN"], None)
        df = df.dropna(subset=["SIGLA DO DOCUMENTO", "NOME DO DOCUMENTO", "RESPONSÁVEL"], how="all")
        
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

status_documento = df["STATUS DO DOCUMENTO NORMATIVO"].fillna("NÃO INFORMADO")
total_docs = len(df)

aprovados = len(df[status_documento == "APROVADO"])
verf_1 = len(df[status_documento == "AGUARD_DEV_DO_SETOR"])
verf_2 = len(df[status_documento == "EM VERF INTERNA"])

col_media_dias = "MÉDIA I.A.A.A" if "MÉDIA I.A.A.A" in df.columns else df.columns[-1]

try:
    media_valores = pd.to_numeric(df[col_media_dias], errors='coerce').dropna()
    media_dias_total = round(media_valores.mean(), 1) if len(media_valores) > 0 else 0.0
except:
    media_dias_total = 0.0

# Linha de blocos de métricas originais com todos os emojis mantidos
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric(label="📄 Total de Documentos", value=total_docs)
m2.metric(label="✅ Aprovados", value=aprovados)
m3.metric(label="📅 T - 1º Verf", value=verf_1)
m4.metric(label="🔍 T - 2º Verf", value=verf_2)
m5.metric(label="⏳ Temp Total até Aprov", value=f"{media_dias_total} dias" if media_dias_total > 0 else "N/A")

st.markdown("---")

row1_col1, row1_col2 = st.columns(2)

with row1_col1:
    st.subheader("1 Doc. por Status")
    status_alvo = ["APROVADO", "EM VERF INTERNA", "AGUARD_DEV_DO_SETOR", "CANCELADO"]
    df_g1_filtrado = df[df["STATUS DO DOCUMENTO NORMATIVO"].isin(status_alvo)]
    df_g1 = df_g1_filtrado["STATUS DO DOCUMENTO NORMATIVO"].value_counts().reset_index()
    df_g1.columns = ["STATUS DO DOCUMENTO NORMATIVO", "Quantidade"]
    
    if not df_g1.empty:
        fig1 = px.pie(df_g1, names="STATUS DO DOCUMENTO NORMATIVO", values="Quantidade", hole=0.4, color="STATUS DO DOCUMENTO NORMATIVO", color_discrete_map=mapa_cores_status)
        fig1.update_traces(textinfo='value+label', textposition='inside')
        fig1.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
        st.plotly_chart(fig1, use_container_width=True)

with row1_col2:
    st.subheader("2 Doc. Aprovados por Tipo")
    df_g2_filtrado = df[df["STATUS DO DOCUMENTO NORMATIVO"] == "APROVADO"]
    df_g2 = df_g2_filtrado["SIGLA DO DOCUMENTO"].value_counts().reset_index()
    df_g2.columns = ["Tipo de Documento", "Quantidade Aprovada"]
    
    if not df_g2.empty:
        fig2 = px.bar(df_g2, x="Quantidade Aprovada", y="Tipo de Documento", text="Quantidade Aprovada", orientation="h", color="Tipo de Documento", color_discrete_sequence=cor_sequencia)
        fig2.update_traces(textposition="outside")
        fig2.update_layout(margin=dict(l=20, r=20, t=30, b=20), showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")

# ==================== BLOCO DO GRÁFICO 3 MANTIDO FUNCIONANDO ====================
st.subheader("3 Validade por Tipo de Documentos")
col_real_g3 = None
for col in df.columns:
    if "VENCIDO" in col or "VALIDADE" in col:
        col_real_g3 = col
        break

if col_real_g3:
    df_g3_base = df.dropna(subset=[col_real_g3]).copy()
    df_g3_base[col_real_g3] = df_g3_base[col_real_g3].astype(str).str.upper().str.strip()
    df_g3_base[col_real_g3] = df_g3_base[col_real_g3].replace({
        "VENCIDO": "Vencidos",
        "VÁLIDO": "Válidos",
        "NO PRAZO": "No Prazo",
        "A": "Prestes a Vencer",
        "PRESTES A VENCER": "Prestes a Vencer"
    })
    
    status_validade_disponiveis = ["Válidos", "Vencidos", "No Prazo", "Prestes a Vencer"]
    validade_selecionada = st.multiselect("Filtrar Status de Validade:", options=status_validade_disponiveis, default=status_validade_disponiveis)
    
    df_g3_filtrado_val = df_g3_base[df_g3_base[col_real_g3].isin(validade_selecionada)]
    
    if not df_g3_filtrado_val.empty:
        df_g3_counts = df_g3_filtrado_val.groupby(["SIGLA DO DOCUMENTO", col_real_g3]).size().reset_index(name="Quantidade")
        
        fig3 = px.bar(df_g3_counts, 
                      x="SIGLA DO DOCUMENTO", 
                      y="Quantidade", 
                      color=col_real_g3, 
                      text="Quantidade", 
                      barmode="group", 
                      labels={"SIGLA DO DOCUMENTO": "Tipo de Documento", col_real_g3: "Status de Validade"},
                      color_discrete_sequence=cor_sequencia)
        
        fig3.update_traces(textposition="outside")
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("Nenhum dado encontrado para os status de validade selecionados.")
else:
    st.warning("Coluna de Validade não encontrada no arquivo carregado.")
# ======================================================================

st.markdown("---")

df_prof = df.copy()
if "RESPONSÁVEL" in df_prof.columns:
    df_prof["RESPONSÁVEL"] = df_prof["RESPONSÁVEL"].replace({"SONALIA": "OUTROS", "SABRINA": "OUTROS", "SONALHYA": "OUTROS"})

if "RESPONSÁVEL" in df_prof.columns:
    df_prof["RESPONSÁVEL_COMPARA"] = df_prof["RESPONSÁVEL"].apply(normalizar_texto)

st.subheader("4 Documentos por profissional")
if "RESPONSÁVEL" in df_prof.columns:
    profissionais_lista = sorted([p for p in df_prof["RESPONSÁVEL"].dropna().unique().tolist() if p != "OUTROS"])
    prof_selecionado_g4 = st.selectbox("Filtrar por profissional para o Gráfico 4:", options=["Todos"] + profissionais_lista)
    
    prof_sel_g4_limpo = normalizar_texto(prof_selecionado_g4)
    
    df_g4 = df_prof.copy() if prof_selecionado_g4 == "Todos" else df_prof[df_prof["RESPONSÁVEL_COMPARA"] == prof_sel_g4_limpo]
    df_g4_counts = df_g4.groupby(["RESPONSÁVEL", "STATUS DO DOCUMENTO NORMATIVO"]).size().reset_index(name="Quantidade")
    
    if not df_g4_counts.empty:
        fig4 = px.bar(df_g4_counts, x="Quantidade", y="RESPONSÁVEL", color="STATUS DO DOCUMENTO NORMATIVO", barmode="group", orientation="h", height=450, text="Quantidade", labels={"RESPONSÁVEL": "Profissional", "Quantidade": "N° de Documentos"}, color_discrete_map=mapa_cores_status)
        fig4.update_traces(textposition="outside")
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("Nenhum dado de profissional encontrado para gerar o Gráfico 4.")

st.markdown("---")

# ==================== BLOCO DO GRÁFICO 5 REMAPEADO E SEGURO ====================
st.subheader("5 Documentos por Tipo / Profissional")
if "RESPONSÁVEL" in df_prof.columns:
    prof_selecionado_g5 = st.selectbox("Selecione o Responsável para Filtrar a Análise Cruzada:", options=profissionais_lista, key="sb_grafico_5_novo")
    
    prof_sel_g5_limpo = normalizar_texto(prof_selecionado_g5)
    
