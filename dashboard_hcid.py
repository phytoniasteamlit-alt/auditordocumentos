import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (Modo Amplo)
# ==============================================================================
st.set_page_config(
    page_title="Painel de Estágios - HCID & ANEXO",
    layout="wide",
    initial_sidebar_state="expanded"
)

def normalizar_texto(texto):
    if pd.isna(texto) or not isinstance(texto, str):
        return ""
    texto = texto.strip().upper()
    texto = "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto

# --- CABEÇALHO SUPERIOR ---
header_left, header_right = st.columns(2)

with header_left:
    st.markdown("<h1 style='margin: 0; padding: 0; font-size: 2.2rem;'>📊 Painel de Indicadores de Estágio</h1>", unsafe_allow_html=True)

with header_right:
    st.markdown(
        """
        <div style="text-align: right; line-height: 1.2; padding-bottom: 10px;">
            <span style="font-size: 16px; font-weight: bold;">🏥 Hospital da Cidade</span><br>
            <span style="font-size: 14px; color: #888;">👩‍💼 Coord: Verônica Azevedo</span>
        </div>
        """, 
        unsafe_allow_html=True
    )

st.markdown("---")

# ==============================================================================
# 2. PAINEL DE CONTROLE (SIDEBAR) & CARREGAMENTO DAS ABAS
# ==============================================================================
st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha de Estágios (.xlsx):", type=["xlsx"])

st.sidebar.markdown("---")
st.sidebar.subheader("🎨 Customização Visual")
paleta_selecionada = st.sidebar.selectbox(
    "Tema de Cores Geral:",
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

tipo_grafico_5 = st.sidebar.radio("Estilo dos Gráficos de Setor:", options=["Barras Verticais", "Barras Horizontais"], index=1)

if uploaded_file is not None:
    try:
        df_hcid = pd.read_excel(uploaded_file, sheet_name="HCID")
        df_anexo = pd.read_excel(uploaded_file, sheet_name="ANEXO")
        
        for df_aba in [df_hcid, df_anexo]:
            for col in df_aba.columns:
                df_aba.rename(columns={col: col.strip().upper()}, inplace=True)
            for col in df_aba.columns:
                if df_aba[col].dtype == "object":
                    df_aba[col] = df_aba[col].astype(str).str.strip()
                    
        df_hcid["VAGAS"] = pd.to_numeric(df_hcid["VAGAS"], errors='coerce').fillna(0).astype(int)
        df_anexo["VAGAS"] = pd.to_numeric(df_anexo["VAGAS"], errors='coerce').fillna(0).astype(int)
        
    except Exception as e:
        st.error(f"Erro ao processar as abas 'HCID' ou 'ANEXO'. Verifique se o arquivo final contém apenas esses dois nomes de aba. Erro: {e}")
        st.stop()
else:
    st.info("💡 Por favor, use o menu lateral para carregar a sua planilha Excel estruturada na vertical.")
    st.stop()

# ==============================================================================
# 3. BLOCO 1: GRÁFICOS DO HCID
# ==============================================================================
st.markdown("<h2 style='color: #2ca02c;'>🏢 Indicadores Exclusivos - HCID</h2>", unsafe_allow_html=True)
st.markdown("---")

categorias_hcid = sorted(df_hcid["CATEGORIA PROFISSIONAL"].unique().tolist())
filtro_cat_hcid = st.sidebar.multiselect("Filtrar Profissões (HCID):", options=categorias_hcid, default=categorias_hcid)
df_hcid_filtrado = df_hcid[df_hcid["CATEGORIA PROFISSIONAL"].isin(filtro_cat_hcid)]

r1_c1, r1_col2 = st.columns(2)
with r1_c1:
    st.subheader("1. Total de Vagas de Estágio no HCID")
    st.metric(label="Vagas Totais Disponibilizadas", value=df_hcid_filtrado["VAGAS"].sum())

with r1_col2:
    st.subheader("2. Total de Setores Disponibilizados p/ Estágio no HCID")
    st.metric(label="Setores com Campos Ativos", value=df_hcid_filtrado["SETOR"].nunique())

st.markdown("---")

r2_c1, r2_c2 = st.columns(2)
with r2_c1:
    st.subheader("3. Setores Disponibilizados para Realização de Estágio no HCID")
    df_g3 = df_hcid_filtrado.groupby("SETOR")["VAGAS"].sum().reset_index()
    ori_3 = "h" if tipo_grafico_5 == "Barras Horizontais" else "v"
    x_v, y_v = ("VAGAS", "SETOR") if ori_3 == "h" else ("SETOR", "VAGAS")
    fig3 = px.bar(df_g3, x=x_v, y=y_v, text="VAGAS", orientation=ori_3, color="SETOR", color_discrete_sequence=cor_sequencia)
    fig3.update_traces(textposition="outside", textfont=dict(size=14))
    fig3.update_layout(showlegend=False, height=500)
    st.plotly_chart(fig3, use_container_width=True)

with r2_c2:
    st.subheader("4. Categorias Profissionais Contempladas no Estágio por Setor no HCID")
    df_g4 = df_hcid_filtrado.groupby(["SETOR", "CATEGORIA PROFISSIONAL"])["VAGAS"].sum().reset_index()
    ori_4 = "h" if tipo_grafico_5 == "Barras Horizontais" else "v"
    x_v4, y_v4 = ("VAGAS", "SETOR") if ori_4 == "h" else ("SETOR", "VAGAS")
    fig4 = px.bar(df_g4, x=x_v4, y=y_v4, color="CATEGORIA PROFISSIONAL", orientation=ori_4, barmode="stack", color_discrete_sequence=cor_sequencia)
    fig4.update_layout(height=500)
    st.plotly_chart(fig4, use_container_width=True)

st.markdown("---")

r3_c1, r3_c2, r3_c3 = st.columns(3)
with r3_c1:
    st.subheader("5. Total de Vagas por Sub-Setor no HCID")
    df_g5 = df_hcid_filtrado.groupby("SUB_SETOR")["VAGAS"].sum().reset_index()
    fig5 = px.bar(df_g5, x="VAGAS", y="SUB_SETOR", orientation="h", color_discrete_sequence=cor_sequencia)
    fig5.update_layout(height=450)
    st.plotly_chart(fig5, use_container_width=True)

with r3_c2:
    st.subheader("6. Total de Vagas do HCID por Turno")
    df_g6 = df_hcid_filtrado.groupby("TURNO")["VAGAS"].sum().reset_index()
    fig6 = px.bar(df_g6, x="TURNO", y="VAGAS", text="VAGAS", color="TURNO", color_discrete_sequence=px.colors.qualitative.Pastel)
    fig6.update_traces(textposition="outside", textfont=dict(size=15))
    fig6.update_layout(showlegend=False, height=450)
    st.plotly_chart(fig6, use_container_width=True)

with r3_c3:
    st.subheader("7. Total de Estagiários por Turno por Setor Geral no HCID")
    df_g7 = df_hcid_filtrado.groupby(["SETOR", "TURNO"])["VAGAS"].sum().reset_index()
    fig7 = px.bar(df_g7, x="SETOR", y="VAGAS", color="TURNO", barmode="group", color_discrete_sequence=px.colors.qualitative.Safe)
    fig7.update_layout(height=450, xaxis_tickangle=-45)
    st.plotly_chart(fig7, use_container_width=True)

# ==============================================================================
# 4. BLOCO 2: GRÁFICOS DO ANEXO
# ==============================================================================
st.markdown("<br><br>---", unsafe_allow_html=True)
st.markdown("<h2 style='color: #d62728;'>🏢 Indicadores Exclusivos - ANEXO</h2>", unsafe_allow_html=True)
st.markdown("---")

categorias_anexo = sorted(df_anexo["CATEGORIA PROFISSIONAL"].unique().tolist()) if not df_anexo.empty else []
filtro_cat_anexo = st.sidebar.multiselect("Filtrar Profissões (Anexo):", options=categorias_anexo, default=categorias_anexo)
df_anexo_filtrado = df_anexo[df_anexo["CATEGORIA PROFISSIONAL"].isin(filtro_cat_anexo)] if not df_anexo.empty else df_anexo

ax_r1_c1, ax_r1_col2 = st.columns(2)
with ax_r1_c1:
    st.subheader("1. Total de Vagas de Estágio no Anexo")
    st.metric(label="Vagas Totais (Anexo)", value=df_anexo_filtrado["VAGAS"].sum() if not df_anexo_filtrado.empty else 0)

with ax_r1_col2:
    st.subheader("2. Total de Setores Disponibilizados por Campo de Estágio no Anexo")
    st.metric(label="Setores Ativos (Anexo)", value=df_anexo_filtrado["SETOR"].nunique() if not df_anexo_filtrado.empty else 0)

st.markdown("---")

ax_r2_c1, ax_r2_c2 = st.columns(2)
with ax_r2_c1:
    st.subheader("3. Setores Disponibilizados para Realização de Estágio no Anexo")
    if not df_anexo_filtrado.empty:
        df_ax3 = df_anexo_filtrado.groupby("SETOR")["VAGAS"].sum().reset_index()
        fig_ax3 = px.bar(df_ax3, x=x_v, y=y_v, text="VAGAS", orientation=ori_3, color="SETOR", color_discrete_sequence=cor_sequencia)
        fig_ax3.update_traces(textposition="outside", textfont=dict(size=14))
        fig_ax3.update_layout(showlegend=False, height=500)
        st.plotly_chart(fig_ax3, use_container_width=True)

with ax_r2_c2:
    st.subheader("4. Categorias Profissionais Contempladas no Estágio por Setor no Anexo")
    if not df_anexo_filtrado.empty:
        df_ax4 = df_anexo_filtrado.groupby(["SETOR", "CATEGORIA PROFISSIONAL"])["VAGAS"].sum().reset_index()
        fig_ax4 = px.bar(df_ax4, x=x_v4, y=y_v4, color="CATEGORIA PROFISSIONAL", orientation=ori_4, barmode="stack", color_discrete_sequence=cor_sequencia)
        fig_ax4.update_layout(height=500)
        st.plotly_chart(fig_ax4, use_container_width=True)

# ==============================================================================
# 5. RODAPÉ - ASSINATURA DO PROGRAMADOR
# ==============================================================================
st.markdown("<br><br><br>", unsafe_allow_html=True)
col_assinatura, _ = st.columns(2)
with col_assinatura:
    st.markdown(
        """
        <div style="padding: 10px; border-top: 1px solid #ddd; color: #555; font-size: 0.9rem;">
            👨‍💻 <b>Programador:</b> Ezequias S. Santos Naqh / Nsp
        </div>
        """,
        unsafe_allow_html=True
    )
