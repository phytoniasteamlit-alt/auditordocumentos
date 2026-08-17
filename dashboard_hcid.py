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

# Lógica robusta de leitura de planilhas para evitar travar com nomes de abas ou colunas diferentes
if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        abas_disponiveis = excel_file.sheet_names
        
        # Procura de forma inteligente a aba do HCID (aceita HCID, HCID_BDD, HCID1...)
        aba_hcid_real = "HCID"
        for aba in abas_disponiveis:
            if "HCID" in aba.upper():
                aba_hcid_real = aba
                break
                
        # Procura de forma inteligente a aba do ANEXO
        aba_anexo_real = "ANEXO"
        for aba in abas_disponiveis:
            if "ANEXO" in aba.upper():
                aba_anexo_real = aba
                break
        
        df_hcid = pd.read_excel(uploaded_file, sheet_name=aba_hcid_real)
        df_anexo = pd.read_excel(uploaded_file, sheet_name=aba_anexo_real if aba_anexo_real in abas_disponiveis else abas_disponiveis[-1])
        
        # Padronização e limpeza forçada de cabeçalhos
        for df_aba in [df_hcid, df_anexo]:
            for col in df_aba.columns:
                df_aba.rename(columns={col: str(col).strip().upper()}, inplace=True)
            for col in df_aba.columns:
                if df_aba[col].dtype == "object":
                    df_aba[col] = df_aba[col].astype(str).str.strip()
                    
        # Mapeamento dinâmico de colunas para blindar o erro 'VAGAS'
        def descobrir_colunas(df_tratar):
            c_setor, c_sub, c_cat, c_turno, c_vagas = "SETOR", "SUB_SETOR", "CATEGORIA PROFISSIONAL", "TURNO", "VAGAS"
            for col in df_tratar.columns:
                if "SUB" in col: c_sub = col
                elif "SETOR" in col or "CAMPO" in col: c_setor = col
                elif "PROF" in col or "CAT" in col: c_cat = col
                elif "TURN" in col: c_turno = col
                elif "VAGA" in col or "TOTAL" in col or "QTD" in col: c_vagas = col
            return c_setor, c_sub, c_cat, c_turno, c_vagas

        hc_setor, hc_sub, hc_cat, hc_turno, hc_vagas = descobrir_colunas(df_hcid)
        ax_setor, ax_sub, ax_cat, ax_turno, ax_vagas = descobrir_colunas(df_anexo)
        
        # Conversão numérica limpa e segura
        df_hcid[hc_vagas] = pd.to_numeric(df_hcid[hc_vagas], errors='coerce').fillna(0).astype(int)
        df_anexo[ax_vagas] = pd.to_numeric(df_anexo[ax_vagas], errors='coerce').fillna(0).astype(int)
        
    except Exception as e:
        st.error(f"Erro crítico no mapeamento das colunas da planilha. Detalhes: {e}")
        st.stop()
else:
    st.info("💡 Por favor, use o menu lateral para carregar a sua planilha Excel estruturada na vertical.")
    st.stop()

# ==============================================================================
# 3. BLOCO 1: GRÁFICOS DO HCID
# ==============================================================================
st.markdown("<h2 style='color: #2ca02c;'>🏢 Indicadores Exclusivos - HCID</h2>", unsafe_allow_html=True)
st.markdown("---")

if hc_cat in df_hcid.columns:
    categorias_hcid = sorted(df_hcid[hc_cat].dropna().unique().tolist())
    filtro_cat_hcid = st.sidebar.multiselect("Filtrar Profissões (HCID):", options=categorias_hcid, default=categorias_hcid)
    df_hcid_filtrado = df_hcid[df_hcid[hc_cat].isin(filtro_cat_hcid)]
else:
    df_hcid_filtrado = df_hcid

r1_c1, r1_col2 = st.columns(2)
with r1_c1:
    st.subheader("1. Total de Vagas de Estágio no HCID")
    st.metric(label="Vagas Totais Disponibilizadas", value=df_hcid_filtrado[hc_vagas].sum())

with r1_col2:
    st.subheader("2. Total de Setores Disponibilizados p/ Estágio no HCID")
    st.metric(label="Setores com Campos Ativos", value=df_hcid_filtrado[hc_setor].nunique())

st.markdown("---")

r2_c1, r2_c2 = st.columns(2)
with r2_c1:
    st.subheader("3. Setores Disponibilizados para Realização de Estágio no HCID")
    df_g3 = df_hcid_filtrado.groupby(hc_setor)[hc_vagas].sum().reset_index()
    ori_3 = "h" if tipo_grafico_5 == "Barras Horizontais" else "v"
    x_v, y_v = (hc_vagas, hc_setor) if ori_3 == "h" else (hc_setor, hc_vagas)
    fig3 = px.bar(df_g3, x=x_v, y=y_v, text=hc_vagas, orientation=ori_3, color=hc_setor, color_discrete_sequence=cor_sequencia)
    fig3.update_traces(textposition="outside", textfont=dict(size=14))
    fig3.update_layout(showlegend=False, height=500)
    st.plotly_chart(fig3, use_container_width=True)

with r2_c2:
    st.subheader("4. Categorias Profissionais Contempladas no Estágio por Setor no HCID")
    df_g4 = df_hcid_filtrado.groupby([hc_setor, hc_cat])[hc_vagas].sum().reset_index()
    ori_4 = "h" if tipo_grafico_5 == "Barras Horizontais" else "v"
    x_v4, y_v4 = (hc_vagas, hc_setor) if ori_4 == "h" else (hc_setor, hc_vagas)
    fig4 = px.bar(df_g4, x=x_v4, y=y_v4, color=hc_cat, orientation=ori_4, barmode="stack", color_discrete_sequence=cor_sequencia)
    fig4.update_layout(height=500, legend=dict(title_text="Profissão"))
    st.plotly_chart(fig4, use_container_width=True)

st.markdown("---")

r3_c1, r3_c2, r3_c3 = st.columns(3)
with r3_c1:
    st.subheader("5. Total de Vagas por Sub-Setor no HCID")
    df_g5 = df_hcid_filtrado.groupby(hc_sub)[hc_vagas].sum().reset_index()
    fig5 = px.bar(df_g5, x=hc_vagas, y=hc_sub, orientation="h", color_discrete_sequence=cor_sequencia)
    fig5.update_layout(height=450)
    st.plotly_chart(fig5, use_container_width=True)

with r3_c2:
    st.subheader("6. Total de Vagas do HCID por Turno")
    df_g6 = df_hcid_filtrado.groupby(hc_turno)[hc_vagas].sum().reset_index()
    fig6 = px.bar(df_g6, x=hc_turno, y=hc_vagas, text=hc_vagas, color=hc_turno, color_discrete_sequence=px.colors.qualitative.Pastel)
    fig6.update_traces(textposition="outside", textfont=dict(size=15))
    fig6.update_layout(showlegend=False, height=450)
    st.plotly_chart(fig6, use_container_width=True)

with r3_c3:
    st.subheader("7. Total de Estagiários por Turno por Setor Geral no HCID")
    df_g7 = df_hcid_filtrado.groupby([hc_setor, hc_turno])[hc_vagas].sum().reset_index()
    fig7 = px.bar(df_g7, x=hc_setor, y=hc_vagas, color=hc_turno, barmode="group", color_discrete_sequence=px.colors.qualitative.Safe)
    fig7.update_layout(height=450, xaxis_tickangle=-45)
    st.plotly_chart(fig7, use_container_width=True)

# ==============================================================================
# 4. BLOCO 2: GRÁFICOS DO ANEXO
# ==============================================================================
st.markdown("<br><br>---", unsafe_allow_html=True)
st.markdown("<h2 style='color: #d62728;'>🏢 Indicadores Exclusivos - ANEXO</h2>", unsafe_allow_html=True)
st.markdown("---")

if ax_cat in df_anexo.columns and not df_anexo.empty:
    categorias_anexo = sorted(df_anexo[ax_cat].dropna().unique().tolist())
    filtro_cat_anexo = st.sidebar.multiselect("Filtrar Profissões (Anexo):", options=categorias_anexo, default=categorias_anexo)
    df_anexo_filtrado = df_anexo[df_anexo[ax_cat].isin(filtro_cat_anexo)]
else:
    df_anexo_filtrado = df_anexo

ax_r1_c1, ax_r1_col2 = st.columns(2)
with ax_r1_c1:
    st.subheader("1. Total de Vagas de Estágio no Anexo")
    st.metric(label="Vagas Totais (Anexo)", value=df_anexo_filtrado[ax_vagas].sum() if not df_anexo_filtrado.empty else 0)

with ax_r1_col2:
    st.subheader("2. Total de Setores Disponibilizados por Campo de Estágio no Anexo")
    st.metric(label="Setores Ativos (Anexo)", value=df_anexo_filtrado[ax_setor].nunique() if not df_anexo_filtrado.empty else 0)

st.markdown("---")

ax_r2_c1, ax_r2_c2 = st.columns(2)
with ax_r2_c1:
