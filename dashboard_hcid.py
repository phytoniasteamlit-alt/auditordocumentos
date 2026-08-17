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

# Lógica robusta e inteligente para encontrar as colunas de forma automática
if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        abas_disponiveis = excel_file.sheet_names
        
        # Localiza de forma flexível qual aba contém os dados do HCID
        aba_hcid_real = None
        for opcao in ["HCID_BDD", "HCID", "HCID1", "DADOS"]:
            if opcao in abas_disponiveis:
                aba_hcid_real = opcao
                break
        if not aba_hcid_real:
            aba_hcid_real = abas_disponiveis[0]
                
        # Localiza de forma flexível a aba de ANEXO
        aba_anexo_real = None
        for opcao in ["ANEXO", "ANEXO2", "ANEXOS"]:
            if opcao in abas_disponiveis:
                aba_anexo_real = opcao
                break
                
        df_hcid = pd.read_excel(uploaded_file, sheet_name=aba_hcid_real)
        
        if aba_anexo_real and aba_anexo_real in abas_disponiveis:
            df_anexo = pd.read_excel(uploaded_file, sheet_name=aba_anexo_real)
        else:
            df_anexo = pd.DataFrame(columns=df_hcid.columns)
            
        # Função interna de varredura inteligente para mapear as colunas de forma definitiva
        def processar_mapeamento_inteligente(df_aba):
            if df_aba.empty:
                return df_aba, "SETOR", "SUB_SETOR", "CATEGORIA PROFISSIONAL", "TURNO", "VAGAS"
            
            # Limpa espaços em branco dos títulos das colunas
            df_aba.columns = [str(c).strip() for c in df_aba.columns]
            
            c_setor, c_sub, c_cat, c_turno, c_vagas = None, None, None, None, None
            
            # 1. Encontra a coluna de Vagas localizando qual delas é numérica de verdade
            for col in df_aba.columns:
                v_num = pd.to_numeric(df_aba[col], errors='coerce').dropna()
                if len(v_num) > 0 and v_num.sum() > len(v_num): 
                    c_vagas = col
                    break
            
            # 2. Se a busca numérica falhar, procura por aproximação de texto
            if not c_vagas:
                for col in df_aba.columns:
                    if "VAGA" in col.upper() or "TOTAL" in col.upper() or "QTD" in col.upper():
                        c_vagas = col
                        break
            
            # 3. Mapeia o resto das colunas de texto por termos aproximados
            for col in df_aba.columns:
                if col == c_vagas:
                    continue
                col_upper = col.upper()
                if "SUB" in col_upper: c_sub = col
                elif "SETOR" in col_upper or "CAMPO" in col_upper: c_setor = col
                elif "PROF" in col_upper or "CAT" in col_upper: c_cat = col
                elif "TURN" in col_upper: c_turno = col
            
            # Define nomes padrão de segurança caso a planilha venha sem algum título
            c_setor = c_setor if c_setor else (df_aba.columns[0] if len(df_aba.columns) > 0 else "SETOR")
            c_sub = c_sub if c_sub else (df_aba.columns[1] if len(df_aba.columns) > 1 else "SUB_SETOR")
            c_cat = c_cat if c_cat else (df_aba.columns[2] if len(df_aba.columns) > 2 else "CATEGORIA PROFISSIONAL")
            c_turno = c_turno if c_turno else (df_aba.columns[3] if len(df_aba.columns) > 3 else "TURNO")
            c_vagas = c_vagas if c_vagas else df_aba.columns[-1]
            
            # Limpa os textos internos das células
            for col in [c_setor, c_sub, c_cat, c_turno]:
                if col in df_aba.columns and df_aba[col].dtype == "object":
                    df_aba[col] = df_aba[col].astype(str).str.strip()
                    
            # Converte as vagas para números inteiros limpos
            df_aba[c_vagas] = pd.to_numeric(df_aba[c_vagas], errors='coerce').fillna(0).astype(int)
            return df_aba, c_setor, c_sub, c_cat, c_turno, c_vagas

        df_hcid, hc_setor, hc_sub, hc_cat, hc_turno, hc_vagas = processar_mapeamento_inteligente(df_hcid)
        df_anexo, ax_setor, ax_sub, ax_cat, ax_turno, ax_vagas = processar_mapeamento_inteligente(df_anexo)
        
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
