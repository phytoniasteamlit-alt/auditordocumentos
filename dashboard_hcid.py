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

header_left.markdown("<h1 style='margin: 0; padding: 0; font-size: 2.2rem;'>📊 Painel de Indicadores de Estágio</h1>", unsafe_allow_html=True)

header_right.markdown(
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

# Lógica robusta de leitura inteligente de abas e colunas
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
            if df_anexo.dropna(how="all").empty:
                df_anexo = pd.DataFrame(columns=["SETOR", "SUB_SETOR", "CATEGORIA PROFISSIONAL", "TURNO", "VAGAS"])
        else:
            df_anexo = pd.DataFrame(columns=["SETOR", "SUB_SETOR", "CATEGORIA PROFISSIONAL", "TURNO", "VAGAS"])
            
        # Função interna de varredura inteligente corrigida contra erros de tipo 'Index'
        def processar_mapeamento_inteligente(df_aba):
            if df_aba.empty:
                df_vazia = pd.DataFrame(columns=["SETOR", "SUB_SETOR", "CATEGORIA PROFISSIONAL", "TURNO", "VAGAS"])
                return df_vazia, "SETOR", "SUB_SETOR", "CATEGORIA PROFISSIONAL", "TURNO", "VAGAS"
            
            df_aba.columns = [str(c).strip() for c in df_aba.columns]
            c_setor, c_sub, c_cat, c_turno, c_vagas = None, None, None, None, None
            
            # 1. Encontra a coluna de Vagas localizando qual delas é numérica de verdade
            for col in df_aba.columns:
                v_num = pd.to_numeric(df_aba[col], errors='coerce').dropna()
                if len(v_num) > 0 and v_num.sum() > len(v_num): 
                    c_vagas = col
                    break
            
            # 2. Se a busca numérica falhar, procura por texto
            if not c_vagas:
                for col in df_aba.columns:
                    if "VAGA" in col.upper() or "TOTAL" in col.upper() or "QTD" in col.upper():
                        c_vagas = col
                        break
            
            # 3. Mapeia o resto das colunas por termos aproximados
            for col in df_aba.columns:
                if col == c_vagas:
                    continue
                col_upper = col.upper()
                if "SUB" in col_upper: c_sub = col
                elif "SETOR" in col_upper or "CAMPO" in col_upper: c_setor = col
                elif "PROF" in col_upper or "CAT" in col_upper: c_cat = col
                elif "TURN" in col_upper: c_turno = col
            
            # Garante que as variáveis contenham strings puras em vez de listas de índices do pandas
            c_setor = str(c_setor) if c_setor else "SETOR"
            c_sub = str(c_sub) if c_sub else "SUB_SETOR"
            c_cat = str(c_cat) if c_cat else "CATEGORIA PROFISSIONAL"
            c_turno = str(c_turno) if c_turno else "TURNO"
            c_vagas = str(c_vagas) if c_vagas else "VAGAS"
            
            # Força a criação das colunas caso faltem na tabela carregada
            for col_nome in [c_setor, c_sub, c_cat, c_turno, c_vagas]:
                if col_nome not in df_aba.columns:
                    df_aba[col_nome] = "NÃO INFORMADO" if col_nome != c_vagas else 0
            
            for col in [c_setor, c_sub, c_cat, c_turno]:
                if df_aba[col].dtype == "object":
                    df_aba[col] = df_aba[col].astype(str).str.strip()
                    
            df_aba[c_vagas] = pd.to_numeric(df_aba[c_vagas], errors='coerce').fillna(0).astype(int)
            return df_aba, c_setor, c_sub, c_cat, c_turno, c_vagas

        df_hcid, hc_setor, hc_sub, hc_cat, hc_turno, hc_vagas = processar_mapeamento_inteligente(df_hcid)
        df_anexo, ax_setor, ax_sub, ax_cat, ax_turno, ax_vagas = processar_mapeamento_inteligente(df_anexo)
        
        # Cria a engenharia clean de Setor - Sub-setor separada por tabela de forma segura
        for d_f, s_t, s_b in [(df_hcid, hc_setor, hc_sub), (df_anexo, ax_setor, ax_sub)]:
            if not d_f.empty:
                d_f["LOCAL_COMBINADO"] = d_f.apply(
                    lambda r: f"{r[s_t]}" if str(r[s_t]).upper() == str(r[s_b]).upper() else f"{r[s_t]} - {r[s_b]}",
                    axis=1
                )
            else:
                d_f["LOCAL_COMBINADO"] = pd.Series(dtype=str)
        
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

if hc_cat in df_hcid.columns and not df_hcid.empty:
    categorias_hcid = sorted(df_hcid[hc_cat].dropna().unique().tolist())
    filtro_cat_hcid = st.sidebar.multiselect("Filtrar Profissões (HCID):", options=categorias_hcid, default=categorias_hcid)
    df_hcid_filtrado = df_hcid[df_hcid[hc_cat].isin(filtro_cat_hcid)]
else:
    df_hcid_filtrado = df_hcid

r1_c1, r1_col2 = st.columns(2)
if not df_hcid_filtrado.empty and "LOCAL_COMBINADO" in df_hcid_filtrado.columns:
    r1_c1.metric(label="Vagas de Estágio por Turno no HCID", value=int(df_hcid_filtrado.groupby(["LOCAL_COMBINADO", hc_turno])[hc_vagas].sum().max()))
    r1_col2.metric(label="Áreas de Estágio Ativas no HCID", value=df_hcid_filtrado["LOCAL_COMBINADO"].nunique())
else:
    r1_c1.metric(label="Vagas de Estágio por Turno no HCID", value=0)
    r1_col2.metric(label="Áreas de Estágio Ativas no HCID", value=0)

st.markdown("---")

r2_c1, r2_c2 = st.columns(2)

# Gráfico 3
if not df_hcid_filtrado.empty:
    df_g3 = df_hcid_filtrado.groupby("LOCAL_COMBINADO")[hc_vagas].mean().reset_index()
    df_g3[hc_vagas] = df_g3[hc_vagas].round(1)
    ori_3 = "h" if tipo_grafico_5 == "Barras Horizontais" else "v"
    x_v, y_v = (hc_vagas, "LOCAL_COMBINADO") if ori_3 == "h" else ("LOCAL_COMBINADO", hc_vagas)
    fig3 = px.bar(df_g3, x=x_v, y=y_v, text=hc_vagas, orientation=ori_3, color="LOCAL_COMBINADO", color_discrete_sequence=cor_sequencia, title="3. Vagas Disponibilizadas por Turno / Campo de Estágio no HCID")
    fig3.update_traces(textposition="outside", textfont=dict(size=14))
    fig3.update_layout(showlegend=False, height=650)
    r2_c1.plotly_chart(fig3, use_container_width=True)
else:
    r2_c1.info("Nenhum dado do HCID selecionado nos filtros laterais.")

# Gráfico 4
if not df_hcid_filtrado.empty:
    df_g4 = df_hcid_filtrado.groupby(["LOCAL_COMBINADO", hc_cat])[hc_vagas].mean().reset_index()
    df_g4[hc_vagas] = df_g4[hc_vagas].round(1)
    ori_4 = "h" if tipo_grafico_5 == "Barras Horizontais" else "v"
    x_v4, y_v4 = (hc_vagas, "LOCAL_COMBINADO") if ori_4 == "h" else ("LOCAL_COMBINADO", hc_vagas)
