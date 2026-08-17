import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (Modo Amplo e Executivo)
# ==============================================================================
st.set_page_config(
    page_title="Painel Executivo de Estágios - HCID",
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
        <span style="font-size: 16px; font-weight: bold;">🏥 Hospital da Cidade Dr. Jackson Lago</span><br>
        <span style="font-size: 14px; color: #888;">👩‍💼 Coordenação: Verônica Azevedo</span>
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
    index=0
)

if paleta_selecionada == "Tons Pastéis":
    cor_sequencia = px.colors.qualitative.Pastel
elif paleta_selecionada == "Vibrante":
    cor_sequencia = px.colors.qualitative.Prism
elif paleta_selecionada == "Esmeralda":
    cor_sequencia = px.colors.sequential.Mint
else:
    cor_sequencia = ["#008080", "#4682B4", "#20B2AA", "#5F9EA0", "#B0C4DE"]

# Lógica de leitura de abas
if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        abas_disponiveis = excel_file.sheet_names
        
        aba_hcid_real = None
        for opcao in ["HCID_BDD", "HCID", "HCID1", "DADOS"]:
            if opcao in abas_disponiveis:
                aba_hcid_real = opcao
                break
        if not aba_hcid_real:
            aba_hcid_real = abas_disponiveis if abas_disponiveis else None
                
        # AJUSTE CRÍTICO: Pula as 3 primeiras linhas (linha 4 do Excel vira o cabeçalho real)
        df_hcid_bruto = pd.read_excel(uploaded_file, sheet_name=aba_hcid_real, skiprows=3)
        
        def processar_mapeamento_executivo(df_aba):
            if df_aba.empty:
                return pd.DataFrame(), "SETOR", "SUB_SETOR", "CATEGORIA", "VAGAS_MANHA", "VAGAS_TARDE", "VAGAS_TOTAL"
            
            # Limpeza e padronização dos nomes das colunas
            df_aba.columns = [str(c).strip().replace('\n', ' ') for c in df_aba.columns]
            
            c_setor, c_sub, c_cat = None, None, None
            c_manha, c_tarde = None, None
            
            # Mapeamento inteligente baseado no seu layout real da imagem
            for col in df_aba.columns:
                col_upper = col.upper()
                if "SUB" in col_upper: c_sub = col
                elif "SETOR" in col_upper or "CAMPO" in col_upper: c_setor = col
                elif "PROF" in col_upper or "CAT" in col_upper: c_cat = col
                elif "MANH" in col_upper: c_manha = col
                elif "TARD" in col_upper: c_tarde = col
            
            # Fallbacks seguros caso os nomes variem de leve
            c_setor = c_setor or "SETOR"
            c_sub = c_sub or "SUB_SETOR"
            c_cat = c_cat or "CATEGORIA PROFISSIONAL"
            
            # Garante a existência das colunas de texto para evitar erros
            if c_setor not in df_aba.columns: df_aba[c_setor] = "NÃO INFORMADO"
            if c_sub not in df_aba.columns: df_aba[c_sub] = ""
            if c_cat not in df_aba.columns: df_aba[c_cat] = "NÃO INFORMADO"
            
            # Preenche células mescladas para baixo (Mapeia PS, SADT, etc para as linhas filhas)
            df_aba[c_setor] = df_aba[c_setor].ffill()
            df_aba[c_sub] = df_aba[c_sub].fillna("")
            df_aba[c_cat] = df_aba[c_cat].ffill()
            
            # Limpa o texto "4 por turno" isolando o número 4
            def limpar_vagas(valor):
                if pd.isna(valor):
                    return 0
                v_str = "".join(filter(str.isdigit, str(valor)))
                return int(v_str) if v_str != "" else 0
            
            df_aba["VAGAS_MANHA"] = df_aba[c_manha].apply(limpar_vagas) if c_manha else 0
            df_aba["VAGAS_TARDE"] = df_aba[c_tarde].apply(limpar_vagas) if c_tarde else 0
            df_aba["VAGAS_TOTAL"] = df_aba["VAGAS_MANHA"] + df_aba["VAGAS_TARDE"]
            
            # Concatena Setor + Subsetor de forma elegante
            df_aba["LOCAL_COMBINADO"] = df_aba.apply(
                lambda r: f"{r[c_setor]}" if (not r[c_sub] or str(r[c_setor]).upper() == str(r[c_sub]).upper()) else f"{r[c_setor]} - {r[c_sub]}",
                axis=1
            )
            
            # Elimina linhas em branco e subtotais antigos da planilha
            df_aba = df_aba[
                (~df_aba[c_setor].astype(str).str.upper().str.contains("TOTAL|QUANTITATIVO|HOSPITAL", na=False)) & 
                (df_aba["VAGAS_TOTAL"] > 0)
            ]
            
            return df_aba, c_setor, c_sub, c_cat
            
        df_hcid, hc_setor, hc_sub, hc_cat = processar_mapeamento_executivo(df_hcid_bruto)
        
    except Exception as e:
        st.error(f"Erro ao processar estrutura da aba HCID_BDD: {e}")
        st.stop()
else:
    st.info("💡 Por favor, use o menu lateral para carregar a sua planilha Excel estruturada na vertical (HCID_BDD).")
    st.stop()

# ==============================================================================
# 3. APRESENTAÇÃO DE INDICADORES (FOCO EXECUTIVO)
# ==============================================================================
st.markdown("<h2 style='color: #008080;'>🏢 Análise Estratégica de Capacidade - HCID</h2>", unsafe_allow_html=True)
st.markdown("---")

# Filtros com verificação de segurança para não carregar vazios
if hc_cat in df_hcid.columns and not df_hcid.empty:
    categorias_disponiveis = sorted([str(x) for x in df_hcid[hc_cat].dropna().unique() if str(x).strip() != ""])
    filtro_cat = st.sidebar.multiselect("Filtrar por Categoria Profissional:", options=categorias_disponiveis, default=categorias_disponiveis)
    df_filtrado = df_hcid[df_hcid[hc_cat].isin(filtro_cat)]
else:
    df_filtrado = df_hcid

# --- LINHA 1: MÉTRICAS DE IMPACTO IMEDIATO ---
if not df_filtrado.empty:
    m1, m2, m3, m4 = st.columns(4)
    
    total_geral = int(df_filtrado["VAGAS_TOTAL"].sum())
    total_manha = int(df_filtrado["VAGAS_MANHA"].sum())
    total_tarde = int(df_filtrado["VAGAS_TARDE"].sum())
    total_setores = df_filtrado["LOCAL_COMBINADO"].nunique()
    
    m1.metric(label="🎯 Capacidade Total de Vagas", value=f"{total_geral} Vagas")
    m2.metric(label="🌅 Alocação Período Manhã", value=f"{total_manha} Estagiários")
    m3.metric(label="🌇 Alocação Período Tarde", value=f"{total_tarde} Estagiários")
    m4.metric(label="📍 Setores/Subsetores Atendidos", value=total_setores)
    
    st.markdown("---")
    
    # --- LINHA 2: GRÁFICOS ---
    g1, g2 = st.columns(2)
    
    with g1:
        st.markdown("#### 📊 Distribuição de Vagas por Setor e Turno")
        df_grafico = df_filtrado.groupby("LOCAL_COMBINADO")[["VAGAS_MANHA", "VAGAS_TARDE"]].sum().reset_index()
        df_longo = df_grafico.melt(id_vars="LOCAL_COMBINADO", value_vars=["VAGAS_MANHA", "VAGAS_TARDE"], 
                                  var_name="TURNO", value_name="VAGAS")
        df_longo["TURNO"] = df_longo["TURNO"].map({"VAGAS_MANHA": "Manhã", "VAGAS_TARDE": "Tarde"})
        
        fig_barra = px.bar(
            df_longo,
            x="VAGAS",
            y="LOCAL_COMBINADO",
            color="TURNO",
            orientation="h",
            color_discrete_map={"Manhã": "#4682B4", "Tarde": "#FF8C00"},
            text="VAGAS"
        )
        fig_barra.update_layout(barmode="stack", height=500, yaxis={'categoryorder':'total ascending'}, legend_title_text="Turno")
        fig_barra.update_traces(textposition="inside")
        st.plotly_chart(fig_barra, use_container_width=True)
    
    with g2:
        st.markdown("#### 🎯 Ocupação por Categoria Profissional")
        df_pizza = df_filtrado.groupby(hc_cat)["VAGAS_TOTAL"].sum().reset_index()
        fig_pizza = px.pie(
            df_pizza, 
            values="VAGAS_TOTAL", 
            names=hc_cat, 
            hole=0.4,
            color_discrete_sequence=cor_sequencia
        )
        fig_pizza.update_traces(textinfo="percent+value")
        fig_pizza.update_layout(height=500, legend=dict(orientation="h", y=-0.1))
        st.plotly_chart(fig_pizza, use_container_width=True)
    
    st.markdown("---")
    
    # --- LINHA 3: TABELA DETALHADA ---
    st.markdown("### 📋 Resumo Executivo para Auditoria (Diretoria)")
    df_tabela_entrega = df_filtrado[[hc_setor, hc_sub, hc_cat, "VAGAS_MANHA", "VAGAS_TARDE", "VAGAS_TOTAL"]].rename(
        columns={
            hc_setor: "Setor Hospitalar",
            hc_sub: "Sub-Setor / Ala",
            hc_cat: "Categoria Profissional",
            "VAGAS_MANHA": "Vagas Manhã",
            "VAGAS_TARDE": "Vagas Tarde",
            "VAGAS_TOTAL": "Total de Vagas"
        }
    )
