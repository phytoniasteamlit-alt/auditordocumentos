import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (Interface Dashboard Executivo)
# ==============================================================================
st.set_page_config(
    page_title="Painel Geral de Estágios - HCID & ANEXOS",
    layout="wide",
    initial_sidebar_state="expanded"
)

def normalizar_texto(texto):
    if pd.isna(texto) or not isinstance(texto, str):
        return ""
    texto = texto.strip().upper()
    texto = "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto

# --- CABEÇALHO SUPERIOR (Igual ao Modelo Solicitado) ---
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
# 2. PAINEL DE CONTROLE (SIDEBAR) & CONFIGURAÇÃO VISUAL
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

# ==============================================================================
# 3. MOTOR DE PROCESSAMENTO DE DADOS (Mapeamento Inteligente por Aba)
# ==============================================================================
if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        abas_disponiveis = excel_file.sheet_names
        
        # Localização exata da aba HCID
        aba_hcid_real = None
        for opcao in ["HCID_BDD", "HCID", "HCID1", "DADOS"]:
            if opcao in abas_disponiveis:
                aba_hcid_real = opcao
                break
                
        # Localização exata da aba ANEXO
        aba_anexo_real = None
        for opcao in ["ANEXO", "ANEXO2", "ANEXOS"]:
            if opcao in abas_disponiveis:
                aba_anexo_real = opcao
                break
        
        # Função interna de tratamento e limpeza de dados
        def extrair_e_limpar_dados(sheet_name):
            if not sheet_name or sheet_name not in abas_disponiveis:
                return pd.DataFrame(), "SETOR", "SUB_SETOR", "CATEGORIA"
            
            df_bruto = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
            
            # Varredura para encontrar a linha real do cabeçalho
            linha_cabecalho = 0
            for idx, row in df_bruto.iterrows():
                row_str = " ".join([str(x).upper() for x in row.dropna()])
                if "CATEGOR" in row_str or "PROFISS" in row_str or "SETOR" in row_str:
                    linha_cabecalho = idx
                    break
            
            df_aba = df_bruto.iloc[linha_cabecalho+1:].copy()
            df_aba.columns = [str(c).strip().replace('\n', ' ') for c in df_bruto.iloc[linha_cabecalho]]
            
            # Tratamento de colunas nulas resultantes de células mescladas
            novas_colunas = []
            ultima_coluna_valida = "COLUNA"
            for col in df_aba.columns:
                if pd.isna(col) or "UNNAMED" in str(col).upper() or str(col).strip() == "":
                    novas_colunas.append(ultima_coluna_valida)
                else:
                    ultima_coluna_valida = str(col).strip()
                    novas_colunas.append(ultima_coluna_valida)
            df_aba.columns = novas_colunas
            
            c_setor, c_sub, c_cat = None, None, None
            for col in set(df_aba.columns):
                col_upper = col.upper()
                if "SUB" in col_upper: c_sub = col
                elif "SETOR" in col_upper or "CAMPO" in col_upper: c_setor = col
                elif "PROF" in col_upper or "CAT" in col_upper: c_cat = col
            
            c_setor = c_setor or df_aba.columns[0]
            c_sub = c_sub or (df_aba.columns[1] if len(df_aba.columns) > 1 else df_aba.columns[0])
            c_cat = c_cat or (df_aba.columns[2] if len(df_aba.columns) > 2 else df_aba.columns[0])
            
            df_aba[c_setor] = df_aba[c_setor].ffill()
            df_aba[c_sub] = df_aba[c_sub].fillna("")
            df_aba[c_cat] = df_aba[c_cat].ffill()
            
            def limpar_vagas(valor):
                if pd.isna(valor): return 0
                v_str = "".join(filter(str.isdigit, str(valor)))
                return int(v_str) if v_str != "" else 0
            
            # Extração posicional padrão para turnos (Manhã na coluna 4, Tarde na coluna 5)
            if df_aba.shape[1] >= 5:
                df_aba["VAGAS_MANHA"] = df_aba.iloc[:, 3].apply(limpar_vagas)
                df_aba["VAGAS_TARDE"] = df_aba.iloc[:, 4].apply(limpar_vagas)
            else:
                df_aba["VAGAS_MANHA"] = df_aba.iloc[:, -1].apply(limpar_vagas)
                df_aba["VAGAS_TARDE"] = 0
                
            df_aba["VAGAS_TOTAL"] = df_aba["VAGAS_MANHA"] + df_aba["VAGAS_TARDE"]
            
            df_aba["LOCAL_COMBINADO"] = df_aba.apply(
                lambda r: f"{r[c_setor]}" if (not r[c_sub] or str(r[c_setor]).upper() == str(r[c_sub]).upper() or str(r[c_sub]).strip() == "") else f"{r[c_setor]} - {r[c_sub]}",
                axis=1
            )
            
            # Expurgar linhas de totais acumulados no Excel para evitar duplicidade nos gráficos
            df_aba = df_aba[
                (~df_aba[c_setor].astype(str).str.upper().str.contains("TOTAL|QUANTITATIVO|HOSPITAL", na=False)) & 
                (df_aba["VAGAS_TOTAL"] > 0)
            ]
            return df_aba, c_setor, c_sub, c_cat

        # Processamento isolado de cada aba
        df_hcid, hc_setor, hc_sub, hc_cat = extrair_e_limpar_dados(aba_hcid_real)
        df_anexo, ax_setor, ax_sub, ax_cat = extrair_e_limpar_dados(aba_anexo_real)
        
    except Exception as e:
        st.error(f"Erro no processamento da planilha: {e}")
        st.stop()
else:
    st.info("💡 Por favor, arraste ou carregue sua planilha Excel para estruturar os painéis automaticamente.")
    st.stop()

# ==============================================================================
# 4. QUADRO 1: INDICADORES EXCLUSIVOS - HCID
# ==============================================================================
st.markdown("<h2 style='color: #008080; border-bottom: 2px solid #008080;'>🏢 QUADRO I - Mapeamento de Vagas Exclusivo HCID</h2>", unsafe_allow_html=True)

if not df_hcid.empty:
    # Métricas Estruturais (Filtro invisível por aba para não misturar)
    m1, m2, m3 = st.columns(3)
    t_vagas_hcid = int(df_hcid["VAGAS_TOTAL"].sum())
    t_setores_hcid = df_hcid["LOCAL_COMBINADO"].nunique()
    
    # Gráficos 1 e 2: Exibidos como KPIs Executivos de Impacto
    m1.metric(label="📊 1. Total de Vagas de Estágio no HCID", value=f"{t_vagas_hcid} Vagas")
    m2.metric(label="📍 2. Total de Setores Disponibilizados p/ Campo no HCID", value=t_setores_hcid)
    
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    
    with c1:
        # Gráfico 3: Setores disponibilizados para realização de estágio no HCID
        df_g3 = df_hcid.groupby("LOCAL_COMBINADO")["LOCAL_COMBINADO"].count().reset_index(name="Contagem")
        fig3 = px.bar(df_g3, x="Contagem", y="LOCAL_COMBINADO", orientation="h", title="3. Setores Disponibilizados para Estágio (HCID)", color_discrete_sequence=["#008080"])
        st.plotly_chart(fig3, use_container_width=True)
        
        # Gráfico 5: Total de vagas de estágio disponibilizados por setor no HCID
        df_g5 = df_hcid.groupby("LOCAL_COMBINADO")["VAGAS_TOTAL"].sum().reset_index()
        fig5 = px.bar(df_g5, x="VAGAS_TOTAL", y="LOCAL_COMBINADO", orientation="h", title="5. Total de Vagas Disponibilizadas por Setor (HCID)", color_discrete_sequence=["#4682B4"])
        st.plotly_chart(fig5, use_container_width=True)

    with c2:
        # Gráfico 4: Categorias profissionais contempladas no estágio por setor no HCID
        df_g4 = df_hcid.groupby(["LOCAL_COMBINADO", hc_cat])["VAGAS_TOTAL"].sum().reset_index()
        fig4 = px.bar(df_g4, x="VAGAS_TOTAL", y="LOCAL_COMBINADO", color=hc_cat, barmode="stack", title="4. Categorias Profissionais Contempladas por Setor (HCID)", color_discrete_sequence=cor_sequencia)
        st.plotly_chart(fig4, use_container_width=True)
        
        # Gráfico 6: Total de vagas de estágio do HCID por turno no HCID
        df_g6 = pd.DataFrame({
            "Turno": ["Manhã", "Tarde"],
            "Vagas": [df_hcid["VAGAS_MANHA"].sum(), df_hcid["VAGAS_TARDE"].sum()]
        })
        fig6 = px.bar(df_g6, x="Turno", y="Vagas", text="Vagas", title="6. Total de Vagas de Estágio por Turno (HCID)", color="Turno", color_discrete_map={"Manhã": "#4682B4", "Tarde": "#FF8C00"})
