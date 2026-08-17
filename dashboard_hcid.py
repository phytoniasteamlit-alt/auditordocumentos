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
    texto = texto.strip().upper().replace('\n', ' ').replace('\r', ' ')
    texto = "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return " ".join(texto.split())

# --- CABEÇALHO SUPERIOR ---
header_left, header_right = st.columns(2)
header_left.markdown("<h1 style='margin: 0; padding: 0; font-size: 2.2rem;'>📊 Painel de Indicadores de Estágio</h1>", unsafe_allow_html=True)
header_right.markdown(
    """
    <div style="text-align: right; line-height: 1.3; padding-bottom: 10px;">
        <span style="font-size: 16px; font-weight: bold;">🏥 Hospital da Cidade Dr. Jackson Lago</span><br>
        <span style="font-size: 14px; color: #ccc;">👩‍💼 Coordenação: Verônica Azevedo</span><br>
        <span style="font-size: 13px; color: #aaa; font-weight: 500;">📌 Setor Nep / Nepex</span>
    </div>
    """, 
    unsafe_allow_html=True
)
st.markdown("---")

# ==============================================================================
# 2. PAINEL DE CONTROLE (SIDEBAR)
# ==============================================================================
st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha de Estágios (.xlsx):", type=["xlsx"])

# ==============================================================================
# 3. FUNÇÃO DE TRATAMENTO DE DADOS (ISOLADA)
# ==============================================================================
def extrair_e_limpar_dados(uploaded_file, sheet_name):
    if not sheet_name:
        return pd.DataFrame()
    
    df_bruto = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    
    linha_cabecalho = 0
    for idx, row in df_bruto.iterrows():
        row_str = " ".join([str(x).upper() for x in row.dropna()])
        if "CATEGOR" in row_str or "PROFISS" in row_str or "SETOR" in row_str:
            linha_cabecalho = idx
            break
    
    cabecalhos = [str(c).strip().replace('\n', ' ') for c in df_bruto.iloc[linha_cabecalho]]
    df_dados = df_bruto.iloc[linha_cabecalho+1:].copy()
    df_dados.columns = cabecalhos
    
    idx_setor, idx_sub, idx_cat, idx_manha, idx_tarde = 0, 1, 2, 3, 4
    for idx_c, col_nome in enumerate(df_dados.columns):
        c_norm = normalizar_texto(col_nome)
        if "SUB" in c_norm: idx_sub = idx_c
        elif "SETOR" in c_norm or "CAMPO" in c_norm: idx_setor = idx_c
        elif "PROF" in c_norm or "CAT" in c_norm: idx_cat = idx_c
        elif "MANH" in c_norm: idx_manha = idx_c
        elif "TARD" in c_norm: idx_tarde = idx_c

    def extrair_inteiro(valor):
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan": 
            return 0
        v_str = "".join(filter(str.isdigit, str(valor)))
        return int(v_str) if v_str != "" else 0

    df_limpo = pd.DataFrame()
    df_limpo["SETOR"] = df_dados.iloc[:, idx_setor].astype(str).str.strip().ffill()
    df_limpo["SUB_SETOR"] = df_dados.iloc[:, idx_sub].fillna("GERAL").astype(str).str.strip()
    df_limpo["CATEGORIA"] = df_dados.iloc[:, idx_cat].fillna("NÃO ESPECIFICADO").astype(str).str.strip()
    df_limpo["MANHÃ"] = df_dados.iloc[:, idx_manha].apply(extrair_inteiro)
    df_limpo["TARDE"] = df_dados.iloc[:, idx_tarde].apply(extrair_inteiro)
    df_limpo["TOTAL_VAGAS"] = df_limpo["MANHÃ"] + df_limpo["TARDE"]
    
    linhas_validas = []
    for _, row in df_limpo.iterrows():
        txt_s = str(row["SETOR"]).upper()
        txt_c = str(row["CATEGORIA"]).upper()
        if "TOTAL" in txt_s or "TOTAL" in txt_c or txt_s == "NAN" or txt_c == "NAN":
            linhas_validas.append(False)
        else:
            linhas_validas.append(True)
    
    df_final = df_limpo[linhas_validas].copy()
    return df_final[df_final["TOTAL_VAGAS"] > 0]

# ==============================================================================
# 4. FUNÇÃO DE RENDERIZAÇÃO VISUAL (EVITA DUPLICAÇÃO DE CÓDIGO)
# ==============================================================================
def renderizar_painel_etapas(df_alvo, nome_aba_excel, chave_unica):
    if df_alvo.empty:
        st.warning(f"Nenhum dado válido ou ativo foi processado na aba '{nome_aba_excel}'. Verifique as colunas da planilha.")
        return

    st.caption(f"📂 Lendo dados da Aba: **'{nome_aba_excel}'**")

    # --- TOP CARD METRICS ---
    total_geral = df_alvo["TOTAL_VAGAS"].sum()
    total_m = df_alvo["MANHÃ"].sum()
    total_t = df_alvo["TARDE"].sum()
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Capacidade Total de Vagas", f"{total_geral} Vagas")
    m2.metric("Turno Manhã (Total)", f"{total_m} M")
    m3.metric("Turno Tarde (Total)", f"{total_t} T")
    
    st.markdown("---")
    
    col_esquerda, col_direita = st.columns(2)
    
    with col_esquerda:
        st.markdown("### 📊 Etapa 1: Visão Macro por Setor")
        st.caption("Volume macro consolidado por área principal.")
        
        df_macro = df_alvo.groupby("SETOR")["TOTAL_VAGAS"].sum().reset_index()
        df_macro = df_macro.sort_values(by="TOTAL_VAGAS", ascending=True)
        
        fig_macro = px.bar(
            df_macro,
            x="TOTAL_VAGAS",
            y="SETOR",
            orientation="h",
            labels={"TOTAL_VAGAS": "Total de Vagas", "SETOR": "Setor Principal"},
            color="TOTAL_VAGAS",
            color_continuous_scale=px.colors.sequential.Tealgrn
        )
        fig_macro.update_layout(showlegend=False, height=400, margin=dict(l=20, r=20, t=10, b=10))
        st.plotly_chart(fig_macro, use_container_width=True, key=f"macro_{chave_unica}")

    with col_direita:
        st.markdown("### 🔍 Etapa 2: Detalhar Sub-Setores e Turnos")
        st.caption("Filtre um setor para abrir o destrinchamento de sub-áreas e turnos.")
        
        setores_disponiveis = sorted(df_alvo["SETOR"].unique())
        setor_selecionado = st.selectbox("Selecione o Setor Principal:", setores_disponiveis, key=f"sel_{chave_unica}")
        
        df_filtrado = df_alvo[df_alvo["SETOR"] == setor_selecionado].copy()
        
        df_melted = df_filtrado.melt(
            id_vars=["SUB_SETOR", "CATEGORIA"],
            value_vars=["MANHÃ", "TARDE"],
            var_name="TURNO",
            value_name="VAGAS"
        )
        df_melted = df_melted[df_melted["VAGAS"] > 0]
        
        if not df_melted.empty:
            df_melted["SUB_E_CAT"] = df_melted["SUB_SETOR"] + " (" + df_melted["CATEGORIA"] + ")"
            
            fig_detalhe = px.bar(
                df_melted,
                x="VAGAS",
                y="SUB_E_CAT",
                color="TURNO",
                orientation="h",
                labels={"VAGAS": "Quantidade de Vagas", "SUB_E_CAT": "Sub-Setor (Profissão)"},
                color_discrete_map={"MANHÃ": "#008080", "TARDE": "#FF7F50"}
            )
            fig_detalhe.update_layout(barmode="stack", height=400, margin=dict(l=20, r=20, t=10, b=10))
            st.plotly_chart(fig_detalhe, use_container_width=True, key=f"detalhe_{chave_unica}")
        else:
            st.info("Nenhuma vaga ativa encontrada para os parâmetros do setor selecionado.")

    st.markdown("---")
    with st.expander("📄 Ver tabela de dados tratados desta unidade"):
        st.dataframe(df_alvo[["SETOR", "SUB_SETOR", "CATEGORIA", "MANHÃ", "TARDE", "TOTAL_VAGAS"]], use_container_width=True)

# ==============================================================================
# 5. EXECUÇÃO DO FLUXO PRINCIPAL
# ==============================================================================
if uploaded_file is not None:
    excel_file = pd.ExcelFile(uploaded_file)
    abas_disponiveis = excel_file.sheet_names
    
    # Identifica as abas dinamicamente ou por posição
    aba_hcid_real = next((op for op in ["HCID_BDD", "HCID", "HCID1", "DADOS"] if op in abas_disponiveis), abas_disponiveis[0])
    aba_anexo_real = next((op for op in ["ANEXO", "ANEXO2", "ANEXOS"] if op in abas_disponiveis), None)
    if aba_anexo_real is None and len(abas_disponiveis) > 1:
        aba_anexo_real = abas_disponiveis[1]
        
    df_hcid = extrair_e_limpar_dados(uploaded_file, aba_hcid_real)
    
    df_anexo = pd.DataFrame()
    if aba_anexo_real:
        df_anexo = extrair_e_limpar_dados(uploaded_file, aba_anexo_real)

    # Cria as abas de navegação visual
    tab_hcid, tab_anexos = st.tabs(["🏥 Hospital Geral (HCID)", "🏢 Unidades Anexas"])
    
    with tab_hcid:
        renderizar_painel_etapas(df_hcid, aba_hcid_real, "hcid")
        
    with tab_anexos:
        if aba_anexo_real:
            renderizar_painel_etapas(df_anexo, aba_anexo_real, "anexos")
        else:
            st.info("Sua planilha possui apenas 1 aba de dados ativos. Se possuir anexos, adicione-os na aba seguinte do arquivo Excel.")
else:
    st.info("💡 Por favor, arraste ou carregue sua planilha Excel para estruturar os painéis automaticamente.")
