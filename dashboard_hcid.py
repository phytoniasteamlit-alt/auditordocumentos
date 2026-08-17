import streamlit as st
import pandas as pd
import plotly.express as px

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (Interface Dashboard Executivo)
# ==============================================================================
st.set_page_config(
    page_title="Painel Geral de Estágios - HCID & ANEXOS",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
# 3. MOTOR DE PROCESSAMENTO POR COORDENADAS FÍSICAS ESTRITAS
# ==============================================================================
def extrair_e_limpar_dados(uploaded_file, sheet_name):
    if not sheet_name:
        return pd.DataFrame()
    
    # Lê a planilha bruta sem pular linhas para preservar o alinhamento das colunas A, B, C, D, E
    df_bruto = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    
    if df_bruto.empty or len(df_bruto) <= 5:
        return pd.DataFrame()
        
    # Pula as linhas de títulos e a linha mesclada superior (Linhas 1 a 4 do Excel)
    # Isola os dados começando exatamente da linha onde está escrito "Manhã" e "Tarde" (Linha 5 indexada como 4)
    df_dados = df_bruto.iloc[4:].copy()
    
    # MAPEAMENTO ESTREITO POR COORDENADA FÍSICA DE COLUNA (A=0, B=1, C=2, D=3, E=4)
    idx_setor = 0  # Coluna A
    idx_sub   = 1  # Coluna B
    idx_cat   = 2  # Coluna C
    idx_manha = 3  # Coluna D
    idx_tarde = 4  # Coluna E

    def extrair_inteiro(valor):
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan": 
            return None  # Retorna None para o ffill propagar as vagas do topo do bloco
        # Remove os textos complementares ("por turno") e extrai apenas os numéricos
        v_str = "".join(filter(str.isdigit, str(valor)))
        return int(v_str) if v_str != "" else 0

    df_limpo = pd.DataFrame()
    
    # Tratamento estrutural do Setor Principal considerando mesclagem vertical (ffill)
    df_limpo["SETOR"] = df_dados.iloc[:, idx_setor].astype(str).str.strip()
    df_limpo["SETOR"] = df_limpo["SETOR"].replace(["nan", "NAN", ""], pd.NA).ffill()
    
    # Tratamento de sub-áreas e profissões
    df_limpo["SUB_SETOR"] = df_dados.iloc[:, idx_sub].fillna("").astype(str).str.strip().replace(["nan", "NAN"], "")
    df_limpo["CATEGORIA"] = df_dados.iloc[:, idx_cat].fillna("").astype(str).str.strip().replace(["nan", "NAN"], "")
    
    # Extração isolada e limpa das colunas físicas D e E
    df_limpo["MANHÃ"] = df_dados.iloc[:, idx_manha].apply(extrair_inteiro)
    df_limpo["TARDE"] = df_dados.iloc[:, idx_tarde].apply(extrair_inteiro)
    
    # PREENCHIMENTO EM CASCATA: Alimenta as subprofissões com as vagas do cabeçalho do bloco
    df_limpo["MANHÃ"] = df_limpo["MANHÃ"].ffill().fillna(0).astype(int)
    df_limpo["TARDE"] = df_limpo["TARDE"].ffill().fillna(0).astype(int)
    df_limpo["TOTAL_VAGAS"] = df_limpo["MANHÃ"] + df_limpo["TARDE"]
    
    # Filtro rígido para expurgar linhas de cabeçalho, lixo ou totalizadores parciais
    linhas_validas = []
    for _, row in df_limpo.iterrows():
        txt_s = str(row["SETOR"]).upper()
        txt_c = str(row["CATEGORIA"]).upper()
        if "TOTAL" in txt_s or "TOTAL" in txt_c or txt_c == "" or "SETOR" in txt_s or "CATEGOR" in txt_c or "MANH" in txt_c or "TARD" in txt_s:
            linhas_validas.append(False)
        else:
            linhas_validas.append(True)
            
    df_final = df_limpo[linhas_validas].copy()
    df_final["SUB_SETOR"] = df_final["SUB_SETOR"].apply(lambda x: "GERAL" if x == "" else x)
    
    return df_final[df_final["TOTAL_VAGAS"] > 0]

# ==============================================================================
# 4. FUNÇÃO DE RENDERIZAÇÃO DO DASHBOARD PROGRESSIVO POR ETAPAS
# ==============================================================================
def renderizar_painel_etapas(df_alvo, nome_aba_excel, chave_unica):
    if df_alvo.empty:
        st.warning(f"Nenhum dado ativo foi localizado na aba '{nome_aba_excel}'. Verifique o formato do arquivo.")
        return

    st.caption(f"📂 Fonte dos dados ativa: Aba **'{nome_aba_excel}'**")

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
            color_continuous_scale=px.colors.sequential.Tealgrn,
            text_auto=True
        )
        fig_macro.update_layout(showlegend=False, height=450, margin=dict(l=20, r=35, t=10, b=10))
        fig_macro.update_traces(textposition="outside", cliponaxis=False)
        st.plotly_chart(fig_macro, use_container_width=True, key=f"macro_{chave_unica}")

    with col_direita:
        st.markdown("### 🔍 Etapa 2: Detalhar Sub-Setores e Turnos")
        st.caption("Filtre um setor para abrir o destrinchamento de sub-areas, turnos e profissões.")
        
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
                color_discrete_map={"MANHÃ": "#008080", "TARDE": "#FF7F50"},
                text_auto=True 
            )
            fig_detalhe.update_layout(barmode="stack", height=450, margin=dict(l=20, r=35, t=10, b=10))
            fig_detalhe.update_traces(textposition="inside", insidetextanchor="middle")
            st.plotly_chart(fig_detalhe, use_container_width=True, key=f"detalhe_{chave_unica}")
        else:
            st.info("Nenhuma vaga ativa encontrada para os parâmetros do setor selecionado.")

    st.markdown("---")
    with st.expander("📄 Ver tabela de dados tratados desta unidade"):
        st.dataframe(df_alvo[["SETOR", "SUB_SETOR", "CATEGORIA", "MANHÃ", "TARDE", "TOTAL_VAGAS"]], use_container_width=True)

# ==============================================================================
# 5. EXECUÇÃO DO FLUXO PRINCIPAL (LINHA SEGUIDA TOTALMENTE ALINHADA)
# ==============================================================================
if uploaded_file is not None:
    excel_file = pd.ExcelFile(uploaded_file)
    abas_disponiveis = excel_file.sheet_names
    
    # Identifica as abas disponíveis de maneira direta
    aba_hcid_real = "HCID_BDD" if "HCID_BDD" in abas_disponiveis else abas_disponiveis
    aba_anexo_real = "ANEXO" if "ANEXO" in abas_disponiveis else ("ANEXO2" if "ANEXO2" in abas_disponiveis else None)
    
    if aba_anexo_real is None and len(abas_disponiveis) > 1:
        aba_anexo_real = abas_disponiveis if abas_disponiveis != aba_hcid_real else None
        
    df_hcid = extrair_e_limpar_dados(uploaded_file, aba_hcid_real)
    df_anexo = extrair_e_limpar_dados(uploaded_file, aba_anexo_real) if aba_anexo_real else pd.DataFrame()

    tab_hcid, tab_anexos = st.tabs(["🏥 Hospital Geral (HCID)", "🏢 Unidades Anexas"])
    
    with tab_hcid:
        renderizar_painel_etapas(df_hcid, aba_hcid_real, "hcid")
        
    with tab_anexos:
        if aba_anexo_real and not df_anexo.empty:
            renderizar_painel_etapas(df_anexo, aba_anexo_real, "anexos")
        else:
            st.info("Sua planilha possui apenas 1 aba de dados ativos. Se tiver anexos, adicione a segunda aba ao arquivo Excel.")
else:
