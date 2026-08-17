import streamlit as st
import pandas as pd
import plotly.express as px

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (Interface Dashboard Executivo)
# ==============================================================================
st.set_page_config(
    page_title="Painel de Controle de Estágios - HCID & ANEXOS",
    layout="wide",
    initial_sidebar_state="expanded"
)

def extrair_numero(valor):
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan":
        return 0
    v_str = "".join(filter(str.isdigit, str(valor)))
    return int(v_str) if v_str != "" else 0

# --- CABEÇALHO SUPERIOR INSTITUCIONAL ---
header_left, header_right = st.columns(2)
header_left.markdown("<h1 style='margin: 0; padding: 0; font-size: 2.2rem;'>📊 Painel Unificado de Indicadores de Estágio</h1>", unsafe_allow_html=True)
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
# 2. PAINEL DE CONTROLE E CUSTOMIZAÇÃO VISUAL (SIDEBAR)
# ==============================================================================
st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha de Estágios (.xlsx):", type=["xlsx"])

st.sidebar.markdown("---")
st.sidebar.subheader("🎨 Customização Interativa")

paleta_selecionada = st.sidebar.selectbox(
    "Tema de Cores dos Gráficos:",
    options=["Padrão Hospitalar", "Tons Pastéis", "Vibrante", "Esmeralda"],
    index=0
)

estilo_grafico = st.sidebar.radio(
    "Orientação das Barras:",
    options=["Barras Verticais", "Barras Horizontais"],
    index=1
)

if paleta_selecionada == "Tons Pastéis":
    seq_cores = px.colors.qualitative.Pastel
elif paleta_selecionada == "Vibrante":
    seq_cores = px.colors.qualitative.Prism
elif paleta_selecionada == "Esmeralda":
    seq_cores = px.colors.sequential.Mint
else:
    seq_cores = ["#4682B4", "#008080", "#20B2AA", "#5F9EA0", "#B0C4DE"]

is_vert = estilo_grafico == "Barras Verticais"

# ==============================================================================
# 3. MOTOR AMBIDESTRO PROTEGIDO CONTRA SINTAXES DIVERGENTES DE ABAS
# ==============================================================================
def extrair_dados_seguros(uploaded_file, numero_posicao_aba):
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=numero_posicao_aba, header=None)
        if df_raw.empty or len(df_raw) <= 7:
            return pd.DataFrame()
            
        # Determina os cabeçalhos de turnos para proteção horizontal
        linha_turnos = pd.Series(df_raw.iloc[5, :]).ffill().fillna("").astype(str).tolist()
        
        # Isola o corpo de dados reais (Linha 8 física / índice 7)
        df_corpo = df_raw.iloc[7:].copy().reset_index(drop=True)
        
        # Correção vertical em cascata para herdar células mescladas de setores
        setores_col = df_corpo.iloc[:, 0].astype(str).str.strip().replace(["nan", "NAN", ""], None).ffill().fillna("GERAL")
        sub_setores_col = df_corpo.iloc[:, 1].astype(str).str.strip().replace(["nan", "NAN", ""], None).ffill().fillna("GERAL")
        categorias_col = df_corpo.iloc[:, 2].astype(str).str.strip().replace(["nan", "NAN", ""], None).ffill().fillna("NÃO ESPECIFICADO")
        
        registros_vagas = []
        
        # LÓGICA DE CAPTURA INTELIGENTE POR TIPO DE ABA
        if numero_posicao_aba == 0:
            # ABA 1 (HCID_BDD): Lê estritamente as colunas D e E fixas de forma limpa
            vagas_m_padrao = df_corpo.iloc[:, 3].astype(str).str.strip().replace(["nan", "NAN", ""], None).ffill().fillna("0")
            vagas_t_padrao = df_corpo.iloc[:, 4].astype(str).str.strip().replace(["nan", "NAN", ""], None).ffill().fillna("0")
            
            for idx_row in range(len(df_corpo)):
                q_m = extrair_numero(vagas_m_padrao.iloc[idx_row])
                q_t = extrair_numero(vagas_t_padrao.iloc[idx_row])
                
                if q_m > 0:
                    registros_vagas.append({"SETOR": str(setores_col.iloc[idx_row]).upper(), "SUB_SETOR": str(sub_setores_col.iloc[idx_row]).upper(), "CATEGORIA": str(categorias_col.iloc[idx_row]).upper(), "TURNO": "MANHÃ", "VAGAS": q_m})
                if q_t > 0:
                    registros_vagas.append({"SETOR": str(setores_col.iloc[idx_row]).upper(), "SUB_SETOR": str(sub_setores_col.iloc[idx_row]).upper(), "CATEGORIA": str(categorias_col.iloc[idx_row]).upper(), "TURNO": "TARDE", "VAGAS": q_t})
                    
        else:
            # ABA 2 (ANEXO): Varre a matriz de forma ampla tratando as colunas com dados em branco
            num_colunas_total = len(df_raw.columns)
            for idx_row in range(len(df_corpo)):
                setor_a = str(setores_col.iloc[idx_row]).upper()
                cat_a = str(categorias_col.iloc[idx_row]).upper()
                
                if "TOTAL" in setor_a or "TOTAL" in cat_a or cat_a == "" or setor_a == "SETOR":
                    continue
                    
                # Varre a linha horizontalmente pulando as primeiras colunas estruturais
                for col_idx in range(8, num_colunas_total):
                    vaga_bruta = df_corpo.iloc[idx_row, col_idx]
                    qtd_vagas = extrair_numero(vaga_bruta)
                    
                    # Ignora células em branco. Elas reaparecerão no gráfico automaticamente ao serem preenchidas
                    if qtd_vagas > 0:
                        turno_atual = str(linha_turnos[col_idx]).strip().upper()
                        final_turno = "MANHÃ" if "MANH" in turno_atual else "TARDE"
                        
                        registros_vagas.append({
                            "SETOR": setor_a,
                            "SUB_SETOR": str(sub_setores_col.iloc[idx_row]).upper(),
                            "CATEGORIA": cat_a,
                            "TURNO": final_turno,
                            "VAGAS": qtd_vagas
                        })
                        
        return pd.DataFrame(registros_vagas)
    except:
        return pd.DataFrame()

# ==============================================================================
# 4. AMBIENTE DE NAVEGAÇÃO POR ABAS CONSOLIDADAS MANTIDO ISOLADO
# ==============================================================================
if uploaded_file is not None:
    excel_file = pd.ExcelFile(uploaded_file)
    abas_planilha = excel_file.sheet_names
    
    # Processamento assíncrono blindado contra quebras estruturais
    df_hcid = extrair_dados_seguros(uploaded_file, 0)
    df_anexos = extrair_dados_seguros(uploaded_file, 1) if len(abas_planilha) > 1 else pd.DataFrame()

    tab_hcid, tab_anexos = st.tabs(["🏥 Hospital Geral (HCID)", "🏢 Unidades Anexas"])

    # --------------------------------==========================================
    # CONTEÚDO EXCLUSIVO DA GUIA 1: SOMENTE HCID
    # --------------------------------==========================================
    with tab_hcid:
        st.markdown("<div style='background-color: #1a2a3a; padding: 12px; border-radius: 5px; margin-bottom: 20px;'><h2 style='margin:0; font-size:1.4rem; color:#fff;'>📊 QUADRO DE INDICADORES - SOMENTE HCID</h2></div>", unsafe_allow_html=True)
        
        if not df_hcid.empty:
            t_vagas_h = df_hcid["VAGAS"].sum()
            t_setores_h = df_hcid["SETOR"].nunique()
            t_m_h = df_hcid[df_hcid["TURNO"] == "MANHÃ"]["VAGAS"].sum()
            t_t_h = df_hcid[df_hcid["TURNO"] == "TARDE"]["VAGAS"].sum()
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("📑 Total de vagas de estágio geral HCID", f"{t_vagas_h} Vagas")
            c2.metric("🏥 Total de setores disponibilizados no HCID", f"{t_setores_h} Setores")
            c3.metric("☀️ Total de vagas de estágio do HCID por turno manhã", f"{t_m_h} M")
            c4.metric("🌙 Total de vagas de estágio do HCID tarde", f"{t_t_h} T")
            
            st.markdown("<br>", unsafe_allow_html=True)
            col1_h, col2_h = st.columns(2)
            
            with col1_h:
                st.markdown("##### 1️⃣ Total de vagas e de estágio no HCID")
                df_g1_h = df_hcid.groupby("SETOR")["VAGAS"].sum().reset_index()
                f1_h = px.bar(df_g1_h, x="SETOR" if is_vert else "VAGAS", y="VAGAS" if is_vert else "SETOR", orientation="v" if is_vert else "h", text_auto=True, color_discrete_sequence=seq_cores)
                f1_h.update_layout(height=320, margin=dict(l=10,r=15,t=10,b=10), xaxis_title="Setor" if is_vert else "Vagas", yaxis_title="Vagas" if is_vert else "Setor")
                f1_h.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(f1_h, use_container_width=True)
                
                st.markdown("##### 3️⃣ Setores disponibilizados para realização de estágio no HCID")
                df_g3_h = df_hcid.groupby("SETOR")["VAGAS"].sum().reset_index().sort_values(by="VAGAS", ascending=True)
                f3_h = px.bar(df_g3_h, x="SETOR" if is_vert else "VAGAS", y="VAGAS" if is_vert else "SETOR", orientation="v" if is_vert else "h", text_auto=True, color="VAGAS", color_continuous_scale=px.colors.sequential.Tealgrn)
                f3_h.update_layout(height=320, coloraxis_showscale=False, margin=dict(l=10,r=15,t=10,b=10), xaxis_title="Setor" if is_vert else "Vagas", yaxis_title="Vagas" if is_vert else "Setor")
