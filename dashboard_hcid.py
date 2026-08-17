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
# 3. MOTOR INTELIGENTE MATRICIAL (VARREDURA REAL DE MESES E TEXTOS VAZIOS)
# ==============================================================================
def processar_calendario_dinamico(uploaded_file, numero_posicao_aba):
    try:
        # Lê a aba inteira desde o topo para não perder os meses mesclados
        df_raw = pd.read_excel(uploaded_file, sheet_name=numero_posicao_aba, header=None)
        if df_raw.empty or len(df_raw) <= 7:
            return pd.DataFrame()
            
        # Captura as linhas horizontais estruturais do cabeçalho de datas
        # Linha 4 (índice 3) = Meses, Linha 5 (índice 4) = Dias, Linha 6 (índice 5) = Turnos
        linha_meses = pd.Series(df_raw.iloc[3, :]).ffill().fillna("").astype(str).tolist()
        linha_dias = pd.Series(df_raw.iloc[4, :]).ffill().fillna("").astype(str).tolist()
        linha_turnos = pd.Series(df_raw.iloc[5, :]).ffill().fillna("").astype(str).tolist()
        
        # Isola o corpo de dados reais ignorando os títulos (Linha 8 física / índice 7 em diante)
        df_corpo = df_raw.iloc[7:].copy().reset_index(drop=True)
        
        # Preenchimento em cascata vertical das informações estruturais dos setores
        df_corpo.iloc[:, 0] = df_corpo.iloc[:, 0].replace(["nan", "NAN", ""], None).ffill().fillna("GERAL")
        df_corpo.iloc[:, 1] = df_corpo.iloc[:, 1].replace(["nan", "NAN", ""], None).ffill().fillna("GERAL")
        df_corpo.iloc[:, 2] = df_corpo.iloc[:, 2].replace(["nan", "NAN", ""], None).ffill().fillna("NÃO ESPECIFICADO")
        
        registros_vagas = []
        num_colunas_total = len(df_raw.columns)
        
        for idx_row in range(len(df_corpo)):
            setor = str(df_corpo.iloc[idx_row, 0]).strip().upper()
            sub_setor = str(df_corpo.iloc[idx_row, 1]).strip().upper()
            categoria = str(df_corpo.iloc[idx_row, 2]).strip().upper()
            
            # Filtro rígido contra linhas pretas vazias de TOTAL nativas do Excel
            if "TOTAL" in setor or "TOTAL" in categoria or categoria == "" or setor == "SETOR" or setor == "GERAL":
                continue
                
            # Varre horizontalmente todas as colunas de calendário a partir da coluna de índice 8
            for col_idx in range(8, num_colunas_total):
                vaga_bruta = df_corpo.iloc[idx_row, col_idx]
                qtd_vagas = extrair_numero(vaga_bruta)
                
                # REGRA DE OURO SOLICITADA: Célula vazia ou zerada não entra no banco e some do gráfico!
                if qtd_vagas > 0:
                    mes_name = str(linha_meses[col_idx]).strip().upper()
                    dia_name = str(linha_dias[col_idx]).strip().upper()
                    turno_name = str(linha_turnos[col_idx]).strip().upper()
                    
                    # Filtra colunas válidas que pertencem aos meses de estágio
                    if any(m in mes_name for m in ["AGO", "SET", "OUT", "NOV", "DEZ"]) or "VAGAS" in mes_name:
                        if "VAGAS" in mes_name or mes_name == "": 
                            mes_name = "AGOSTO"
                            
                        final_turno = "MANHÃ" if "MANH" in turno_name or "MANH" in dia_name else "TARDE"
                        
                        registros_vagas.append({
                            "SETOR": setor,
                            "SUB_SETOR": sub_setor if sub_setor != "" else "GERAL",
                            "CATEGORIA": categoria,
                            "MÊS": mes_name,
                            "DIA_SEMANA": dia_name if any(d in dia_name for d in ["SEG", "TER", "QUA", "QUI", "SEX"]) else "SEGUNDA",
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
    
    # Processa as matrizes de calendário de forma 100% independente por índice de aba
    df_hcid = processar_calendario_dinamico(uploaded_file, 0)
    df_anexos = processar_calendario_dinamico(uploaded_file, 1) if len(abas_planilha) > 1 else pd.DataFrame()

    tab_hcid, tab_anexos = st.tabs(["🏥 Hospital Geral (HCID)", "🏢 Unidades Anexas"])

    # --------------------------------==========================================
    # CONTEÚDO EXCLUSIVO DA GUIA 1: SOMENTE HCID (CALENDÁRIO REAL ATIVO)
    # --------------------------------==========================================
    with tab_hcid:
        st.markdown("<div style='background-color: #1a2a3a; padding: 12px; border-radius: 5px; margin-bottom: 20px;'><h2 style='margin:0; font-size:1.4rem; color:#fff;'>📊 QUADRO DE INDICADORES - SOMENTE HCID (ESCALA REAL)</h2></div>", unsafe_allow_html=True)
        
        if not df_hcid.empty:
            # Filtro interativo de Meses no topo da aba para auditar o calendário de forma organizada
            mes_selecionado_h = st.selectbox("📅 Selecione o Mês para Visualizar o Calendário (HCID):", sorted(df_hcid["MÊS"].unique()))
            df_filtro_h = df_hcid[df_hcid["MÊS"] == mes_selecionado_h].copy()
            
            # Caixas de texto de métricas calculando em tempo real com base no mês filtrado
            t_vagas_h = df_filtro_h["VAGAS"].sum()
            t_setores_h = df_filtro_h["SETOR"].nunique()
            t_m_h = df_filtro_h[df_filtro_h["TURNO"] == "MANHÃ"]["VAGAS"].sum()
            t_t_h = df_filtro_h[df_filtro_h["TURNO"] == "TARDE"]["VAGAS"].sum()
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(f"📑 Total de vagas de estágio em {mes_selecionado_h}", f"{t_vagas_h} Vagas")
            c2.metric("🏥 Setores com estagiários ativos", f"{t_setores_h} Setores")
            c3.metric("☀️ Total do turno manhã no mês", f"{t_m_h} M")
            c4.metric("🌙 Total do turno tarde no mês", f"{t_t_h} T")
            
            st.markdown("<br>", unsafe_allow_html=True)
            col1_h, col2_h = st.columns(2)
            
            with col1_h:
                st.markdown("##### 1️⃣ Total de vagas e de estágio no HCID")
                df_g1_h = df_filtro_h.groupby("SETOR")["VAGAS"].sum().reset_index()
                f1_h = px.bar(df_g1_h, x="SETOR" if is_vert else "VAGAS", y="VAGAS" if is_vert else "SETOR", orientation="v" if is_vert else "h", text_auto=True, color_discrete_sequence=seq_cores)
                f1_h.update_layout(height=320, margin=dict(l=10,r=15,t=10,b=10), xaxis_title="Setor" if is_vert else "Vagas", yaxis_title="Vagas" if is_vert else "Setor")
                f1_h.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(f1_h, use_container_width=True)
                
                st.markdown("##### 3️⃣ Setores disponibilizados para realização de estágio no HCID")
                df_g3_h = df_filtro_h.groupby("SETOR")["VAGAS"].sum().reset_index().sort_values(by="VAGAS", ascending=True)
