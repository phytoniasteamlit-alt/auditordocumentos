import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (Interface Dashboard Executivo)
# ==============================================================================
st.set_page_config(
    page_title="Painel de Controle de Estágios - HCID & ANEXOS",
    layout="wide",
    initial_sidebar_state="expanded"
)

def normalizar_texto(texto):
    if pd.isna(texto) or not isinstance(texto, str):
        return ""
    texto = texto.strip().upper().replace('\n', ' ').replace('\r', ' ')
    texto = "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return " ".join(texto.split())

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
# 2. PAINEL DE CONTROLE (SIDEBAR)
# ==============================================================================
st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha de Estágios (.xlsx):", type=["xlsx"])

# ==============================================================================
# 3. MOTOR MATRICIAL DE PROCESSAMENTO ISOLADO POR ABA
# ==============================================================================
def processar_aba_matricial(uploaded_file, sheet_name_fallback, padrao_procurado):
    excel_file = pd.ExcelFile(uploaded_file)
    abas = excel_file.sheet_names
    
    # Caça a aba exata por correspondência de texto para evitar quebras por digitação
    aba_real = next((op for op in padrao_procurado if op in abas), None)
    if not aba_real:
        aba_real = sheet_name_fallback if sheet_name_fallback in abas else None
        
    if not aba_real:
        return pd.DataFrame()
        
    df_raw = pd.read_excel(uploaded_file, sheet_name=aba_real, header=None)
    if df_raw.empty or len(df_raw) <= 7:
        return pd.DataFrame()
        
    # Tratamento horizontal do calendário mesclado (Meses, Dias e Turnos)
    linha_meses = pd.Series(df_raw.iloc[3, :]).ffill().fillna("").astype(str).tolist()
    linha_dias = pd.Series(df_raw.iloc[4, :]).ffill().fillna("").astype(str).tolist()
    linha_turnos = pd.Series(df_raw.iloc[5, :]).ffill().fillna("").astype(str).tolist()
    
    # Isola o corpo de registros (Linha 8 física / índice 7 em diante)
    df_corpo = df_raw.iloc[7:].copy().reset_index(drop=True)
    
    # Tratamento vertical das colunas estruturais textuais
    df_corpo.iloc[:, 0] = df_corpo.iloc[:, 0].replace(["nan", "NAN", ""], pd.NA).ffill().fillna("GERAL")
    df_corpo.iloc[:, 1] = df_corpo.iloc[:, 1].replace(["nan", "NAN", ""], pd.NA).ffill().fillna("GERAL")
    df_corpo.iloc[:, 2] = df_corpo.iloc[:, 2].replace(["nan", "NAN", ""], pd.NA).ffill().fillna("NÃO ESPECIFICADO")
    
    registros_vagas = []
    num_colunas_total = len(df_raw.columns)
    
    for idx_row in range(len(df_corpo)):
        setor = str(df_corpo.iloc[idx_row, 0]).strip().upper()
        sub_setor = str(df_corpo.iloc[idx_row, 1]).strip().upper()
        categoria = str(df_corpo.iloc[idx_row, 2]).strip().upper()
        
        if "TOTAL" in setor or "TOTAL" in categoria or categoria == "" or setor == "SETOR" or setor == "GERAL":
            continue
            
        for col_idx in range(8, num_colunas_total):
            vaga_bruta = df_corpo.iloc[idx_row, col_idx]
            qtd_vagas = extrair_numero(vaga_bruta)
            
            # Recuperação em cascata se a célula diária pontual estiver sem preenchimento
            if qtd_vagas == 0:
                turno_atual = str(linha_turnos[col_idx]).strip().upper()
                vaga_padrao_celula = df_corpo.iloc[idx_row, 3] if "MANH" in turno_atual else df_corpo.iloc[idx_row, 4]
                if pd.notna(vaga_bruta) and str(vaga_bruta).strip() != "" and str(vaga_bruta).strip().lower() != "nan":
                    qtd_vagas = extrair_numero(vaga_padrao_celula)
            
            if qtd_vagas > 0:
                mes_name = str(linha_meses[col_idx]).strip().upper()
                dia_name = str(linha_dias[col_idx]).strip().upper()
                turno_name = str(linha_turnos[col_idx]).strip().upper()
                
                if any(m in mes_name for m in ["AGO", "SET", "OUT", "NOV", "DEZ"]) or "VAGAS" in mes_name:
                    if "VAGAS" in mes_name or mes_name == "": 
                        mes_name = "AGOSTO"
                        
                    final_turno = "MANHÃ" if "MANH" in turno_name or "MANH" in dia_name else "TARDE"
                        
                    registros_vagas.append({
                        "SETOR": setor,
                        "SUB_SETOR": sub_setor,
                        "CATEGORIA": categoria,
                        "MÊS": mes_name,
                        "DIA_SEMANA": dia_name if any(d in dia_name for d in ["SEG", "TER", "QUA", "QUI", "SEX"]) else "SEGUNDA",
                        "TURNO": final_turno,
                        "VAGAS": qtd_vagas
                    })
                    
    return pd.DataFrame(registros_vagas)

# ==============================================================================
# 4. EXECUÇÃO DO PROCESSAMENTO EM AMBIENTES ISOLADOS (SEM MISTURA DE DADOS)
# ==============================================================================
if uploaded_file is not None:
    # Processa estritamente o bloco HCID
    df_hcid = processar_aba_matricial(uploaded_file, "HCID", ["HCID_BDD", "HCID", "HCID1"])
    
    # Processa estritamente o bloco ANEXOS
    df_anexos = processar_aba_matricial(uploaded_file, "ANEXO", ["ANEXO", "ANEXO2", "ANEXOS"])

    # ==========================================================================
    # QUADRO 1: CONJUNTO EXCLUSIVO HOSPITAL GERAL (HCID)
    # ==========================================================================
    st.markdown("<div style='background-color: #1a2a3a; padding: 12px; border-radius: 5px; margin-bottom: 20px;'><h2 style='margin:0; font-size:1.6rem; color:#fff;'>🏥 QUADRO DE INDICADORES - SOMENTE HCID</h2></div>", unsafe_allow_html=True)
    
    if not df_hcid.empty:
        # Caixas de Texto com as métricas propostas integradas
        t_vagas_h = df_hcid["VAGAS"].sum()
        t_setores_h = df_hcid["SETOR"].nunique()
        t_m_h = df_hcid[df_hcid["TURNO"] == "MANHÃ"]["VAGAS"].sum()
        t_t_h = df_hcid[df_hcid["TURNO"] == "TARDE"]["VAGAS"].sum()
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total de vagas de estágio geral HCID", f"{t_vagas_h} Vagas")
        c2.metric("Total de setores disponibilizados no HCID", f"{t_setores_h} Setores")
        c3.metric("Total de vagas por turno MANHÃ (HCID)", f"{t_m_h} M")
        c4.metric("Total de vagas por turno TARDE (HCID)", f"{t_t_h} T")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- RENDERIZAÇÃO DOS 7 GRÁFICOS DO HCID ---
        col1_h, col2_h = st.columns(2)
        
        with col1_h:
            # Gráfico 1: Total de vagas de estágio geral no HCID (Visão Cronológica Mensal)
            st.markdown("##### Gráfico 1: Total de vagas e de estágio no HCID")
            df_m_h = df_hcid.groupby("MÊS")["VAGAS"].sum().reset_index()
            f1_h = px.bar(df_m_h, x="MÊS", y="VAGAS", color="MÊS", text_auto=True, color_discrete_sequence=px.colors.qualitative.Set2)
            f1_h.update_layout(height=300, showlegend=False, margin=dict(l=10,r=10,t=10,b=10))
            st.plotly_chart(f1_h, use_container_width=True)
            
            # Gráfico 3: Setores disponibilizados para realização de estágio no HCID
            st.markdown("##### Gráfico 3: Setores disponibilizados para realização de estágio o HCID")
            df_s_h = df_hcid.groupby("SETOR")["VAGAS"].sum().reset_index().sort_values(by="VAGAS", ascending=True)
            f3_h = px.bar(df_s_h, x="VAGAS", y="SETOR", orientation="h", text_auto=True, color="VAGAS", color_continuous_scale=px.colors.sequential.Tealgrn)
            f3_h.update_layout(height=300, coloraxis_showscale=False, margin=dict(l=10,r=10,t=10,b=10))
            st.plotly_chart(f3_h, use_container_width=True)

            # Gráfico 5: Total de vagas de estágio disponibilizados por setor no HCID
            st.markdown("##### Gráfico 5: Total de vagas de estágio disponibilizados por setor no HCID")
            df_v_s_h = df_hcid.groupby("SETOR")["VAGAS"].sum().reset_index()
            f5_h = px.pie(df_v_s_h, values="VAGAS", names="SETOR", hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
            f5_h.update_layout(height=300, margin=dict(l=10,r=10,t=10,b=10))
            st.plotly_chart(f5_h, use_container_width=True)

            # Gráfico 7: Total de estagiários por turno por dia no HCID
            st.markdown("##### Gráfico 7: Total de estagiários por turno por dia no HCID")
            df_d_t_h = df_hcid.groupby(["DIA_SEMANA", "TURNO"])["VAGAS"].sum().reset_index()
