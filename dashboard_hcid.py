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

# --- CABEÇALHO SUPERIOR (Atualizado com Setor Nep / Nepex) ---
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

# NOVO: Customização de estilos das barras exigido na solicitação
estilo_grafico = st.sidebar.radio(
    "Estilo Visual dos Gráficos:",
    options=["Barras Horizontais", "Barras Verticais"],
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

# Configuração da orientação das barras com base no seletor do usuário
sub_or = "h" if estilo_grafico == "Barras Horizontais" else "v"

# ==============================================================================
# 3. MOTOR DE PROCESSAMENTO DE DADOS EXECUTIVO
# ==============================================================================
if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        abas_disponiveis = excel_file.sheet_names
        
        aba_hcid_real = None
        for opcao in ["HCID_BDD", "HCID", "HCID1", "DADOS"]:
            if opcao in abas_disponiveis:
                aba_hcid_real = opcao
                break
                
        aba_anexo_real = None
        for opcao in ["ANEXO", "ANEXO2", "ANEXOS"]:
            if opcao in abas_disponiveis:
                aba_anexo_real = opcao
                break
        
        def extrair_e_limpar_dados(sheet_name):
            if not sheet_name or sheet_name not in abas_disponiveis:
                return pd.DataFrame(), pd.DataFrame(), "SETOR_RAW", "SUB_SETOR_RAW", "CATEGORIA_RAW"
            
            df_bruto = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
            
            linha_cabecalho = 0
            for idx, row in df_bruto.iterrows():
                row_str = " ".join([str(x).upper() for x in row.dropna()])
                if "CATEGOR" in row_str or "PROFISS" in row_str or "SETOR" in row_str:
                    linha_cabecalho = idx
                    break
            
            cabecalhos_originais = [str(c).strip().replace('\n', ' ') for c in df_bruto.iloc[linha_cabecalho]]
            df_aba = df_bruto.iloc[linha_cabecalho+1:].copy()
            df_aba.columns = cabecalhos_originais
            
            col_setor, col_sub, col_cat = None, None, None
            col_manha, col_tarde = None, None
            
            for idx_c, col_nome in enumerate(df_aba.columns):
                c_norm = normalizar_texto(col_nome)
                if "SUB" in c_norm: col_sub = idx_c
                elif "SETOR" in c_norm or "CAMPO" in c_norm: col_setor = idx_c
                elif "PROF" in c_norm or "CAT" in c_norm: col_cat = idx_c
                elif "MANH" in c_norm: col_manha = idx_c
                elif "TARD" in c_norm: col_tarde = idx_c

            col_setor = col_setor if col_setor is not None else 0
            col_sub = col_sub if col_sub is not None else 1
            col_cat = col_cat if col_cat is not None else 2
            col_manha = col_manha if col_manha is not None else 3
            col_tarde = col_tarde if col_tarde is not None else 4
            
            df_final = pd.DataFrame()
            df_final["SETOR_RAW"] = df_aba.iloc[:, col_setor].ffill()
            df_final["SUB_SETOR_RAW"] = df_aba.iloc[:, col_sub].fillna("")
            df_final["CATEGORIA_RAW"] = df_aba.iloc[:, col_cat]
            
            df_final = df_final[df_final["CATEGORIA_RAW"].notna()]
            df_final = df_final[~df_final["SETOR_RAW"].astype(str).str.upper().str.contains("TOTAL|QUANTITATIVO|HOSPITAL", na=False)]
            df_final = df_final[~df_final["CATEGORIA_RAW"].astype(str).str.upper().str.contains("TOTAL|QUANTITATIVO|HOSPITAL", na=False)]
            
            def limpar_vagas(valor):
                if pd.isna(valor) or str(valor).strip() == "": 
                    return 0
                v_str = "".join(filter(str.isdigit, str(valor)))
                return int(v_str) if v_str != "" else 0
            
            df_final["VAGAS_MANHA"] = df_aba.iloc[:, col_manha].apply(limpar_vagas)
            df_final["VAGAS_TARDE"] = df_aba.iloc[:, col_tarde].apply(limpar_vagas)
            df_final["VAGAS_TOTAL"] = df_final["VAGAS_MANHA"] + df_final["VAGAS_TARDE"]
            
            df_final["LOCAL_COMBINADO"] = df_final.apply(
                lambda r: f"{r['SETOR_RAW']} ➔ {r['SUB_SETOR_RAW']}" if (r['SUB_SETOR_RAW'] != "" and str(r['SETOR_RAW']).upper() != str(r['SUB_SETOR_RAW']).upper()) else f"{r['SETOR_RAW']}",
                axis=1
            )
            df_final["LOCAL_E_PROF"] = df_final["LOCAL_COMBINADO"] + " (" + df_final["CATEGORIA_RAW"].astype(str) + ")"
            
            df_ativas = df_final[df_final["VAGAS_TOTAL"] > 0].copy()
            df_inativas = df_final[df_final["VAGAS_TOTAL"] == 0].copy()
            
            return df_ativas, df_inativas, "SETOR_RAW", "SUB_SETOR_RAW", "CATEGORIA_RAW"

        df_hcid, df_hcid_zero, hc_setor, hc_sub, hc_cat = extrair_e_limpar_dados(aba_hcid_real)
        df_anexo, df_anexo_zero, ax_setor, ax_sub, ax_cat = extrair_e_limpar_dados(aba_anexo_real)
        
    except Exception as e:
        st.error(f"Erro no processamento automático dos dados da planilha: {e}")
        st.stop()
else:
    st.info("💡 Por favor, arraste ou carregue sua planilha Excel para estruturar os painéis automaticamente.")
    st.stop()

# ==============================================================================
# 4. QUADRO 1: INDICADORES EXCLUSIVOS - HCID
# ==============================================================================
st.markdown("<h2 style='color: #008080; border-bottom: 2px solid #008080;'>🏢 QUADRO I - Mapeamento de Vagas Exclusivo HCID</h2>", unsafe_allow_html=True)

if not df_hcid.empty:
    m1, m2 = st.columns(2)
    t_vagas_hcid = int(df_hcid["VAGAS_TOTAL"].sum())
    t_setores_hcid = df_hcid["LOCAL_COMBINADO"].nunique()
    
    m1.metric(label="📊 1. Total de Vagas de Estágio no HCID (Soma Geral)", value=f"{t_vagas_hcid} Vagas")
    m2.metric(label="📍 2. Total de Setores Disponibilizados p/ Campo no HCID", value=t_setores_hcid)
    st.markdown("---")
    
    # LINHA 2: GRÁFICOS DE VISÃO GERAL (3 E 4) LADO A LADO
    c1, c2 = st.columns(2)
    with c1:
        df_g3 = df_hcid.groupby("LOCAL_COMBINADO")["LOCAL_COMBINADO"].count().reset_index(name="Contagem")
        x_3, y_3 = ("Contagem", "LOCAL_COMBINADO") if sub_or == "h" else ("LOCAL_COMBINADO", "Contagem")
        fig3 = px.bar(df_g3, x=x_3, y=y_3, orientation=sub_or, text="Contagem", title="3. Setores Disponibilizados para Estágio (HCID)", color_discrete_sequence=["#008080"])
        fig3.update_layout(yaxis={'categoryorder':'total ascending'} if sub_or == "h" else {'categoryorder':'total descending'}, height=450)
        fig3.update_traces(texttemplate='<b>%{text}</b>', textposition='outside', textfont=dict(size=14))
        st.plotly_chart(fig3, use_container_width=True)
    with c2:
        df_g4 = df_hcid.groupby(["LOCAL_COMBINADO", hc_cat])["VAGAS_TOTAL"].sum().reset_index()
        x_4, y_4 = ("VAGAS_TOTAL", "LOCAL_COMBINADO") if sub_or == "h" else ("LOCAL_COMBINADO", "VAGAS_TOTAL")
        fig4 = px.bar(df_g4, x=x_4, y=y_4, color=hc_cat, orientation=sub_or, barmode="stack", text="VAGAS_TOTAL", title="4. Categorias Profissionais Contempladas por Setor (HCID)", color_discrete_sequence=cor_sequencia)
        fig4.update_layout(yaxis={'categoryorder':'total ascending'} if sub_or == "h" else {'categoryorder':'total descending'}, legend_title_text="Profissão", height=450)
        fig4.update_traces(texttemplate='<b>%{text}</b>', textposition='inside', textfont=dict(size=13))
        st.plotly_chart(fig4, use_container_width=True)
        
