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

sub_or = "h" if estilo_grafico == "Barras Horizontais" else "v"

# ==============================================================================
# 3. MOTOR DE PROCESSAMENTO DE DADOS TOTALMENTE REESTRUTURADO E BLINDADO
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
            
            # Localiza a linha do cabeçalho estrutural
            linha_cabecalho = 0
            for idx, row in df_bruto.iterrows():
                row_str = " ".join([str(x).upper() for x in row.dropna()])
                if "CATEGOR" in row_str or "PROFISS" in row_str or "SETOR" in row_str:
                    linha_cabecalho = idx
                    break
            
            cabecalhos_originais = [str(c).strip().replace('\n', ' ') for c in df_bruto.iloc[linha_cabecalho]]
            df_aba = df_bruto.iloc[linha_cabecalho+1:].copy()
            df_aba.columns = cabecalhos_originais
            
            # Busca de índices por correspondência textual aproximada
            idx_setor, idx_sub, idx_cat = 0, 1, 2
            idx_manha, idx_tarde = 3, 4
            
            for idx_c, col_nome in enumerate(df_aba.columns):
                c_norm = normalizar_texto(col_nome)
                if "SUB" in c_norm: idx_sub = idx_c
                elif "SETOR" in c_norm or "CAMPO" in c_norm: idx_setor = idx_c
                elif "PROF" in c_norm or "CAT" in c_norm: idx_cat = idx_c
                elif "MANH" in c_norm: idx_manha = idx_c
                elif "TARD" in c_norm: idx_tarde = idx_c  # CORRIGIDO: alterado de col_tarde para idx_tarde

            # Força a conversão das colunas base do Excel para tipos String limpos
            df_final = pd.DataFrame()
            df_final["SETOR_RAW"] = df_aba.iloc[:, idx_setor].astype(str).str.strip().ffill()
            df_final["SUB_SETOR_RAW"] = df_aba.iloc[:, idx_sub].fillna("").astype(str).str.strip()
            df_final["CATEGORIA_RAW"] = df_aba.iloc[:, idx_cat].fillna("").astype(str).str.strip()
            
            # --- PROTEÇÃO ABSOLUTA CONTRA AMBIGUIDADE (Filtro por loops nativos) ---
            # Remove linhas em branco que ficaram perdidas no Excel
            df_final = df_final[(df_final["CATEGORIA_RAW"] != "") & (df_final["CATEGORIA_RAW"] != "nan")]
            
            # Nova filtragem sem utilizar strings dinâmicas no Pandas (Resolve o erro vermelho)
            linhas_validas = []
            for idx_r, row_f in df_final.iterrows():
                setor_upper = str(row_f["SETOR_RAW"]).upper()
                cat_upper = str(row_f["CATEGORIA_RAW"]).upper()
                
                # Ignora linhas que contenham ruídos do Excel como subtotais ou títulos duplicados
                if "TOTAL" in setor_upper or "QUANTITATIVO" in setor_upper or "HOSPITAL" in setor_upper:
                    linhas_validas.append(False)
                elif "TOTAL" in cat_upper or "CATEGOR" in cat_upper or "PROFISS" in cat_upper:
                    linhas_validas.append(False)
                else:
                    linhas_validas.append(True)
                    
            df_final = df_final[linhas_validas].copy()
            
            # Limpeza das expressões textuais das vagas ('4 por turno' -> 4)
            def limpar_vagas(valor):
                if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan": 
                    return None  # Retorna None para aplicar o ffill do setor pai nas sub-especialidades
                v_str = "".join(filter(str.isdigit, str(valor)))
                return int(v_str) if v_str != "" else 0
            
            # Vincula e limpa os turnos coletando as vagas originais do Excel
            df_final["VAGAS_MANHA"] = df_aba.loc[df_final.index, df_aba.columns[idx_manha]].apply(limpar_vagas)
            df_final["VAGAS_TARDE"] = df_aba.loc[df_final.index, df_aba.columns[idx_tarde]].apply(limpar_vagas)
            
            # Preenchimento em cascata para herdar as vagas se a linha de baixo estiver em branco
            df_final["VAGAS_MANHA"] = df_final["VAGAS_MANHA"].ffill().fillna(0).astype(int)
            df_final["VAGAS_TARDE"] = df_final["VAGAS_TARDE"].ffill().fillna(0).astype(int)
            df_final["VAGAS_TOTAL"] = df_final["VAGAS_MANHA"] + df_final["VAGAS_TARDE"]
            
            # Construção elegante do identificador combinado setor ➔ subsetor
            df_final["LOCAL_COMBINADO"] = df_final.apply(
                lambda r: f"{r['SETOR_RAW']} ➔ {r['SUB_SETOR_RAW']}" if (r['SUB_SETOR_RAW'] != "" and r['SETOR_RAW'].upper() != r['SUB_SETOR_RAW'].upper()) else f"{r['SETOR_RAW']}",
                axis=1
            )
            df_final["LOCAL_E_PROF"] = df_final["LOCAL_COMBINADO"] + " (" + df_final["CATEGORIA_RAW"] + ")"
            
            # Divide os dataframes entre ativos (com vagas) e inativos (vagas zeradas)
            df_ativas = df_final[df_final["VAGAS_TOTAL"] > 0].copy()
            df_inativas = df_final[df_final["VAGAS_TOTAL"] == 0].copy()
            
            return df_ativas, df_inativas, "SETOR_RAW", "SUB_SETOR_RAW", "CATEGORIA_RAW"

        df_hcid, df_hcid_zero, hc_setor, hc_sub, hc_cat = extrair_e_limpar_dados(aba_hcid_real)
        df_anexo, df_anexo_zero, ax_setor, ax_sub, ax_cat = extrair_e_limpar_dados(aba_anexo_real)
        
    except Exception as e:
        st.error(f"Erro inesperado durante a leitura dos dados: {e}")
        st.stop()
else:
    st.info("💡 Por favor, arraste ou carregue sua planilha Excel para estruturar os painéis automaticamente.")
    st.stop()

# ==============================================================================
# FUNÇÃO AUXILIAR: GERA O TEXTO DE DESTRINCHAMENTO AUTOMÁTICO
# ==============================================================================
def gerar_texto_distribuicao(df_filtrado):
    textos = []
    if df_filtrado.empty:
        return textos
    setores_unicos = df_filtrado["SETOR_RAW"].unique()
    for s in setores_unicos:
        df_s = df_filtrado[df_filtrado["SETOR_RAW"] == s]
        total_s = df_s["VAGAS_TOTAL"].sum()
        
        texto_setor = f"📌 **Setor {s}**: Disponibiliza um total de **{total_s} vagas** para campo de estágio. "
        sub_detalhes = []
        
        sub_unicos = df_s["SUB_SETOR_RAW"].unique()
        for sub in sub_unicos:
            df_sub = df_s[df_s["SUB_SETOR_RAW"] == sub]
            total_sub = df_sub["VAGAS_TOTAL"].sum()
            m_sub = df_sub["VAGAS_MANHA"].sum()
            t_sub = df_sub["VAGAS_TARDE"].sum()
