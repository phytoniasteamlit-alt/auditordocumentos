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
# 3. MOTOR DE PROCESSAMENTO DE DADOS BLINDADO (BUSCA DINÂMICA)
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
            
            # Carrega a planilha sem cabeçalho para varrer as linhas e achar onde começam os dados reais
            df_bruto = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
            
            # 1. Localiza a linha correta do cabeçalho
            linha_cabecalho = 0
            for idx, row in df_bruto.iterrows():
                row_str = " ".join([str(x).upper() for x in row.dropna()])
                if "CATEGOR" in row_str or "PROFISS" in row_str or "SETOR" in row_str:
                    linha_cabecalho = idx
                    break
            
            # 2. Descobre dinamicamente os índices das colunas analisando as linhas próximas ao cabeçalho
            cabecalho_linha = df_bruto.iloc[linha_cabecalho]
            sub_cabecalho_linha = df_bruto.iloc[linha_cabecalho+1] if (linha_cabecalho + 1 < len(df_bruto)) else cabecalho_linha
            
            idx_setor, idx_sub, idx_cat = 0,  1,  2
            idx_manha, idx_tarde = 3,  4
            
            # Varre as primeiras linhas para identificar onde estão as palavras chaves das colunas
            for r_idx in range(max(0, linha_cabecalho-2), min(len(df_bruto), linha_cabecalho+3)):
                for c_idx in range(len(df_bruto.columns)):
                    celula = normalizar_texto(df_bruto.iloc[r_idx, c_idx])
                    if "SUB" in celula: idx_sub = c_idx
                    elif "SETOR" in celula or "CAMPO" in celula: idx_setor = c_idx
                    elif "PROF" in celula or "CAT" in celula: idx_cat = c_idx
                    elif "MANH" in celula: idx_manha = c_idx
                    elif "TARD" in celula: idx_tarde = c_idx

            # Corta a tabela mantendo apenas os dados úteis
            df_aba = df_bruto.iloc[linha_cabecalho+1:].copy()
            
            # Monta o DataFrame final limpo
            df_final = pd.DataFrame()
            df_final["SETOR_RAW"] = df_aba.iloc[:, idx_setor].ffill()
            df_final["SUB_SETOR_RAW"] = df_aba.iloc[:, idx_sub].fillna("")
            df_final["CATEGORIA_RAW"] = df_aba.iloc[:, idx_cat]
            
            # Limpa ruídos de linhas vazias e textos de cabeçalho duplicados ou totais fixos do Excel
            df_final = df_final[df_final["CATEGORIA_RAW"].notna()]
            df_final["CATEGORIA_STR"] = df_final["CATEGORIA_RAW"].astype(str).str.strip()
            df_final = df_final[(df_final["CATEGORIA_STR"] != "") & (~df_final["CATEGORIA_STR"].str.upper().contains("CATEGOR|PROFISS|TOTAL", na=False))]
            df_final = df_final[~df_final["SETOR_RAW"].astype(str).str.upper().str.contains("TOTAL|QUANTITATIVO|HOSPITAL", na=False)]
            
            # Limpeza via Expressão Regular (Regex) para converter textos como "4 por turno" em inteiros puros
            def limpar_vagas(valor):
                if pd.isna(valor) or str(valor).strip() == "": 
                    return 0
                v_str = "".join(filter(str.isdigit, str(valor)))
                return int(v_str) if v_str != "" else 0
            
            df_final["VAGAS_MANHA"] = df_aba.loc[df_final.index, idx_manha].apply(limpar_vagas)
            df_final["VAGAS_TARDE"] = df_aba.loc[df_final.index, idx_tarde].apply(limpar_vagas)
            df_final["VAGAS_TOTAL"] = df_final["VAGAS_MANHA"] + df_final["VAGAS_TARDE"]
            
            # Montagem estruturada do identificador combinado setor ➔ subsetor
            df_final["LOCAL_COMBINADO"] = df_final.apply(
                lambda r: f"{r['SETOR_RAW']} ➔ {r['SUB_SETOR_RAW']}" if (r['SUB_SETOR_RAW'] != "" and str(r['SETOR_RAW']).upper() != str(r['SUB_SETOR_RAW']).upper()) else f"{r['SETOR_RAW']}",
                axis=1
            )
            df_final["LOCAL_E_PROF"] = df_final["LOCAL_COMBINADO"] + " (" + df_final["CATEGORIA_RAW"].astype(str) + ")"
            
            # Separação estrutural de dados ativos (com vagas) vs dados inativos (sem vagas alocadas)
            df_ativas = df_final[df_final["VAGAS_TOTAL"] > 0].copy()
            df_inativas = df_final[df_final["VAGAS_TOTAL"] == 0].copy()
            
            return df_ativas, df_inativas, "SETOR_RAW", "SUB_SETOR_RAW", "CATEGORIA_RAW"

        df_hcid, df_hcid_zero, hc_setor, hc_sub, hc_cat = extrair_e_limpar_dados(aba_hcid_real)
        df_anexo, df_anexo_zero, ax_setor, ax_sub, ax_cat = extrair_e_limpar_dados(aba_anexo_real)
        
    except Exception as e:
        st.error(f"Erro no processamento dinâmico da planilha: {e}")
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
            profissoes = ", ".join(df_sub["CATEGORIA_RAW"].unique())
            
            nome_sub = f"na ala/área **{sub}**" if sub != "" else "na área geral"
            sub_detalhes.append(
                f"destas, **{total_sub} estão alocadas** {nome_sub} (composta por: {profissoes}), sendo **{m_sub} no turno da manhã** e **{t_sub} no turno da tarde**"
            )
            
        texto_setor += " Deste montante, " + "; ".join(sub_detalhes) + "."
        textos.append(texto_setor)
    return textos

# ==============================================================================
# 4. QUADRO I - HCID (ORGANIZAÇÃO EM ABAS INTERNAS)
# ==============================================================================
