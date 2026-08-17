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
                return pd.DataFrame(), "SETOR", "SUB_SETOR", "CATEGORIA"
            
            # Carrega a planilha sem cabeçalho para fazer a varredura manual
            df_bruto = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
            
            # Localiza a linha correta das colunas (Ex: linha onde tem "Setor" ou "Categorias")
            linha_cabecalho = 0
            for idx, row in df_bruto.iterrows():
                row_str = " ".join([str(x).upper() for x in row.dropna()])
                if "CATEGOR" in row_str or "PROFISS" in row_str or "SETOR" in row_str:
                    linha_cabecalho = idx
                    break
            
            # Corta a tabela a partir da linha localizada
            df_aba = df_bruto.iloc[linha_cabecalho+1:].copy()
            valores_cabecalho = [str(c).strip().replace('\n', ' ') for c in df_bruto.iloc[linha_cabecalho]]
            
            # Tratamento de colunas com mesmo nome vindas do Excel
            novas_colunas = []
            ultima_coluna_valida = "COLUNA"
            for col in valores_cabecalho:
                if pd.isna(col) or "UNNAMED" in str(col).upper() or str(col).strip() == "":
                    novas_colunas.append(ultima_coluna_valida)
                else:
                    ultima_coluna_valida = str(col).strip()
                    novas_colunas.append(ultima_coluna_valida)
            df_aba.columns = novas_colunas
            
            # Renomeia por posição física para garantir segurança total contra erros de digitação
            df_aba = df_aba.rename(columns={
                df_aba.columns[0]: "SETOR_RAW",
                df_aba.columns[1]: "SUB_SETOR_RAW",
                df_aba.columns[2]: "CATEGORIA_RAW"
            })
            
            # Preenche os nomes dos setores para as linhas de especialidades de baixo
            df_aba["SETOR_RAW"] = df_aba["SETOR_RAW"].ffill()
            df_aba["SUB_SETOR_RAW"] = df_aba["SUB_SETOR_RAW"].fillna("")
            
            # Descarta linhas completamente vazias ou que sejam separadores de 'TOTAL' do Excel
            df_aba = df_aba[df_aba["CATEGORIA_RAW"].notna()]
            df_aba = df_aba[~df_aba["SETOR_RAW"].astype(str).str.upper().str.contains("TOTAL|QUANTITATIVO|HOSPITAL", na=False)]
            df_aba = df_aba[~df_aba["CATEGORIA_RAW"].astype(str).str.upper().str.contains("TOTAL|QUANTITATIVO|HOSPITAL", na=False)]
            
            # Função limpa texto isolando apenas dígitos
            def limpar_vagas(valor):
                if pd.isna(valor) or str(valor).strip() == "": 
                    return 0
                v_str = "".join(filter(str.isdigit, str(valor)))
                return int(v_str) if v_str != "" else 0
            
            # Captura Manhã (Coluna índice 3) e Tarde (Coluna índice 4)
            df_aba["VAGAS_MANHA"] = df_aba.iloc[:, 3].apply(limpar_vagas)
            df_aba["VAGAS_TARDE"] = df_aba.iloc[:, 4].apply(limpar_vagas)
            df_aba["VAGAS_TOTAL"] = df_aba["VAGAS_MANHA"] + df_aba["VAGAS_TARDE"]
            
            # Monta o nome combinado estruturado para os eixos dos gráficos
            df_aba["LOCAL_COMBINADO"] = df_aba.apply(
                lambda r: f"{r['SETOR_RAW']}" if (not r['SUB_SETOR_RAW'] or str(r['SETOR_RAW']).upper() == str(r['SUB_SETOR_RAW']).upper() or str(r['SUB_SETOR_RAW']).strip() == "") else f"{r['SETOR_RAW']} - {r['SUB_SETOR_RAW']}",
                axis=1
            )
            
            return df_aba, "SETOR_RAW", "SUB_SETOR_RAW", "CATEGORIA_RAW"

        df_hcid, hc_setor, hc_sub, hc_cat = extrair_e_limpar_dados(aba_hcid_real)
        df_anexo, ax_setor, ax_sub, ax_cat = extrair_e_limpar_dados(aba_anexo_real)
        
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
    
    m1.metric(label="📊 1. Total de Vagas de Estágio no HCID", value=f"{t_vagas_hcid} Vagas")
    m2.metric(label="📍 2. Total de Setores Disponibilizados p/ Campo no HCID", value=t_setores_hcid)
    
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    
    with c1:
        df_g3 = df_hcid.groupby("LOCAL_COMBINADO")["LOCAL_COMBINADO"].count().reset_index(name="Contagem")
        fig3 = px.bar(df_g3, x="Contagem", y="LOCAL_COMBINADO", orientation="h", title="3. Setores Disponibilizados para Estágio (HCID)", color_discrete_sequence=["#008080"])
        fig3.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig3, use_container_width=True)
        
        df_g5 = df_hcid.groupby("LOCAL_COMBINADO")["VAGAS_TOTAL"].sum().reset_index()
        fig5 = px.bar(df_g5, x="VAGAS_TOTAL", y="LOCAL_COMBINADO", orientation="h", title="5. Total de Vagas Disponibilizadas por Setor (HCID)", color_discrete_sequence=["#4682B4"])
        fig5.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig5, use_container_width=True)

    with c2:
        df_g4 = df_hcid.groupby(["LOCAL_COMBINADO", hc_cat])["VAGAS_TOTAL"].sum().reset_index()
        fig4 = px.bar(df_g4, x="VAGAS_TOTAL", y="LOCAL_COMBINADO", color=hc_cat, barmode="stack", title="4. Categorias Profissionais Contempladas por Setor (HCID)", color_discrete_sequence=cor_sequencia)
        fig4.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig4, use_container_width=True)
        
        df_g6 = pd.DataFrame({
            "Turno": ["Manhã", "Tarde"],
            "Vagas": [df_hcid["VAGAS_MANHA"].sum(), df_hcid["VAGAS_TARDE"].sum()]
        })
        fig6 = px.bar(df_g6, x="Turno", y="Vagas", text="Vagas", title="6. Total de Vagas de Estágio por Turno (HCID)", color="Turno", color_discrete_map={"Manhã": "#4682B4", "Tarde": "#FF8C00"})
        st.plotly_chart(fig6, use_container_width=True)
        
    st.markdown("#### 📅 7. Total Estagiários por Turno por Dia (HCID)")
    df_g7 = df_hcid.groupby(hc_cat)[["VAGAS_MANHA", "VAGAS_TARDE"]].sum().reset_index()
    df_g7_melt = df_g7.melt(id_vars=hc_cat, value_vars=["VAGAS_MANHA", "VAGAS_TARDE"], var_name="Turno", value_name="Vagas")
