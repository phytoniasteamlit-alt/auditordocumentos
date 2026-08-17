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
# 3. MOTOR DE PROCESSAMENTO DE DADOS EXECUTIVO (CALIBRADO POR TEXTO)
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
                return pd.DataFrame(), "SETOR_RAW", "SUB_SETOR_RAW", "CATEGORIA_RAW"
            
            df_bruto = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
            
            # Localiza a linha correta das colunas (Onde tem "Setor" ou "Categorias")
            linha_cabecalho = 0
            for idx, row in df_bruto.iterrows():
                row_str = " ".join([str(x).upper() for x in row.dropna()])
                if "CATEGOR" in row_str or "PROFISS" in row_str or "SETOR" in row_str:
                    linha_cabecalho = idx
                    break
            
            # Extrai os cabeçalhos textuais limpos diretamente da linha localizada
            cabecalhos_originais = [str(c).strip().replace('\n', ' ') for c in df_bruto.iloc[linha_cabecalho]]
            
            # Copia os dados abaixo do cabeçalho
            df_aba = df_bruto.iloc[linha_cabecalho+1:].copy()
            df_aba.columns = cabecalhos_originais
            
            # Identificação das colunas dinamicamente por texto purificado
            col_setor, col_sub, col_cat = None, None, None
            col_manha, col_tarde = None, None
            
            for idx_c, col_nome in enumerate(df_aba.columns):
                c_norm = normalizar_texto(col_nome)
                if "SUB" in c_norm:
                    col_sub = idx_c
                elif "SETOR" in c_norm or "CAMPO" in c_norm:
                    col_setor = idx_c
                elif "PROF" in c_norm or "CAT" in c_norm:
                    col_cat = idx_c
                elif "MANH" in c_norm:
                    col_manha = idx_c
                elif "TARD" in c_norm:
                    col_tarde = idx_c

            # Fallbacks baseados em posições caso o texto mude drasticamente
            col_setor = col_setor if col_setor is not None else 0
            col_sub = col_sub if col_sub is not None else 1
            col_cat = col_cat if col_cat is not None else 2
            col_manha = col_manha if col_manha is not None else 3
            col_tarde = col_tarde if col_tarde is not None else 4
            
            # Cria colunas amigáveis padronizadas no DataFrame final
            df_final = pd.DataFrame()
            df_final["SETOR_RAW"] = df_aba.iloc[:, col_setor]
            df_final["SUB_SETOR_RAW"] = df_aba.iloc[:, col_sub] if col_sub < df_aba.shape[1] else ""
            df_final["CATEGORIA_RAW"] = df_aba.iloc[:, col_cat] if col_cat < df_aba.shape[1] else ""
            
            # Preenche os nomes dos setores para baixo (Forward Fill) nas sub-especialidades
            df_final["SETOR_RAW"] = df_final["SETOR_RAW"].ffill()
            df_final["SUB_SETOR_RAW"] = df_final["SUB_SETOR_RAW"].fillna("")
            
            # Elimina linhas nulas nas profissões e marcadores de "TOTAL" do Excel
            df_final = df_final[df_final["CATEGORIA_RAW"].notna()]
            df_final = df_final[~df_final["SETOR_RAW"].astype(str).str.upper().str.contains("TOTAL|QUANTITATIVO|HOSPITAL", na=False)]
            df_final = df_final[~df_final["CATEGORIA_RAW"].astype(str).str.upper().str.contains("TOTAL|QUANTITATIVO|HOSPITAL", na=False)]
            
            # Filtro numérico para limpar expressões textuais tipo "4 por turno"
            def limpar_vagas(valor):
                if pd.isna(valor) or str(valor).strip() == "": 
                    return 0
                v_str = "".join(filter(str.isdigit, str(valor)))
                return int(v_str) if v_str != "" else 0
            
            # Captura os dados numéricos baseados nas colunas corretas de texto encontradas
            df_final["VAGAS_MANHA"] = df_aba.iloc[:, col_manha].apply(limpar_vagas)
            df_final["VAGAS_TARDE"] = df_aba.iloc[:, col_tarde].apply(limpar_vagas)
            df_final["VAGAS_TOTAL"] = df_final["VAGAS_MANHA"] + df_final["VAGAS_TARDE"]
            
            # Montagem estruturada do eixo vertical combinando locais
            df_final["LOCAL_COMBINADO"] = df_final.apply(
                lambda r: f"{r['SETOR_RAW']}" if (not r['SUB_SETOR_RAW'] or str(r['SETOR_RAW']).upper() == str(r['SUB_SETOR_RAW']).upper() or str(r['SUB_SETOR_RAW']).strip() == "") else f"{r['SETOR_RAW']} - {r['SUB_SETOR_RAW']}",
                axis=1
            )
            
            return df_final, "SETOR_RAW", "SUB_SETOR_RAW", "CATEGORIA_RAW"

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
