import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA (Interface Dashboard Executivo)
# ==============================================================================
st.set_page_config(
    page_title="Painel de Estágios Executivo - HCID",
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
# 3. MOTOR DE PROCESSAMENTO MATRICIAL (CORREÇÃO DE SINAL DE TURNOS)
# ==============================================================================
if uploaded_file is not None:
    excel_file = pd.ExcelFile(uploaded_file)
    abas_disponiveis = excel_file.sheet_names
    aba_alvo = "HCID_BDD" if "HCID_BDD" in abas_disponiveis else ("HCID" if "HCID" in abas_disponiveis else abas_disponiveis)
    
    df_raw = pd.read_excel(uploaded_file, sheet_name=aba_alvo, header=None)
    
    # CORREÇÃO CRÍTICA: Aplica ffill() horizontal em TODOS os cabeçalhos para não deixar turnos vazios se perderem
    linha_meses = pd.Series(df_raw.iloc[3, :]).ffill().fillna("").astype(str).tolist()
    linha_dias = pd.Series(df_raw.iloc[4, :]).ffill().fillna("").astype(str).tolist()
    linha_turnos = pd.Series(df_raw.iloc[5, :]).ffill().fillna("").astype(str).tolist() # Adicionado ffill() aqui
    
    # Isola o corpo de dados reais (A partir da Linha 8 física / índice 7)
    df_corpo = df_raw.iloc[7:].copy().reset_index(drop=True)
    
    # Preenchimento em cascata vertical das colunas estruturais mescladas
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
            
            # Se a célula estiver zerada mas possuir registro ativo, tenta herdar do cabeçalho D ou E
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
                        
                    # Mapeamento do Turno corrigido baseado no cabeçalho ffill horizontal
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
                    
    df_master = pd.DataFrame(registros_vagas)
    
    # ==============================================================================
    # 4. EXIBIÇÃO EM CAIXAS DE TEXTO DESTACADAS (MÉTRICAS EXECUTIVAS)
    # ==============================================================================
    if not df_master.empty:
        st.markdown("### 📋 Resumo Executivo de Capacidade (HCID)")
        
        total_vagas_geral = df_master["VAGAS"].sum()
        total_setores = df_master["SETOR"].nunique()
        total_m_geral = df_master[df_master["TURNO"] == "MANHÃ"]["VAGAS"].sum()
        total_t_geral = df_master[df_master["TURNO"] == "TARDE"]["VAGAS"].sum()
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total de vagas de estágio geral HCID", f"{total_vagas_geral} Vagas")
        c2.metric("Total de setores disponibilizados p/ campo de estágio no HCID", f"{total_setores} Setores")
        c3.metric("Total de vagas de estágio do HCID por turno manhã", f"{total_m_geral} M")
        c4.metric("Total de vagas de estágio do HCID tarde", f"{total_t_geral} T")
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("##### 🗓️ Total de estagiários por dia")
        
        col_seletor, _ = st.columns(2)
        dias_disponiveis = sorted(df_master["DIA_SEMANA"].unique())
        dia_selecionado = col_seletor.selectbox("Selecione o Dia da Semana para Filtrar os Totais:", dias_disponiveis)
        
        vagas_dia_m = df_master[(df_master["DIA_SEMANA"] == dia_selecionado) & (df_master["TURNO"] == "MANHÃ")]["VAGAS"].sum()
        vagas_dia_t = df_master[(df_master["DIA_SEMANA"] == dia_selecionado) & (df_master["TURNO"] == "TARDE")]["VAGAS"].sum()
        
        cc1, cc2 = st.columns(2)
        cc1.info(f"🟢 **Total de estagiários por dia turno manhã ({dia_selecionado}):** {vagas_dia_m} alunos em campo.")
        cc2.warning(f"🟠 **Total de estagiários por dia turno tarde ({dia_selecionado}):** {vagas_dia_t} alunos em campo.")
        
        st.markdown("---")
        
        # ==============================================================================
        # 5. RENDERIZAÇÃO DOS GRÁFICOS INTERATIVOS EXECUTIVOS
        # ==============================================================================
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.markdown("### 🏢 Setores disponibilizados para realização de estágio no hcid")
            df_g1 = df_master.groupby("SETOR")["VAGAS"].sum().reset_index().sort_values(by="VAGAS", ascending=True)
            
            fig1 = px.bar(
                df_g1, x="VAGAS", y="SETOR", orientation="h",
                color="VAGAS", color_continuous_scale=px.colors.sequential.Tealgrn, text_auto=True
            )
            fig1.update_layout(showlegend=False, height=400, margin=dict(l=20, r=35, t=10, b=10))
            fig1.update_traces(textposition="outside", cliponaxis=False)
            st.plotly_chart(fig1, use_container_width=True)
            
        with col_g2:
            st.markdown("### 👩‍⚕️ Categorias profissionais contemplados no estagio por setor no hcid")
            setor_selecionado_g2 = st.selectbox("Escolha o Setor para Filtrar as Profissões:", sorted(df_master["SETOR"].unique()))
            df_g2 = df_master[df_master["SETOR"] == setor_selecionado_g2].groupby("CATEGORIA")["VAGAS"].sum().reset_index().sort_values(by="VAGAS", ascending=True)
            
            fig2 = px.bar(
                df_g2, x="VAGAS", y="CATEGORIA", orientation="h",
                color="VAGAS", color_continuous_scale=px.colors.sequential.Bluered, text_auto=True
            )
            fig2.update_layout(showlegend=False, height=335, margin=dict(l=20, r=35, t=10, b=10))
            fig2.update_traces(textposition="outside", cliponaxis=False)
            st.plotly_chart(fig2, use_container_width=True)
            
        # --- GRÁFICO CRONOLÓGICO MENSAL REQUISITADO ---
        st.markdown("---")
        st.markdown("### 📅 Distribuição Mensal Organizada de Vagas Ocupadas por Turno")
