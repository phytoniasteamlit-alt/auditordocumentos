import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA
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
# 2. SIDEBAR
# ==============================================================================
st.sidebar.header("⚙️ Painel de Controle")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha de Estágios (.xlsx):", type=["xlsx"])

# ==============================================================================
# 3. MOTOR DE PROCESSAMENTO DE CALENDÁRIO MATRICIAL (FLUXO LINEAR SEGURO)
# ==============================================================================
if uploaded_file is not None:
    excel_file = pd.ExcelFile(uploaded_file)
    abas_disponiveis = excel_file.sheet_names
    aba_alvo = "HCID_BDD" if "HCID_BDD" in abas_disponiveis else ("HCID" if "HCID" in abas_disponiveis else abas_disponiveis[0])
    
    # Carrega a planilha sem cabeçalho automático
    df_raw = pd.read_excel(uploaded_file, sheet_name=aba_alvo, header=None)
    
    # Identifica as linhas estruturais do cabeçalho complexo (Linhas 3, 4 e 5)
    # Correção do preenchimento: ffill() direto nas séries para evitar deprecation warnings
    linha_meses = df_raw.iloc[2].ffill().fillna("").astype(str).tolist()
    linha_dias = df_raw.iloc[3].ffill().fillna("").astype(str).tolist()
    linha_turnos = df_raw.iloc[4].fillna("").astype(str).tolist()
    
    # Corpo dos dados (Linha 8 física em diante)
    df_corpo = df_raw.iloc[7:].copy()
    
    # Preenchimento em cascata das colunas estruturais mescladas (Setor, Sub-setor, Categoria)
    df_corpo[0] = df_corpo[0].replace(["nan", "NAN", ""], pd.NA).ffill().fillna("GERAL")
    df_corpo[1] = df_corpo[1].replace(["nan", "NAN", ""], pd.NA).ffill().fillna("GERAL")
    df_corpo[2] = df_corpo[2].replace(["nan", "NAN", ""], pd.NA).ffill().fillna("NÃO ESPECIFICADO")
    
    # Lista para consolidar todas as vagas encontradas na matriz para o formato longo
    registros_vagas = []
    
    for idx_row, row in df_corpo.iterrows():
        setor = str(row[0]).strip()
        sub_setor = str(row[1]).strip()
        categoria = str(row[2]).strip()
        
        # Ignora linhas de cabeçalho residual ou totais da planilha
        if "TOTAL" in setor.upper() or "TOTAL" in categoria.upper() or categoria == "" or setor == "GERAL":
            continue
            
        # Varre as colunas de dados a partir da coluna de índice 3 (Coluna D do Excel)
        for col_idx in range(3, len(row)):
            vaga_bruta = row[col_idx]
            qtd_vagas = extrair_numero(vaga_bruta)
            
            if qtd_vagas > 0:
                mes = str(linha_meses[col_idx]).strip().upper()
                dia = str(linha_dias[col_idx]).strip().upper()
                turno = str(linha_turnos[col_idx]).strip().upper()
                
                # Filtro de segurança para validar se a coluna realmente pertence ao calendário
                if any(m in mes for m in ["AGO", "SET", "OUT", "NOV", "DEZ"]) or "VAGAS" in mes:
                    if "VAGAS" in mes or mes == "": 
                        mes = "AGOSTO" 
                    if "MANH" not in turno and "TARD" not in turno: 
                        turno = "MANHÃ" if "MANH" in dia else "TARDE"
                    
                    registros_vagas.append({
                        "SETOR": setor,
                        "SUB_SETOR": sub_setor,
                        "CATEGORIA": categoria,
                        "MÊS": mes,
                        "DIA_SEMANA": dia if any(d in dia for d in ["SEG", "TER", "QUA", "QUI", "SEX"]) else "SEGUNDA",
                        "TURNO": "MANHÃ" if "MANH" in turno else "TARDE",
                        "VAGAS": qtd_vagas
                    })
                    
    df_master = pd.DataFrame(registros_vagas)
    
    # ==============================================================================
    # 4. EXIBIÇÃO DAS CAIXAS DE TEXTO (MÉTRICAS / CARDS SOLICITADOS)
    # ==============================================================================
    if not df_master.empty:
        st.markdown("### 📋 Resumo Executivo de Capacidade (HCID)")
        
        # Cálculos macro unificados
        total_vagas_geral = df_master["VAGAS"].sum()
        total_setores = df_master["SETOR"].nunique()
        total_manha = df_master[df_master["TURNO"] == "MANHÃ"]["VAGAS"].sum()
        total_tarde = df_master[df_master["TURNO"] == "TARDE"]["VAGAS"].sum()
        
        # Primeira linha de caixas de texto (Métricas Principais)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total de Vagas de Estágio Geral", f"{total_vagas_geral} Vagas")
        c2.metric("Setores para Campo de Estágio", f"{total_setores} Setores Ativos")
        c3.metric("Total Turno Manhã (Consolidado)", f"{total_manha} M")
        c4.metric("Total Turno Tarde (Consolidado)", f"{total_tarde} T")
        
        # Segunda linha de caixas de texto (Média Diária Proposta)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("##### 🗓️ Distribuição Diária de Estagiários")
        
        c_dia_m, c_dia_t = st.columns(2)
        
        # Cria um seletor interativo para os dias da semana
        dia_selecionado = c_dia_m.selectbox("Selecione o Dia da Semana para Auditar:", sorted(df_master["DIA_SEMANA"].unique()))
        
        vagas_dia_m = df_master[(df_master["DIA_SEMANA"] == dia_selecionado) & (df_master["TURNO"] == "MANHÃ")]["VAGAS"].sum()
        vagas_dia_t = df_master[(df_master["DIA_SEMANA"] == dia_selecionado) & (df_master["TURNO"] == "TARDE")]["VAGAS"].sum()
        
        cc1, cc2 = st.columns(2)
        cc1.info(f"🟢 **Total de Estagiários na {dia_selecionado} (Turno Manhã):** {vagas_dia_m} alunos em campo.")
        cc2.warning(f"🟠 **Total de Estagiários na {dia_selecionado} (Turno Tarde):** {vagas_dia_t} alunos em campo.")

        st.markdown("---")
        
        # ==============================================================================
        # 5. RENDERIZAÇÃO DOS GRÁFICOS SOLICITADOS
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
            setor_filtro_g2 = st.selectbox("Escolha o Setor para Analisar as Profissões:", sorted(df_master["SETOR"].unique()))
            df_g2 = df_master[df_master["SETOR"] == setor_filtro_g2].groupby("CATEGORIA")["VAGAS"].sum().reset_index().sort_values(by="VAGAS", ascending=True)
            
            fig2 = px.bar(
                df_g2, x="VAGAS", y="CATEGORIA", orientation="h",
                color="VAGAS", color_continuous_scale=px.colors.sequential.Bluered, text_auto=True
            )
            fig2.update_layout(showlegend=False, height=335, margin=dict(l=20, r=35, t=10, b=10))
            fig2.update_traces(textposition="outside", cliponaxis=False)
            st.plotly_chart(fig2, use_container_width=True)

        # --- GRÁFICO POR MÊS REESTRUTURADO ---
        st.markdown("---")
        st.markdown("### 📅 Distribuição Mensal Organizada de Vagas Ocupadas por Turno")
        st.caption("Separação por mês indicando a quantidade de alunos de manhã e tarde e quais setores ocupam essas vagas.")
        
        df_mensal = df_master.groupby(["MÊS", "TURNO", "SETOR"])["VAGAS"].sum().reset_index()
        
        ordem_meses = ["AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
        df_mensal["MÊS"] = pd.Categorical(df_mensal["MÊS"], categories=ordem_meses, ordered=True)
        df_mensal = df_mensal.dropna(subset=["MÊS"]).sort_values(by=["MÊS", "TURNO"])
        
