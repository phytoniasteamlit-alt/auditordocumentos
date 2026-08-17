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
# 3. MOTOR DE PROCESSAMENTO CORRIGIDO PARA ESTRUTURA MATRICIAL
# ==============================================================================
if uploaded_file is not None:
    excel_file = pd.ExcelFile(uploaded_file)
    abas_disponiveis = excel_file.sheet_names
    aba_alvo = "HCID_BDD" if "HCID_BDD" in abas_disponiveis else ("HCID" if "HCID" in abas_disponiveis else abas_disponiveis)
    
    # Carrega a tabela bruta sem cabeçalho automático
    df_raw = pd.read_excel(uploaded_file, sheet_name=aba_alvo, header=None)
    
    # Processamento dos cabeçalhos horizontais (Meses na linha 3, Dias na linha 4, Turnos na linha 5)
    # Extrai cada linha bruta como uma série e propaga as células mescladas na horizontal antes de converter para lista
    linha_meses = df_raw.iloc[2].ffill().fillna("").astype(str).tolist()
    linha_dias = df_raw.iloc[3].ffill().fillna("").astype(str).tolist()
    linha_turnos = df_raw.iloc[4].fillna("").astype(str).tolist()
    
    # Isola o corpo de dados reais (A partir da Linha 8 física / índice 7)
    df_corpo = df_raw.iloc[7:].copy()
    
    # Preenchimento em cascata vertical das colunas estruturais mescladas (Setor, Sub-setor, Categoria)
    df_corpo.iloc[:, 0] = df_corpo.iloc[:, 0].replace(["nan", "NAN", ""], pd.NA).ffill().fillna("GERAL")
    df_corpo.iloc[:, 1] = df_corpo.iloc[:, 1].replace(["nan", "NAN", ""], pd.NA).ffill().fillna("GERAL")
    df_corpo.iloc[:, 2] = df_corpo.iloc[:, 2].replace(["nan", "NAN", ""], pd.NA).ffill().fillna("NÃO ESPECIFICADO")
    
    registros_vagas = []
    
    # Varredura matricial convertendo a planilha em banco de dados linear
    for _, row in df_corpo.iterrows():
        # CORREÇÃO: Utiliza o fatiamento correto por índice de posição física da série
        setor = str(row.iloc[0]).strip().upper()
        sub_setor = str(row.iloc[1]).strip().upper()
        categoria = str(row.iloc[2]).strip().upper()
        
        # Filtro rígido para descartar cabeçalhos duplicados, linhas de lixo ou subtotais
        if "TOTAL" in setor or "TOTAL" in categoria or categoria == "" or setor == "SETOR" or setor == "GERAL":
            continue
            
        # Percorre as colunas de calendário a partir da coluna indexada 8 (Onde começam as vagas de Agosto)
        for col_idx in range(8, len(row)):
            vaga_bruta = row.iloc[col_idx]
            qtd_vagas = extrair_numero(vaga_bruta)
            
            if qtd_vagas > 0:
                mes_nome = str(linha_meses[col_idx]).strip().upper()
                dia_nome = str(linha_dias[col_idx]).strip().upper()
                turno_nome = str(linha_turnos[col_idx]).strip().upper()
                
                # Validação se a coluna pertence à matriz de agendamentos temporais
                if any(m in mes_nome for m in ["AGO", "SET", "OUT", "NOV", "DEZ"]) or "VAGAS" in mes_nome:
                    if "VAGAS" in mes_nome or mes_nome == "": 
                        mes_nome = "AGOSTO"
                    if "MANH" not in turno_nome and "TARD" not in turno_nome:
                        turno_nome = "MANHÃ" if "MANH" in dia_nome else "TARDE"
                        
                    registros_vagas.append({
                        "SETOR": setor,
                        "SUB_SETOR": sub_setor,
                        "CATEGORIA": categoria,
                        "MÊS": mes_nome,
                        "DIA_SEMANA": dia_nome if any(d in dia_nome for d in ["SEG", "TER", "QUA", "QUI", "SEX"]) else "SEGUNDA",
                        "TURNO": "MANHÃ" if "MANH" in turno_nome else "TARDE",
                        "VAGAS": qtd_vagas
                    })
                    
    df_master = pd.DataFrame(registros_vagas)
    
    # ==============================================================================
    # 4. EXIBIÇÃO EM CAIXAS DE TEXTO DESTACADAS (MÉTRICAS EXECUTIVAS)
    # ==============================================================================
    if not df_master.empty:
        st.markdown("### 📋 Resumo Executivo de Capacidade (HCID)")
        
        # Cálculos consolidados para as caixas de texto pedidas
        total_vagas_geral = df_master["VAGAS"].sum()
        total_setores = df_master["SETOR"].nunique()
        total_m_geral = df_master[df_master["TURNO"] == "MANHÃ"]["VAGAS"].sum()
        total_t_geral = df_master[df_master["TURNO"] == "TARDE"]["VAGAS"].sum()
        
        # Primeira linha de Caixas de Texto (Métricas de Vagas Globais)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total de vagas de estágio geral HCID", f"{total_vagas_geral} Vagas")
        c2.metric("Total de setores disponibilizados p/ campo de estágio no HCID", f"{total_setores} Setores")
        c3.metric("Total de vagas de estágio do HCID por turno manhã", f"{total_m_geral} M")
        c4.metric("Total de vagas de estágio do HCID tarde", f"{total_t_geral} T")
        
        # Segunda linha de Caixas de Texto (Auditoria de Alunos por Dia de Semana)
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
        # 5. RENDERIZAÇÃO DOS GRÁFICOS INTERATIVOS SOLICITADOS
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
            
        # --- GRÁFICO CRONOLÓGICO POR MÊS SOLICITADO ---
        st.markdown("---")
