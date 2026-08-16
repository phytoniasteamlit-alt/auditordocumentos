import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Configuração da página profissional (Layout Expandido)
st.set_page_config(
    page_title="Painel de Indicadores Hospitalares",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Assinatura do Programador Fixada no Canto Inferior Esquerdo
st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        left: 15px;
        bottom: 15px;
        font-size: 14px;
        font-weight: bold;
        color: #888888;
        z-index: 9999;
        background-color: rgba(14, 17, 23, 0.9);
        padding: 5px 10px;
        border-radius: 5px;
    }
    </style>
    <div class="footer">👨‍💻 Ezequias S. Santos Naqh / Nsp</div>
    """,
    unsafe_allow_html=True
)

# 3. Menu Lateral de Controle
with st.sidebar:
    st.markdown("## ⚙️ Painel de Controle")
    st.write("Carregar Escala Hospitalar Completa (.xlsx):")
    uploaded_file = st.file_uploader("Upload", type=["xlsx"], label_visibility="collapsed")
    
    st.markdown("---")
    st.markdown("### 🎨 Personalização Visual")
    cor_tema = st.selectbox("Paleta de Cores Pastéis:", ["Azul Soft", "Candy Menta", "Sunset Pastel"])
    
    # Definição das paletas de tons pastéis personalizadas
    if cor_tema == "Azul Soft":
        paleta_pasteis = ["#A0C4FF", "#BDB2FF", "#FFC6FF", "#CAFFBF", "#FDFFB6", "#FFADAD"]
    elif cor_tema == "Candy Menta":
        paleta_pasteis = ["#98D8C8", "#F3B0C3", "#FFD8B3", "#C1C6FC", "#F4F7BB", "#E2C2FF"]
    else:
        paleta_pasteis = ["#FFB7B2", "#FFDAC1", "#E2F0CB", "#B5EAD7", "#C7CEEA", "#FFC6FF"]

# 4. Cabeçalho Principal Customizado (NEP, NEPEX e Setor)
col_titulo, col_info = st.columns([1.2, 1.8])
with col_titulo:
    st.markdown("# 📊 Painel de Indicadores")

with col_info:
    st.markdown(
        """
        <div style="text-align: right; padding-top: 5px; border-left: 2px solid #333; padding-left: 20px;">
            <span style="font-size: 22px; font-weight: bold;">🏥 Hospital da Cidade</span><br>
            <span style="font-size: 16px; font-weight: bold; color: #4EA8DE;">👩‍💼 Coord: Verônica Azevedo</span><br>
            <span style="font-size: 14px; color: #aaaaaa;">📋 Coordenadora do NEP e do NEPEX</span><br>
            <span style="font-size: 13px; color: #888888;">🔬 Setor de Ensino e Pesquisa</span>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

# =========================================================================
# ⚙️ PROCESSADOR INTELIGENTE DE MATRIZ HOSPITALAR (MÁGICA DO PYTHON)
# =========================================================================
def processar_escala_complexa(df_raw):
    df_raw.iloc[:, 0] = df_raw.iloc[:, 0].ffill()
    df_raw.iloc[:, 1] = df_raw.iloc[:, 1].ffill()
    df_raw.iloc[:, 2] = df_raw.iloc[:, 2].ffill()
    
    linha_mes = df_raw.iloc.ffill()
    linha_dia = df_raw.iloc.ffill()
    linha_turno = df_raw.iloc
    
    dados_estruturados = []
    
    for idx_linha in range(7, len(df_raw)):
        linha_atual = df_raw.iloc[idx_linha]
        
        setor_macro = str(linha_atual.iloc[0]).strip()
        sub_setor = str(linha_atual.iloc[1]).strip()
        categoria = str(linha_atual.iloc[2]).strip()
        
        if sub_setor and sub_setor != "None" and sub_setor != "nan" and sub_setor != setor_macro:
            setor_final = f"{setor_macro} - {sub_setor}"
        else:
            setor_final = setor_macro
            
        if any(termo in setor_final.lower() for termo in ["total", "quantitativo", "hospital", "setor"]):
            continue
            
        for idx_col in range(8, len(df_raw.columns)):
            valor_vaga = linha_atual.iloc[idx_col]
            
            if pd.notna(valor_vaga) and (isinstance(valor_vaga, (int, float)) or str(valor_vaga).isdigit()):
                qtd_vagas = int(valor_vaga)
                if qtd_vagas <= 0:
                    continue
                    
                mes = str(linha_mes.iloc[4, idx_col]).strip().upper()
                dia = str(linha_dia.iloc[5, idx_col]).strip().capitalize()
                turno = str(linha_turno.iloc[6, idx_col]).strip()
                
                if mes == "NONE" or "nan" in mes.lower() or not mes: continue
                if dia == "None" or "nan" in dia.lower() or not dia: continue
                
                if "manh" in turno.lower(): turno = "Manhã"
                elif "tard" in turno.lower(): turno = "Tarde"
                elif "noit" in turno.lower(): turno = "Noite"
                else: continue
                
                dados_estruturados.append({
                    "Setor": setor_final,
                    "Categoria Profissional": categoria,
                    "Mês": mes,
                    "Dia da Semana": dia,
                    "Turno": turno,
                    "Vagas Ocupadas": qtd_vagas
                })
                
    return pd.DataFrame(dados_estruturados)

# 5. Renderização e Execução do Upload
if uploaded_file is not None:
    aba_hcid, aba_anexo = st.tabs(["🏢 UNIDADE HCID", "📑 UNIDADE ANEXO"])
    
    # --- BLOCO VISUAL: HCID ---
    with aba_hcid:
        df_raw_hcid = pd.read_excel(uploaded_file, sheet_name="HCID", header=None)
        df_filtro_hcid = processar_escala_complexa(df_raw_hcid)
        
        if not df_filtro_hcid.empty:
            meses_hcid = [m for m in df_filtro_hcid["Mês"].unique() if m and m != "NAN"]
            mes_sel_hcid = st.multiselect("Selecione os Meses para Análise (HCID):", meses_hcid, default=meses_hcid)
            df_final_hcid = df_filtro_hcid[df_filtro_hcid["Mês"].isin(mes_sel_hcid)]
            
            total_geral_vagas = df_final_hcid["Vagas Ocupadas"].sum()
            st.metric(label="📈 1. Total Geral de Vagas de Estágio Ocupadas no HCID", value=f"{total_geral_vagas} Vagas")
            st.markdown("---")
            
            fig2 = px.histogram(df_final_hcid, x="Setor", title="2. Total de Setores Disponibilizados para Campo de Estágio no HCID", color_discrete_sequence=[paleta_pasteis])
            st.plotly_chart(fig2, use_container_width=True)
            
            df_g3 = df_final_hcid.groupby("Setor", as_index=False)["Vagas Ocupadas"].sum()
            fig3 = px.bar(df_g3, x="Setor", y="Vagas Ocupadas", title="3. Setores Disponibilizados para a Realização de Estágio no HCID (Soma de Vagas)", color_discrete_sequence=[paleta_pasteis], text_auto=True)
            st.plotly_chart(fig3, use_container_width=True)
            
            fig4 = px.bar(df_final_hcid, x="Setor", y="Vagas Ocupadas", color="Categoria Profissional", title="4. Categorias Profissionais Contempladas no Estágio por Setor no HCID", barmode="group", color_discrete_sequence=paleta_pasteis)
            st.plotly_chart(fig4, use_container_width=True)
            
            fig5 = px.bar(df_g3, x="Vagas Ocupadas", y="Setor", orientation="h", title="5. Total de Vagas de Estágio Disponibilizadas por Setor no HCID", color_discrete_sequence=[paleta_pasteis], text_auto=True)
            st.plotly_chart(fig5, use_container_width=True)
            
            fig6 = px.pie(df_final_hcid, names="Turno", values="Vagas Ocupadas", title="6. Total de Vagas de Estágio do HCID por Turno", color_discrete_sequence=paleta_pasteis)
            st.plotly_chart(fig6, use_container_width=True)
            
            fig7 = px.bar(df_final_hcid, x="Dia da Semana", y="Vagas Ocupadas", color="Turno", title="7. Total de Estagiários por Turno, por Dia, no HCID", barmode="group", color_discrete_sequence=paleta_pasteis, text_auto=True)
            st.plotly_chart(fig7, use_container_width=True)
        else:
            st.warning("Nenhum dado numérico de vagas encontrado na escala da aba HCID.")

    # --- BLOCO VISUAL: ANEXO ---
    with aba_anexo:
        df_raw_anexo = pd.read_excel(uploaded_file, sheet_name="ANEXO", header=None)
        df_filtro_anexo = processar_escala_complexa(df_raw_anexo)
        
        if not df_filtro_anexo.empty:
            meses_anexo = [m for m in df_filtro_anexo["Mês"].unique() if m and m != "NAN"]
            mes_sel_anexo = st.multiselect("Selecione os Meses para Análise (Anexo):", meses_anexo, default=meses_anexo)
            df_final_anexo = df_filtro_anexo[df_filtro_anexo["Mês"].isin(mes_sel_anexo)]
            
            total_geral_anexo = df_final_anexo["Vagas Ocupadas"].sum()
            st.metric(label="📈 1. Total Geral de Vagas de Estágio Ocupadas no Anexo", value=f"{total_geral_anexo} Vagas")
            st.markdown("---")
            
            fig2_ax = px.histogram(df_final_anexo, x="Setor", title="2. Total de Setores Disponibilizados por Campo de Estágio no Anexo", color_discrete_sequence=[paleta_pasteis])
            st.plotly_chart(fig2_ax, use_container_width=True)
            
            df_g3_ax = df_final_anexo.groupby("Setor", as_index=False)["Vagas Ocupadas"].sum()
            fig3_ax = px.bar(df_g3_ax, x="Setor", y="Vagas Ocupadas", title="3. Setores Disponibilizados para a Realização de Estágio no Anexo", color_discrete_sequence=[paleta_pasteis], text_auto=True)
            st.plotly_chart(fig3_ax, use_container_width=True)
            
            fig4_ax = px.bar(df_final_anexo, x="Setor", y="Vagas Ocupadas", color="Categoria Profissional", title="4. Categorias Profissionais Contempladas no Estágio por Setor no Anexo", barmode="group", color_discrete_sequence=paleta_pasteis)
            st.plotly_chart(fig4_ax, use_container_width=True)
            
            fig5_ax = px.bar(df_g3_ax, x="Vagas Ocupadas", y="Setor", orientation="h", title="5. Total de Vagas de Estágio Disponibilizadas por Setor no Anexo", color_discrete_sequence=[paleta_pasteis], text_auto=True)
            st.plotly_chart(fig5_ax, use_container_width=True)
            
