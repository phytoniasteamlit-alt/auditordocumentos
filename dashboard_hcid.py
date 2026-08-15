import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Configuração da página profissional (Layout Expandido)
st.set_page_config(
    page_title="Painel de Indicadores Norma Zero",
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
    st.markdown("### 🎨 Customização Visual")
    cor_tema = st.selectbox("Tema de Cores Geral:", ["Padrão (Azul)", "Vibrante", "Warm"])
    escala_cores = px.colors.sequential.Blues if cor_tema == "Padrão (Azul)" else (px.colors.sequential.Plasma if cor_tema == "Vibrante" else px.colors.sequential.Sunset)

# 4. Cabeçalho Principal (CORRIGIDO: Inserido o número 2 dentro de st.columns)
col_titulo, col_info = st.columns(2)
with col_titulo:
    st.markdown("# 📊 Painel de Indicadores")
    st.markdown("## Norma Zero")

with col_info:
    st.markdown(
        """
        <div style="text-align: right; padding-top: 15px; border-left: 2px solid #333; padding-left: 20px;">
            <span style="font-size: 20px; font-weight: bold;">🏥 Hospital da Cidade</span><br>
            <span style="font-size: 15px; color: #aaaaaa;">👩‍💼 Coord: Verônica Azevedo</span>
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
    linha_turno = df_raw
    
    dados_estruturados = []
    
    for idx_linha in range(7, len(df_raw)):
        linha_atual = df_raw.iloc[idx_linha]
        
        setor_macro = str(linha_atual.iloc[0]).strip()
        sub_setor = str(linha_atual.iloc[1]).strip()
        categoria = str(linha_atual.iloc[2]).strip()
        
        if sub_setor and sub_setor != "None" and sub_setor != setor_macro:
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
                    
                mes = str(linha_mes.iloc[4, idx_col]).strip()
                dia = str(linha_dia.iloc[5, idx_col]).strip()
                turno = str(linha_turno.iloc[6, idx_col]).strip()
                
                if "manh" in turno.lower(): turno = "Manhã"
                elif "tard" in turno.lower(): turno = "Tarde"
                elif "noit" in turno.lower(): turno = "Noite"
                
                dados_estruturados.append({
                    "Setor": setor_final,
                    "Categoria Profissional": categoria,
                    "Mês": mes if mes != "None" else "Geral",
                    "Dia da Semana": dia if dia != "None" else "Não Informado",
                    "Turno": turno if turno != "None" else "Integral",
                    "Vagas Ocupadas": qtd_vagas
                })
                
    return pd.DataFrame(dados_estruturados)

# 5. Renderização e Execução do Upload
if uploaded_file is not None:
    try:
        aba_hcid, aba_anexo = st.tabs(["🏢 UNIDADE HCID", "📑 UNIDADE ANEXO"])
        
        # --- BLOCO VISUAL: HCID ---
        with aba_hcid:
            df_raw_hcid = pd.read_excel(uploaded_file, sheet_name="HCID", header=None)
            df_filtro_hcid = processar_escala_complexa(df_raw_hcid)
            
            if not df_filtro_hcid.empty:
                meses_hcid = df_filtro_hcid["Mês"].unique()
                mes_sel_hcid = st.multiselect("Selecione os Meses para Análise (HCID):", meses_hcid, default=meses_hcid)
                df_final_hcid = df_filtro_hcid[df_filtro_hcid["Mês"].isin(mes_sel_hcid)]
                
                st.plotly_chart(px.bar(df_final_hcid, y="Vagas Ocupadas", title="1. Total de Vagas de Estágio no HCID", color_discrete_sequence=escala_cores), use_container_width=True)
                st.plotly_chart(px.histogram(df_final_hcid, x="Setor", title="2. Total de Setores Disponibilizados para Campo de Estágio no HCID", color_discrete_sequence=escala_cores), use_container_width=True)
                st.plotly_chart(px.bar(df_final_hcid, x="Setor", y="Vagas Ocupadas", title="3. Setores Disponibilizados para a Realização de Estágio no HCID"), use_container_width=True)
                st.plotly_chart(px.bar(df_final_hcid, x="Setor", y="Vagas Ocupadas", color="Categoria Profissional", title="4. Categorias Profissionais Contempladas no Estágio por Setor no HCID", barmode="group"), use_container_width=True)
                st.plotly_chart(px.bar(df_final_hcid, x="Vagas Ocupadas", y="Setor", orientation="h", title="5. Total de Vagas de Estágio Disponibilizadas por Setor no HCID", color_discrete_sequence=escala_cores), use_container_width=True)
                st.plotly_chart(px.pie(df_final_hcid, names="Turno", values="Vagas Ocupadas", title="6. Total de Vagas de Estágio do HCID por Turno"), use_container_width=True)
                st.plotly_chart(px.bar(df_final_hcid, x="Dia da Semana", y="Vagas Ocupadas", color="Turno", title="7. Total de Estagiários por Turno, por Dia, no HCID", barmode="group"), use_container_width=True)
            else:
                st.warning("Nenhum dado numérico de vagas encontrado na escala da aba HCID.")

        # --- BLOCO VISUAL: ANEXO ---
        with aba_anexo:
            try:
                df_raw_anexo = pd.read_excel(uploaded_file, sheet_name="ANEXO", header=None)
                df_filtro_anexo = processar_escala_complexa(df_raw_anexo)
                
                if not df_filtro_anexo.empty:
                    meses_anexo = df_filtro_anexo["Mês"].unique()
                    mes_sel_anexo = st.multiselect("Selecione os Meses para Análise (Anexo):", meses_anexo, default=meses_anexo)
                    df_final_anexo = df_filtro_anexo[df_filtro_anexo["Mês"].isin(mes_sel_anexo)]
                    
                    st.plotly_chart(px.bar(df_final_anexo, y="Vagas Ocupadas", title="1. Total de Vagas de Estágio no Anexo", color_discrete_sequence=escala_cores), use_container_width=True)
                    st.plotly_chart(px.histogram(df_final_anexo, x="Setor", title="2. Total de Setores Disponibilizados por Campo de Estágio no Anexo", color_discrete_sequence=escala_cores), use_container_width=True)
                    st.plotly_chart(px.bar(df_final_anexo, x="Setor", y="Vagas Ocupadas", title="3. Setores Disponibilizados para a Realização de Estágio no Anexo"), use_container_width=True)
                    st.plotly_chart(px.bar(df_final_anexo, x="Setor", y="Vagas Ocupadas", color="Categoria Profissional", title="4. Categorias Profissionais Contempladas no Estágio por Setor no Anexo", barmode="group"), use_container_width=True)
                    st.plotly_chart(px.bar(df_final_anexo, x="Vagas Ocupadas", y="Setor", orientation="h", title="5. Total de Vagas de Estágio Disponibilizadas por Setor no Anexo", color_discrete_sequence=escala_cores), use_container_width=True)
                    st.plotly_chart(px.pie(df_final_anexo, names="Turno", values="Vagas Ocupadas", title="6. Total de Vagas de Estágio no Anexo por Turno"), use_container_width=True)
                    st.plotly_chart(px.bar(df_final_anexo, x="Dia da Semana", y="Vagas Ocupadas", color="Turno", title="7. Total de Estagiários por Turno e por Dia no Anexo", barmode="group"), use_container_width=True)
                else:
                    st.info("ℹ️ Os gráficos do Anexo serão ativados automaticamente assim que houver registros de vagas maiores que zero na planilha.")
            except Exception:
                st.error("Aba 'ANEXO' não encontrada ou formato incompatível.")

    except Exception as e:
        st.error(f"⚠️ Erro crítico no processamento da matriz: {e}")
