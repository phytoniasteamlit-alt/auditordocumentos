import streamlit as st
import pandas as pd
import altair as alt
import re

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Dashboard Executivo NAQH", page_icon="📊", layout="wide")
st.markdown("# Dashboard Executivo de Indicadores — NAQH")
st.markdown("Monitoramento técnico de conformidade regulatória e volumetria documental.")
st.markdown("---")

#--- FUNÇÃO MODULAR AUXILIAR PARA CRIAR OS GRÁFICOS ---
def plotar_grafico_naqh(dados, coluna_nome, tipo_visual, seletor_clique, cor_barra, titulo_grafico):
    if tipo_visual == "Barra":
        base_chart = alt.Chart(dados).encode(
            x=alt.X(f"{coluna_nome}:N", sort='-y', title=coluna_nome, axis=alt.Axis(labelAngle=0)), # Texto reto na horizontal
            y=alt.Y("Quantidade:Q", title="Quantidade"),
            tooltip=[coluna_nome, 'Quantidade']
        )
        
        bars = base_chart.mark_bar().encode(
            color=alt.condition(seletor_clique, alt.value(cor_barra), alt.value("#4A5260")),
            opacity=alt.condition(seletor_clique, alt.value(1.0), alt.value(0.3))
        ).add_selection(seletor_clique)
        
        text = base_chart.mark_text(
            align='center',
            baseline='bottom',
            dy=-5,
            color='white',
            fontWeight='bold'
        ).encode(text='Quantidade:Q')
        
        chart = (bars + text).properties(height=280, title=titulo_grafico)
        
    else:
        base_chart = alt.Chart(dados).encode(
            theta=alt.Theta("Quantidade:Q", title="Quantidade"),
            color=alt.Color(f"{coluna_nome}:N", title=coluna_nome),
            tooltip=[coluna_nome, 'Quantidade']
        )
        
        arcs = base_chart.mark_arc().encode(
            opacity=alt.condition(seletor_clique, alt.value(1.0), alt.value(0.3))
        ).add_selection(seletor_clique)
        
        text = base_chart.mark_text(radiusOffset=15, color='white', fontWeight='bold').encode(text='Quantidade:Q')
        
        chart = (arcs + text).properties(height=280, title=titulo_grafico)
        
    return chart

#--- 2. BARRA LATERAL (UPLOADS E FILTROS COMPACTOS) ---
st.sidebar.header("⚙️ Painel de Controle")
arquivo_excel = st.sidebar.file_uploader(" Carregar Planilha Excel (.xlsx):", type=["xlsx"])

if arquivo_excel:
    # 'header=2' pula as duas primeiras linhas de título mesclado do NAQH
    df = pd.read_excel(arquivo_excel, engine="openpyxl", header=2)
    df.columns = df.columns.astype(str).str.strip()
    
    if not df.empty:
        # --- BLINDAGEM COMPLETA POR POSIÇÃO FÍSICA (ÍNDICES) ---
        col_tipo = df.columns[0]   # Coluna A (SIGLA DO DOCUMENTO)
        col_setor = df.columns[1]  # Coluna B (SETOR)
        col_responsavel = df.columns[4] if len(df.columns) > 4 else df.columns[4]  # Coluna E (RESPONSÁVEL)
        
        # Posições exatas mapeadas a partir dos prints reais da sua tabela
        col_status = df.columns[16] if len(df.columns) > 16 else df.columns[16]    # Coluna Q (STATUS / OK)
        col_situacao_nome = df.columns[17] if len(df.columns) > 17 else df.columns[17] # Coluna R (SITUAÇÃO)
        
        col_t_v1 = df.columns[19] if len(df.columns) > 19 else df.columns[19]       # Coluna T (I.A.V.1º)
        col_t_v2 = df.columns[20] if len(df.columns) > 20 else df.columns[20]       # Coluna U (I.A.V.2º)
        col_t_total = df.columns[21] if len(df.columns) > 21 else df.columns[21]    # Coluna V (I.A.A.A)

        # --- PROCESSAMENTO MATEMÁTICO DOS TEMPOS DE ANÁLISE ---
        def extrair_dias_puros(valor):
            if pd.isna(valor) or "#" in str(valor):
                return 0.0
            numeros = re.findall(r'\d+', str(valor))
            if numeros:
                return float(numeros[0]) # Pega estritamente apenas o primeiro número (dias corridos)
            return 0.0

        df['T_V1'] = df[col_t_v1].apply(extrair_dias_puros)
        df['T_V2'] = df[col_t_v2].apply(extrair_dias_puros)
        df['T_TOTAL'] = df[col_t_total].apply(extrair_dias_puros)

        # --- SEÇÃO DE FILTROS GLOBAIS NA BARRA LATERAL ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔍 Filtros Globais")
        
        opcoes_tipo = df[col_tipo].dropna().unique().tolist()
        tipos_selecionados = st.sidebar.multiselect(f"Filtrar por Tipo:", options=opcoes_tipo, default=opcoes_tipo)
        
        opcoes_setor = df[col_setor].dropna().unique().tolist()
        setores_selecionados = st.sidebar.multiselect(f"Filtrar por Setor:", options=opcoes_setor, default=opcoes_setor)
        
        # Filtro dinâmico para Adicionar / Remover funcionários responsáveis pela barra lateral
        opcoes_resp = df[col_responsavel].dropna().unique().tolist()
        responsaveis_selecionados = st.sidebar.multiselect("Acrescentar / Retirar Responsáveis:", options=opcoes_resp, default=opcoes_resp)
        
        # Aplicação combinada dos filtros estruturais
        df_base = df[(df[col_setor].isin(setores_selecionados)) & (df[col_tipo].isin(tipos_selecionados))]
        if responsaveis_selecionados:
            df_base = df_base[df_base[col_responsavel].isin(responsaveis_selecionados)]
        
        #--- SELETORES DE INTERATIVIDADE DO ALTAIR ---
        selecao_clique_tipo = alt.selection_single(fields=[col_tipo], name="clique_tipo")
        selecao_clique_setor = alt.selection_single(fields=[col_setor], name="clique_setor")
        selecao_clique_status = alt.selection_single(fields=[col_status], name="clique_status")
        selecao_clique_resp = alt.selection_single(fields=[col_responsavel], name="clique_resp")

        # Agrupamento estruturado para plotagem dos 6 gráficos requisitados
        dados_tipo = df_base[col_tipo].value_counts().reset_index()
        dados_tipo.columns = [col_tipo, 'Quantidade']
        
        dados_setor = df_base[col_setor].value_counts().reset_index()
        dados_setor.columns = [col_setor, 'Quantidade']
        
        dados_status = df_base[col_status].value_counts().reset_index()
        dados_status.columns = [col_status, 'Quantidade']
        
        dados_situacao = df_base[df_base[col_situacao_nome].astype(str).str.contains("#") == False][col_situacao_nome].value_counts().reset_index()
        dados_situacao.columns = [col_situacao_nome, 'Quantidade']
        
        dados_resp = df_base[col_responsavel].value_counts().reset_index()
        dados_resp.columns = [col_responsavel, 'Quantidade']

        #--- 3. MONTAGEM DO LAYOUT DE GRÁFICOS INTERATIVOS LADO A LADO ---
        st.markdown("### 📊 Painel de Análise Gráfica Cruzada")
        st.markdown("*Dica: Use as caixas de seleção acima de cada gráfico para alternar entre Barra ou Pizza!*")
        
        # Bloco 1: POP e Setores
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            tipo_visual_g1 = st.selectbox("Formato (Tipo de Documento):", ["Barra", "Pizza"], key="g1_visual")
            chart_tipo = plotar_grafico_naqh(dados_tipo, col_tipo, tipo_visual_g1, selecao_clique_tipo, "#1f77b4", "Volumetria por Sigla do Documento")
            evento_tipo = st.altair_chart(chart_tipo, use_container_width=True, on_select="rerun")
            
        with col_g2:
            tipo_visual_g2 = st.selectbox("Formato (Setor Requisitante):", ["Barra", "Pizza"], key="g2_visual")
            chart_setor = plotar_grafico_naqh(dados_setor, col_setor, tipo_visual_g2, selecao_clique_setor, "#2ca02c", "Demandas por Setor")
            evento_setor = st.altair_chart(chart_setor, use_container_width=True, on_select="rerun")

        # Bloco 2: Status Executivo e Situação de Prazos (Mapeamento Cirúrgico)
        col_g3, col_g4 = st.columns(2)
        with col_g3:
            tipo_visual_g3 = st.selectbox("Formato (Status Executivo):", ["Barra", "Pizza"], key="g3_visual")
            chart_status = plotar_grafico_naqh(dados_status, col_status, tipo_visual_g3, selecao_clique_status, "#9467bd", "Status Executivo (Aprovados / Cancelados / Em Verificação)")
            evento_status = st.altair_chart(chart_status, use_container_width=True, on_select="rerun")

        with col_g4:
            tipo_visual_g4 = st.selectbox("Formato (Situação de Prazo):", ["Barra", "Pizza"], key="g4_visual")
            if tipo_visual_g4 == "Barra":
                base_situacao = alt.Chart(dados_situacao).encode(
                    x=alt.X(f'{col_situacao_nome}:N', sort='-y', title="Situação", axis=alt.Axis(labelAngle=0)),
                    y=alt.Y('Quantidade:Q', title="Quantidade"),
                    tooltip=[col_situacao_nome, 'Quantidade']
                )
                chart_situacao = base_situacao.mark_bar().encode(color=alt.value("#e377c2")) + base_situacao.mark_text(align='center', baseline='bottom', dy=-5, color='white', fontWeight='bold').encode(text='Quantidade:Q')
                chart_situacao = chart_situacao.properties(height=280, title="Prazos (Válido / Vencido / Prestes a Vencer)")
            else:
                chart_situacao = alt.Chart(dados_situacao).mark_arc().encode(theta=alt.Theta('Quantidade:Q', title="Quantidade"), color=alt.Color(f'{col_situacao_nome}:N', title="Situação"), tooltip=[col_situacao_nome, 'Quantidade']).properties(height=280, title="Prazos (Válido / Vencido / Prestes a Vencer)")
            st.altair_chart(chart_situacao, use_container_width=True)

        # Bloco 3: Responsáveis (Lançamentos) e Gráfico Isolado de Médias das Colunas T, U e V
        col_g5, col_g6 = st.columns(2)
        with col_g5:
            tipo_visual_g5 = st.selectbox("Formato (Funcionário Responsável):", ["Barra", "Pizza"], key="g5_visual")
            chart_resp = plotar_grafico_naqh(dados_resp, col_responsavel, tipo_visual_g5, selecao_clique_resp, "#ff7f0e", "Documentos Lançados por Responsável")
            evento_resp = st.altair_chart(chart_resp, use_container_width=True, on_select="rerun")

        with col_g6:
            st.selectbox("Formato (Tempos Médios):", ["Barra"], disabled=True, key="g6_visual")
