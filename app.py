import streamlit as st
import pandas as pd
import altair as alt

#--- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Dashboard Executivo NAQH", page_icon="📊", layout="wide")

st.markdown("# 📊 Dashboard Executivo de Indicadores — NAQH")
st.markdown("Monitoramento técnico de conformidade regulatória e volumetria documental.")
st.markdown("---")

#--- 2. BARRA LATERAL (UPLOADS E FILTROS COMPACTOS) ---
st.sidebar.header("⚙️ Painel de Controle")

arquivo_excel = st.sidebar.file_uploader("📥 Carregar Planilha Excel (.xlsx):", type=["xlsx"])

if arquivo_excel:
    try:
        # 'header=2' ignora as linhas mescladas de título do topo da sua planilha
        df = pd.read_excel(arquivo_excel, engine="openpyxl", header=2)
        df.columns = df.columns.astype(str)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        if not df.empty:
            col_lista = list(df.columns)
            col_tipo = col_lista[0] if len(col_lista) > 0 else "Tipo"
            col_setor = col_lista[1] if len(col_lista) > 1 else "Setor"
            
            st.sidebar.markdown("---")
            st.sidebar.subheader("🔍 Filtros de Segmentação")
            
            # Filtros na lateral para controle global externo
            tipos_selecionados = st.sidebar.multiselect(f"Filtrar por {col_tipo}:", options=df[col_tipo].dropna().unique(), default=df[col_tipo].dropna().unique())
            setores_selecionados = st.sidebar.multiselect(f"Filtrar por {col_setor}:", options=df[col_setor].dropna().unique(), default=df[col_setor].dropna().unique())

            # Base de dados pré-filtrada pelo menu lateral
            df_base = df[(df[col_setor].isin(setores_selecionados)) & (df[col_tipo].isin(tipos_selecionados))]

            #--- SELETORES DE INTERATIVIDADE CLÁSSICOS (BLINDADOS CONTRA ERROS) ---
            selecao_clique_tipo = alt.selection_single(fields=[col_tipo], name="clique_tipo")
            selecao_clique_setor = alt.selection_single(fields=[col_setor], name="clique_setor")

            # Preparação dos dados agrupados para os gráficos
            dados_tipo = df_base[col_tipo].value_counts().reset_index()
            dados_tipo.columns = [col_tipo, 'Quantidade']

            dados_setor = df_base[col_setor].value_counts().reset_index()
            dados_setor.columns = [col_setor, 'Quantidade']

            #--- 3. MONTAGEM DOS DOIS GRÁFICOS INTERATIVOS LADO A LADO ---
            st.markdown("### 📊 Análise Gráfica de Demandas Interativa")
            st.markdown("💡 *Dica: Clique em qualquer barra de um dos gráficos abaixo para filtrar os números e a tabela automaticamente!*")
            
            col_g1, col_g2 = st.columns(2)

            with col_g1:
                grafico_tipo = alt.Chart(dados_tipo).mark_bar().encode(
                    x=alt.X(f'{col_tipo}:N', sort='-y', title=col_tipo),
                    y=alt.Y('Quantidade:Q', title='Quantidade'),
                    color=alt.condition(selecao_clique_tipo, alt.value("#1f77b4"), alt.value("#4A5260")),
                    opacity=alt.condition(selecao_clique_tipo, alt.value(1.0), alt.value(0.3)),
                    tooltip=[col_tipo, 'Quantidade']
                ).add_selection(selecao_clique_tipo).properties(height=350, title=f"Volumetria por {col_tipo}")
                
                evento_tipo = st.altair_chart(grafico_tipo, use_container_width=True, on_select="rerun")

            with col_g2:
                grafico_setor = alt.Chart(dados_setor).mark_bar().encode(
                    x=alt.X(f'{col_setor}:N', sort='-y', title=col_setor),
                    y=alt.Y('Quantidade:Q', title='Quantidade'),
                    color=alt.condition(selecao_clique_setor, alt.value("#2ca02c"), alt.value("#4A5260")),
                    opacity=alt.condition(selecao_clique_setor, alt.value(1.0), alt.value(0.3)),
                    tooltip=[col_setor, 'Quantidade']
                ).add_selection(selecao_clique_setor).properties(height=350, title=f"Demandas por {col_setor}")
                
                evento_setor = st.altair_chart(grafico_setor, use_container_width=True, on_select="rerun")

            #--- 4. APLICAÇÃO DINÂMICA DO CLIQUE NA BASE DE DATOS ---
            df_filtrado = df_base.copy()
            
            # Filtro inteligente pelo clique no gráfico de tipos
            if evento_tipo and "clique_tipo" in evento_tipo.get("selection", {}):
                valores_clicados_tipo = evento_tipo["selection"]["clique_tipo"]
                if valores_clicados_tipo:
                    df_filtrado = df_filtrado[df_filtrado[col_tipo].isin([v[col_tipo] for v in valores_clicados_tipo])]

            # Filtro inteligente pelo clique no gráfico de setores
            if evento_setor and "clique_setor" in evento_setor.get("selection", {}):
                valores_clicados_setor = evento_setor["selection"]["clique_setor"]
                if valores_clicados_setor:
                    df_filtrado = df_filtrado[df_filtrado[col_setor].isin([v[col_setor] for v in valores_clicados_setor])]

            st.markdown("---")
            
            #--- 5. CARTÕES DESIGN PREMIUM (KPIs RECALCULADOS PELO CLIQUE) ---
            kpi_col1, kpi_kpi2, kpi_col3 = st.columns(3)
            total_docs = len(df_filtrado)
            
            # CORREÇÃO DA VARIÁVEL: Pega estritamente a primeira ocorrência textual para evitar erro de DataFrame
            col_status_lista = [c for c in df.columns if "OK" in str(c).upper() or "STATUS" in str(c).upper() or "SITUA" in str(c).upper()]
            
            if col_status_lista:
                alvo_status = col_status_lista[0]
                aprovados = len(df_filtrado[df_filtrado[alvo_status].astype(str).str.upper().str.contains("APROVADO|OK|SIM|✔️", regex=True)])
            else:
                aprovados = total_docs
                
            taxa_aprovacao = int((aprovados / total_docs) * 100) if total_docs > 0 else 0

            with kpi_col1:
                st.markdown(f"<div style='background-color: #1E222D; padding: 20px; border-radius: 10px; border-left: 5px solid #1f77b4; text-align: center;'><span style='color: #8C949E; font-size: 14px; font-weight: bold; text-transform: uppercase;'>📄 Total Selecionado</span><h2 style='color: #FFFFFF; margin: 10px 0 0 0; font-size: 32px;'>{total_docs}</h2></div>", unsafe_allow_html=True)

            with kpi_kpi2:
                st.markdown(f"<div style='background-color: #1E222D; padding: 20px; border-radius: 10px; border-left: 5px solid #2ca02c; text-align: center;'><span style='color: #8C949E; font-size: 14px; font-weight: bold; text-transform: uppercase;'>✅ Concluídos / Aprovados</span><h2 style='color: #2ca02c; margin: 10px 0 0 0; font-size: 32px;'>{aprovados}</h2></div>", unsafe_allow_html=True)

            with kpi_col3:
                st.markdown(f"<div style='background-color: #1E222D; padding: 20px; border-radius: 10px; border-left: 5px solid #ff7f0e; text-align: center;'><span style='color: #8C949E; font-size: 14px; font-weight: bold; text-transform: uppercase;'>📈 Índice de Eficiência</span><h2 style='color: #ff7f0e; margin: 10px 0 0 0; font-size: 32px;'>{taxa_aprovacao}%</h2></div>", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            
            #--- 6. VISUALIZAÇÃO DA TABELA TÉCNICA DINÂMICA ---
            st.markdown("### 📋 Banco de Dados Estruturado (Filtro Ativo)")
            st.dataframe(df_filtrado, use_container_width=True)
            
    except Exception as e:
        st.error(f"⚠️ Falha técnica ao processar a interatividade do Excel: {e}")
else:
    st.info("💡 **Painel Pronto:** Carregue a planilha Excel no menu lateral para ativar os gráficos interativos por clique.")
