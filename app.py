st.markdown("---")

# Gráfico 5: DOCUMENTOS APROVADOS POR TIPO
st.subheader("5 Documentos Aprovados por Tipo")
tipos_disponiveis = df["SIGLA DO DOCUMENTO"].dropna().unique().tolist()
tipos_selecionados = st.multiselect("Filtrar por Tipo de Documento (Sigla):", options=tipos_disponiveis, default=tipos_disponiveis)

df_g5 = df[(df["STATUS DO DOCUMENTO NORMATIVO"] == "APROVADO") & (df["SIGLA DO DOCUMENTO"].isin(tipos_selecionados))]
df_g5_counts = df_g5["SIGLA DO DOCUMENTO"].value_counts().reset_index()
df_g5_counts.columns = ["Tipo de Documento", "Quantidade Aprovada"]

g5_horizontal = (tipo_grafico_5 == "Barras Horizontais")
eixo_x_5 = "Quantidade Aprovada" if g5_horizontal else "Tipo de Documento"
eixo_y_5 = "Tipo de Documento" if g5_horizontal else "Quantidade Aprovada"
ori_5 = "h" if g5_horizontal else "v"

fig5 = px.bar(df_g5_counts, x=eixo_x_5, y=eixo_y_5, text="Quantidade Aprovada", color="Tipo de Documento", orientation=ori_5, color_discrete_sequence=cor_sequencia)
fig5.update_traces(textposition="outside")
st.plotly_chart(fig5, use_container_width=True)

st.markdown("---")

# Gráfico 6: DETALHAMENTO DE TIPOS DE DOCUMENTO POR PROFISSIONAL
st.subheader("6 Tipos de Documento por Profissional")
prof_selecionado_g6 = st.selectbox("Selecione o Profissional para detalhar os tipos de documentos:", options=profissionais_ativos, key="filtro_prof_g6")

df_g6_filtrado = df[df["RESPONSÁVEL"] == prof_selecionado_g6]
df_g6_counts = df_g6_filtrado.groupby(["SIGLA DO DOCUMENTO"]).size().reset_index(name="Quantidade")

g6_horizontal = (tipo_grafico_6 == "Barras Horizontais")
eixo_x_6 = "Quantidade" if g6_horizontal else "SIGLA DO DOCUMENTO"
eixo_y_6 = "SIGLA DO DOCUMENTO" if g6_horizontal else "Quantidade"
ori_6 = "h" if g6_horizontal else "v"

fig6 = px.bar(df_g6_counts, x=eixo_x_6, y=eixo_y_6, text="Quantidade", color="SIGLA DO DOCUMENTO", orientation=ori_6, color_discrete_sequence=cor_sequencia, labels={"SIGLA DO DOCUMENTO": "Tipo de Documento", "Quantidade": "Total Lançado"})
fig6.update_traces(textposition="outside")
fig6.update_layout(showlegend=False)
st.plotly_chart(fig6, use_container_width=True)
