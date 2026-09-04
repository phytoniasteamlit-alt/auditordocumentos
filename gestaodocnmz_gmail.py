import streamlit as st
import pandas as pd
import zipfile
import mailbox
import io
import re
from docx import Document
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
import plotly.express as px

st.set_page_config(page_title="Auditor NAQH - Hospital da Cidade", layout="wide")

st.title("🏥 Auditoria Científica de Documentos - Norma Zero")
st.subheader("Hospital da Cidade Dr. Jackson Lago")

st.sidebar.header("📥 Upload dos Arquivos Globais")
arquivo_excel = st.sidebar.file_uploader("1. Selecione a Planilha Oficial", type=["xlsx"])
arquivo_zip = st.sidebar.file_uploader("2. Selecione o ZIP do Takeout (1.2 GB)", type=["zip"])

def formatar_data_gmail(data_cabecalho):
    try:
        if data_cabecalho:
            dt = parsedate_to_datetime(data_cabecalho)
            return dt.strftime("%d/%m/%Y")
    except:
        pass
    return None

if arquivo_excel and arquivo_zip:
    st.sidebar.success("Buffers validados. Pronto para processamento local.")
    
    if st.button("🚀 Iniciar Auditoria Reversa e Cruzamento"):
        with st.spinner("Varrendo 1.2 GB de dados e mapeando documentos fantasmas..."):
            try:
                # 1. Carrega dados da planilha
                df_oficial = pd.read_excel(arquivo_excel, dtype=str)
                df_oficial.columns = df_oficial.columns.str.strip()
                
                # Criamos um set com todos os códigos que já existem na planilha para comparação rápida
                codigos_na_planilha = set(df_oficial["CÓD. DO DOCUMENTO"].dropna().str.strip().str.upper())
                
                SEU_EMAIL = "documentos.soc2@gmail.com"
                documentos_fantasmas = [] # Lista para guardar o que não estava no Excel
                
                # 2. Abre o ZIP e processa o MBOX na memória
                with zipfile.ZipFile(arquivo_zip, 'r') as z:
                    path_mbox = [f for f in z.namelist() if f.endswith('.mbox')]
                    with z.open(path_mbox[0]) as mbox_file:
                        mbox = mailbox.mbox(io.BytesIO(mbox_file.read()))
                        
                        # --- ETAPA 1: MAPEAMENTO DE DOCUMENTOS FORA DA PLANILHA (REVERSA) ---
                        for msg in mbox:
                            assunto = str(msg["subject"]).strip()
                            data_msg = formatar_data_gmail(msg["date"])
                            remetente = str(msg["from"])
                            
                            # Captura códigos usando Regex com base no padrão da Norma Zero (POP, ROT, NOR, PROT)
                            encontrados = re.findall(r'\b(POP|ROT|NOR|PROT|REG|MANUAL)_[A-Z0-String0-9_]+', assunto, re.IGNORECASE)
                            
                            if encontrados:
                                # Reconstrói o código achado no assunto
                                match_completo = re.search(r'\b(POP|ROT|NOR|PROT|REG|MANUAL)_[A-Z0-9_]+', assunto, re.IGNORECASE)
                                if match_completo:
                                    codigo_detectado = match_completo.group(0).upper().strip()
                                    
                                    # 🔥 SE O CÓDIGO NÃO ESTIVER NO NOSSO SET DA PLANILHA: Descobrimos um fantasma!
                                    if codigo_detectado not in codigos_na_planilha:
                                        # Evita duplicar o mesmo fantasma na listagem visual
                                        if not any(f['Código'] == codigo_detectado for f in documentos_fantasmas):
                                            documentos_fantasmas.append({
                                                "Código": codigo_detectado,
                                                "Assunto do E-mail": assunto,
                                                "Último Tráfego Detectado": data_msg,
                                                "Origem/Remetente": remetente,
                                                "Status no Sistema": "NÃO CATALOGADO NA PLANILHA"
                                            })
                        
                        # --- ETAPA 2: PREENCHIMENTO SEGURO DA PLANILHA EXISTENTE (Sua Lógica Atual) ---
                        for idx, linha in df_oficial.iterrows():
                            codigo_doc = str(linha["CÓD. DO DOCUMENTO"]).strip().upper()
                            if pd.isna(codigo_doc) or codigo_doc in ["NAN", ""]:
                                continue
                            
                            # Filtro estrito de e-mails para evitar colisões de versão
                            emails_do_doc = [m for m in mbox if codigo_doc in str(m["subject"]).upper()]
                            if not emails_do_doc:
                                continue
                                
                            primeiro_email = emails_do_doc[0]
                            if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]).strip() == "":
                                df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"] = formatar_data_gmail(primeiro_email["date"])
                            
                            df_oficial.at[idx, "1ª VERIFICAÇÃO EZEQUIAS"] = "OK"
                            
                            # Estimativa D+1 se esquecido pelas meninas
                            if pd.isna(df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]) or str(df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]).strip() == "":
                                data_ref = df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]
                                if pd.notna(data_ref):
                                    dt_ref = datetime.strptime(str(data_ref), '%d/%m/%Y')
                                    df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"] = (dt_ref + timedelta(days=1)).strftime('%d/%m/%Y')
                
                st.balloons()
                
                # --- EXIBIÇÃO DO DASHBOARD EM ABAS ---
                aba_principal, aba_fantasmas = st.tabs(["📊 Planilha Atualizada & Gráficos", "🚨 Documentos Omitidos (Não Catalogados)"])
                
                with aba_principal:
                    st.success("Planilha Oficial Sincronizada com Sucesso!")
                    st.dataframe(df_oficial)
                    
                    # Gráfico Dinâmico para a Coordenadora
                    if "STATUS" in df_oficial.columns:
                        df_oficial["STATUS"] = df_oficial["STATUS"].fillna("EM VERIFICAÇÃO")
                        fig = px.pie(df_oficial, names="STATUS", title="Visão Geral do Status dos Documentos")
                        st.plotly_chart(fig, use_container_width=True)
                        
                    # Buffer de download da planilha principal
                    out_oficial = io.BytesIO()
                    with pd.ExcelWriter(out_oficial, engine='openpyxl') as writer:
                        df_oficial.to_excel(writer, index=False)
                    st.download_button("📥 Baixar Planilha Oficial Atualizada", data=out_oficial.getvalue(), file_name="Planilha_Oficial_NAQH.xlsx")
                
                with aba_fantasmas:
                    if documentos_fantasmas:
                        st.error(f"Atenção: O robô detectou {len(documentos_fantasmas)} documentos trafegando no Gmail que NÃO existem na sua planilha!")
                        df_fantasmas = pd.DataFrame(documentos_fantasmas)
                        st.dataframe(df_fantasmas)
                        
                        # Buffer de download para a lista de omitidos
                        out_fantasmas = io.BytesIO()
                        with pd.ExcelWriter(out_fantasmas, engine='openpyxl') as writer:
                            df_fantasmas.to_excel(writer, index=False)
                        st.download_button("📥 Baixar Lista de Documentos Não Catalogados (.xlsx)", data=out_fantasmas.getvalue(), file_name="documentos_fantasmas_encontrados.xlsx")
                    else:
                        st.success("Excelente! Todos os documentos detectados no e-mail já constam na sua planilha oficial.")
                        
            except Exception as e:
                st.error(f"Erro crítico no processamento local: {e}")
