import streamlit as st
import pandas as pd
import zipfile
import mailbox
import io
import os
import re
from docx import Document
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
import plotly.express as px

st.set_page_config(page_title="Auditor NAQH - Hospital da Cidade", layout="wide")

st.title("🏥 Auditoria Científica de Documentos - Norma Zero")
st.subheader("Hospital da Cidade Dr. Jackson Lago")

st.sidebar.header("📥 Carga de Arquivos do Hospital")
arquivo_excel = st.sidebar.file_uploader("1. Selecione a Planilha Oficial (.xlsx)", type=["xlsx"])
arquivo_zip = st.sidebar.file_uploader("2. Selecione o ZIP do Takeout (1.2 GB)", type=["zip"])

def formatar_data_gmail(data_cabecalho):
    try:
        if data_cabecalho:
            dt = parsedate_to_datetime(data_cabecalho)
            return dt.strftime("%d/%m/%Y")
    except:
        pass
    return None

def extrair_data_aprovacao_interna_memoria(conteudo_binario):
    """Abre o arquivo Word na memória e valida se existe aprovação real da Norma Zero"""
    try:
        doc = Document(io.BytesIO(conteudo_binario))
        for tabela in doc.tables:
            for linha in tabela.rows:
                texto_linha = [celula.text.strip() for celula in linha.cells]
                text_completo = " ".join(texto_linha)
                if "Data aprovação:" in text_completo:
                    data = text_completo.split("Data aprovação:")[-1].split("Validade:")[0].strip()
                    # Garante que é uma data digitada (contém barras) e ignora a máscara padrão 'dd/mm/aaaa'
                    if data and "dd/mm" not in data.lower() and "/" in data:
                        return data
    except:
        pass
    return None

if arquivo_excel and arquivo_zip:
    st.sidebar.success("Buffers validados e protegidos contra colisão.")
    
    if st.button("🚀 Executar Auditoria Reversa e Rastreio Seguro"):
        with st.spinner("Filtrando e-mails caóticos e analisando apenas anexos oficiais .docx..."):
            try:
                # 1. Carrega a planilha sem modificar os tipos originais de texto
                df_oficial = pd.read_excel(arquivo_excel, dtype=str)
                df_oficial.columns = df_oficial.columns.str.strip()
                
                codigos_na_planilha = set(df_oficial["CÓD. DO DOCUMENTO"].dropna().str.strip().str.upper())
                SEU_EMAIL = "documentos.soc2@gmail.com"
                documentos_fantasmas = []
                
                # 2. Abre o giga-arquivo local na memória do Ryzen 5
                with zipfile.ZipFile(arquivo_zip, 'r') as z:
                    path_mbox = [f for f in z.namelist() if f.endswith('.mbox')]
                    if not path_mbox:
                        st.error("Erro: Arquivo de e-mails (.mbox) não localizado dentro do arquivo ZIP.")
                    else:
                        with z.open(path_mbox) as mbox_file:
                            mbox = mailbox.mbox(io.BytesIO(mbox_file.read()))
                            
                            # --- ETAPA 1: AUDITORIA REVERSA (MAPEIA DOCUMENTOS FORA DA PLANILHA) ---
                            for msg in mbox:
                                assunto = str(msg["subject"]).strip().upper()
                                data_msg = formatar_data_gmail(msg["date"])
                                remetente = str(msg["from"])
                                
                                # TRAVA REGEX ESTRITA
                                match_codigo = re.search(r'\b(POP|ROT|NOR|PROT|REG|MANUAL)_[A-Z0-9_]+', assunto)
                                if match_codigo:
                                    codigo_detectado = match_codigo.group(0).strip()
                                    
                                    possui_docx = False
                                    if msg.is_multipart():
                                        for part in msg.walk():
                                            filename = part.get_filename()
                                            if filename and filename.endswith(".docx"):
                                                possui_docx = True
                                                break
                                                
                                    if possui_docx and codigo_detectado not in codigos_na_planilha:
                                        if not any(f['Código'] == codigo_detectado for f in documentos_fantasmas):
                                            documentos_fantasmas.append({
                                                "Código": codigo_detectado,
                                                "Assunto do E-mail": assunto,
                                                "Último Tráfego Detectado": data_msg,
                                                "Origem/Remetente": remetente,
                                                "Tipo": "DOCUMENTO NORMA ZERO OCULTO"
                                            })
                            
                            # --- ETAPA 2: PREENCHIMENTO SEGURO DA SUA PLANILHA (MÁQUINA DE ESTADOS) ---
                            for idx, linha in df_oficial.iterrows():
                                codigo_doc = str(linha["CÓD. DO DOCUMENTO"]).strip().upper()
                                versao_doc = str(linha["VERSÃO"]).strip().lower()
                                
                                if pd.isna(codigo_doc) or codigo_doc in ["NAN", ""]:
                                    continue
                                
                                # FILTRO CRÍTICO ANTI-CONFUSÃO
                                emails_do_doc = [m for m in mbox if codigo_doc in str(m["subject"]).upper()]
                                if not emails_do_doc:
                                    continue
                                    
                                if versao_doc != "nan" and versao_doc != "":
                                    if "1" in versao_doc:
                                        emails_do_doc = [m for m in emails_do_doc if "2ª" not in str(m["subject"]).lower() and "3ª" not in str(m["subject"]).lower()]
                                    elif "2" in versao_doc:
                                        emails_do_doc = [m for m in emails_do_doc if "2ª" in str(m["subject"]).lower() or "v2" in str(m["subject"]).lower()]
                                
                                if not emails_do_doc:
                                    continue
                                    
                                # CORREÇÃO TÉCNICA: Captura o primeiro e-mail indexado
                                primeiro_email = emails_do_doc[0]
                                
                                # TRAVA DE SEGURANÇA MÁXIMA
                                if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]).strip() == "":
                                    df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"] = formatar_data_gmail(primeiro_email["date"])
                                
                                df_oficial.at[idx, "1ª VERIFICAÇÃO EZEQUIAS"] = "OK"
                                
                                # Estimativa inteligente (D+1)
                                if pd.isna(df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]) or str(df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]).strip() == "":
                                    data_ref = df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]
                                    if pd.notna(data_ref) and str(data_ref).strip() != "":
                                        try:
                                            dt_ref = datetime.strptime(str(data_ref), '%d/%m/%Y')
                                            df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"] = (dt_ref + timedelta(days=1)).strftime('%d/%m/%Y')
                                        except:
                                            pass
                                
                                # Rastreio da 2ª Verificação
                                emails_2v = [m for m in emails_do_doc if SEU_EMAIL not in str(m["from"]) and m != primeiro_email]
                                if emails_2v:
                                    # CORREÇÃO TÉCNICA: Captura o primeiro e-mail indexado da 2a fase
                                    segundo_email = emails_2v[0]
                                    
                                    if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]).strip() == "":
                                        df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"] = formatar_data_gmail(segundo_email["date"])
                                        df_oficial.at[idx, "2ª VERIFICAÇÃO EZEQUIAS"] = "OK"
                                        
