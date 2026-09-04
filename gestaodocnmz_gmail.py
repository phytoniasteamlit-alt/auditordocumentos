import os
import pandas as pd
from docx import Document
from datetime import datetime, timedelta

PASTA_TRABALHO = r"C:\Caminho\Para\Sua\Pasta\Do\Feriado"
PLANILHA_OFICIAL = "Sua_Planilha_Oficial.xlsx"
PLANILHA_EMAILS = "historico_emails.xlsx" 

caminho_excel = os.path.join(PASTA_TRABALHO, PLANILHA_OFICIAL)
caminho_emails = os.path.join(PASTA_TRABALHO, PLANILHA_EMAILS)

df_oficial = pd.read_excel(caminho_excel, dtype=str)
df_emails = pd.read_excel(caminho_emails)

df_emails['Data de Recebimento'] = pd.to_datetime(df_emails['Data de Recebimento'], errors='coerce')
df_emails = df_emails.sort_values(by='Data de Recebimento').dropna(subset=['Data de Recebimento'])

df_oficial.columns = df_oficial.columns.str.strip()

def extrair_data_aprovacao_interna(caminho_doc):
    try:
        doc = Document(caminho_doc)
        for tabela in doc.tables:
            for linha in tabela.rows:
                texto_linha = [celula.text.strip() for celula in linha.cells]
                text_completo = " ".join(texto_linha)
                if "Data aprovação:" in text_completo:
                    data = text_completo.split("Data aprovação:")[-1].split("Validade:").strip()
                    if data and "dd/mm" not in data.lower() and "/" in data:
                        return data
    except:
        return None
    return None

for arquivo_word in os.listdir(PASTA_TRABALHO):
    if arquivo_word.endswith(".docx") and not arquivo_word.startswith("~$"):
        caminho_doc = os.path.join(PASTA_TRABALHO, arquivo_word)
        
        data_aprovacao_word = extrair_data_aprovacao_interna(caminho_doc)
        cod_documento = arquivo_word.replace(".docx", "") 
        
        historico_do_doc = df_emails[df_emails['Nome do Anexo'].str.contains(arquivo_word, na=False, case=False)]
        
        if historico_do_doc.empty:
            continue
            
        filtro = df_oficial["CÓD. DO DOCUMENTO"].str.strip() == cod_documento if "CÓD. DO DOCUMENTO" in df_oficial.columns else pd.Series([False]*len(df_oficial))
        
        if filtro.any():
            idx = df_oficial[filtro].index
            
            # --- CAPTURA DA 1ª VERIFICAÇÃO ---
            if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]).strip() == "":
                data_h = historico_do_doc['Data de Recebimento'].iloc.strftime('%d/%m/%Y')
                df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"] = data_h
                df_oficial.at[idx, "1ª VERIFICAÇÃO EZEQUIAS"] = "OK"

            # ESTRATÉGIA DE RASTREAMENTO ESTIMADO (Se esqueceram o Início da 1ª)
            if pd.isna(df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]) or str(df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]).strip() == "":
                data_h_dt = pd.to_datetime(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"], format='%d/%m/%Y')
                # Estima início como 1 dia útil após o recebimento
                df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"] = (data_h_dt + timedelta(days=1)).strftime('%d/%m/%Y')

            # --- CAPTURA DA 2ª VERIFICAÇÃO ---
            if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]).strip() == "":
                data_1v_str = df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]
                try:
                    data_1v = pd.to_datetime(data_1v_str, format='%d/%m/%Y')
                    proximos_envios = historico_do_doc[historico_do_doc['Data de Recebimento'] > data_1v]
                    if not proximos_envios.empty:
                        segunda_data = proximos_envios['Data de Recebimento'].iloc.strftime('%d/%m/%Y')
                        df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"] = segunda_data
                        df_oficial.at[idx, "2ª VERIFICAÇÃO EZEQUIAS"] = "OK"
                        
                        # Estima início da 2ª verificação caso esquecido
                        if pd.isna(df_oficial.at[idx, "INÍCIO DA 2ª VERIFICAÇÃO DO RESPONSÁVEL"]) or str(df_oficial.at[idx, "INÍCIO DA 2ª VERIFICAÇÃO DO RESPONSÁVEL"]).strip() == "":
                            df_oficial.at[idx, "INÍCIO DA 2ª VERIFICAÇÃO DO RESPONSÁVEL"] = (proximos_envios['Data de Recebimento'].iloc + timedelta(days=1)).strftime('%d/%m/%Y')
                except:
                    pass

            # --- TRATAMENTO DO GARGALO (SEM RETORNO DO SETOR) ---
            if not pd.isna(df_oficial.at[idx, "FIM DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]) and str(df_oficial.at[idx, "FIM DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]).strip() != "":
                if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]).strip() == "":
                    df_oficial.at[idx, "APÓS 2ª VERIFICAÇÃO, ENVIADO PARA ALTERAÇÕES?"] = "AGUARDANDO SETOR"
                    df_oficial.at[idx, "STATUS"] = "VERIFICADO AGUARDA DEVOLUÇÃO SETOR"

            # --- DATA DE APROVAÇÃO FINAL DO ANEXO ---
            if data_aprovacao_word:
                if pd.isna(df_oficial.at[idx, "DATA DE APROVAÇÃO"]) or str(df_oficial.at[idx, "DATA DE APROVAÇÃO"]).strip() == "":
                    df_oficial.at[idx, "DATA DE APROVAÇÃO"] = data_aprovacao_word
                    df_oficial.at[idx, "STATUS"] = "APROVADO"

df_oficial.to_excel(caminho_excel, index=False)
print("Sincronização concluída com estimativas inteligentes de datas!")
