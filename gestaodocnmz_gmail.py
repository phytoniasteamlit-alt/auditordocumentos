import os
import zipfile
import mailbox
import io
import re
import pandas as pd
from docx import Document
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime

def formatar_data_gmail(data_cabecalho):
    try:
        if data_cabecalho:
            dt = parsedate_to_datetime(data_cabecalho)
            return dt.strftime("%d/%m/%Y")
    except:
        pass
    return None

def extrair_palavras_chave(nome_documento):
    """Gera tokens para busca textual, ignorando conectores, para evitar colisões"""
    nome_limpo = re.sub(r'[^\w\s]', '', str(nome_documento)).upper()
    palavras = [p for p in nome_limpo.split() if len(p) > 3 and p not in ["PARA", "COMO", "COMTN", "DELE"]]
    return words[:2] # Retorna as duas palavras mais marcantes (ex: 'BANHO', 'LEITO')

def rodar_auditoria_anti_colisao(caminho_excel, caminho_zip):
    df_oficial = pd.read_excel(caminho_excel, dtype=str)
    df_oficial.columns = df_oficial.columns.str.strip()
    
    SEU_EMAIL = "documentos.soc2@gmail.com"
    
    with zipfile.ZipFile(caminho_zip, 'r') as z:
        path_mbox = [f for f in z.namelist() if f.endswith('.mbox')]
        with z.open(path_mbox) as mbox_file:
            mbox = mailbox.mbox(io.BytesIO(mbox_file.read()))
            
            for idx, linha in df_oficial.iterrows():
                codigo_doc = str(linha["CÓD. DO DOCUMENTO"]).strip()
                nome_doc = str(linha["NOME DO DOCUMENTO"]).strip()
                versao_planilha = str(linha["VERSÃO"]).strip().lower()
                
                # Se não tem código, o robô adota a estratégia de buscar puramente por palavras-chave do nome
                se_nao_tem_codigo = pd.isna(codigo_doc) or codigo_doc in ["nan", ""]
                
                palavras_chave = extrair_palavras_chave(nome_doc)
                if not keywords:
                    continue
                
                emails_filtrados_compostos = []
                
                for msg in mbox:
                    assunto = str(msg["subject"]).upper()
                    
                    # Regra de busca 1: Se tem código, o código DEVE estar no assunto
                    if not se_nao_tem_codigo and codigo_doc.upper() not in assunto:
                        continue
                        
                    # Regra de busca 2: Pelo menos uma palavra marcante do nome deve bater (Evita o bug do POP_EMTN003)
                    if not any(p in assunto for p in palavras_chave):
                        continue
                        
                    # Regra de busca 3: Tratamento de Versão se estiver preenchida
                    if versao_planilha != "nan" and versao_planilha != "":
                        # Se a planilha diz 2ª versão, ignora e-mails que falem explicitamente de 1ª versão
                        if "2" in versao_planilha and "1ª" in assunto:
                            continue
                            
                    emails_filtrados_compostos.append(msg)
                
                if not emails_filtrados_compostos:
                    continue
                
                # -----------------------------------------------------------------
                # EXECUÇÃO DO PREENCHIMENTO SEGURO (MÁQUINA DE ESTADOS)
                # -----------------------------------------------------------------
                primeiro_email = emails_filtrados_compostos[0]
                
                if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]).strip() == "":
                    df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"] = formatar_data_gmail(primeiro_email["date"])
                
                df_oficial.at[idx, "1ª VERIFICAÇÃO EZEQUIAS"] = "OK"
                
                # [GARGALO DATA MANUAL]: Cálculo baseado no parâmetro cronológico inferido
                if pd.isna(df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]) or str(df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]).strip() == "":
                    data_ref = df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]
                    if pd.notna(data_ref):
                        dt_ref = datetime.strptime(str(data_ref), '%d/%m/%Y')
                        df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"] = (dt_ref + timedelta(days=1)).strftime('%d/%m/%Y')

                # Busca a segunda verificação comparando remetentes posteriores
                emails_2v = [m for m in emails_filtrados_compostos if SEU_EMAIL not in str(m["from"]) and m != primeiro_email]
                if emails_2v:
                    segundo_email = emails_2v[0]
                    if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]).strip() == "":
                        df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"] = formatar_data_gmail(segundo_email["date"])
                        df_oficial.at[idx, "2ª VERIFICAÇÃO EZEQUIAS"] = "OK"

    df_oficial.to_excel(caminho_excel, index=False)
    print("Sincronização anti-colisão finalizada!")
    return df_oficial
