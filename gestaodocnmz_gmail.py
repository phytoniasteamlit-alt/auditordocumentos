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

def extrair_data_aprovacao_interna_memoria(conteudo_binario):
    """Abre o Word direto da memória e lê a terceira linha do cabeçalho da Norma Zero"""
    try:
        doc = Document(io.BytesIO(conteudo_binario))
        for tabela in doc.tables:
            for linha in tabela.rows:
                texto_linha = [celula.text.strip() for celula in linha.cells]
                text_completo = " ".join(texto_linha)
                if "Data aprovação:" in text_completo:
                    # Isola a string após o rótulo
                    data = text_completo.split("Data aprovação:")[-1].split("Validade:")[0].strip()
                    # Valida se é uma data preenchida real e não a máscara padrão
                    if data and "dd/mm" not in data.lower() and "/" in data:
                        return data
    except:
        return None
    return None

def rodar_auditoria_hospital_inteligente(caminho_excel, caminho_zip):
    df_oficial = pd.read_excel(caminho_excel, dtype=str)
    df_oficial.columns = df_oficial.columns.str.strip()
    
    SEU_EMAIL = "documentos.soc2@gmail.com"
    
    # Abre o arquivo de 1,2 GB vindo do WhatsApp
    with zipfile.ZipFile(caminho_zip, 'r') as z:
        path_mbox = [f for f in z.namelist() if f.endswith('.mbox')]
        if not path_mbox:
            print("Erro: Arquivo .mbox não encontrado dentro do ZIP!")
            return
            
        with z.open(path_mbox[0]) as mbox_file:
            mbox = mailbox.mbox(io.BytesIO(mbox_file.read()))
            
            for idx, linha in df_oficial.iterrows():
                codigo_doc = str(linha["CÓD. DO DOCUMENTO"]).strip()
                if pd.isna(codigo_doc) or codigo_doc in ["nan", ""]:
                    continue
                
                # Agrupa todos os e-mails associados ao código do documento
                emails_do_doc = []
                for msg in mbox:
                    if codigo_doc in str(msg["subject"]):
                        emails_do_doc.append(msg)
                
                if not emails_do_doc:
                    continue
                
                # -----------------------------------------------------------------
                # FASE 1: ENTRADA DO FLUXO (1ª VERIFICAÇÃO)
                # -----------------------------------------------------------------
                primeiro_email = emails_do_doc[0]
                
                # Só preenche a data de recebimento se estiver em branco (Garantia de Segurança)
                if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]).strip() == "":
                    data_1v = formatar_data_gmail(primeiro_email["date"])
                    df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"] = data_1v
                
                # Garante o OK na sua coluna de controle
                df_oficial.at[idx, "1ª VERIFICAÇÃO EZEQUIAS"] = "OK"
                
                # [TRATAMENTO DE GARGALO]: Se as meninas esqueceram o Início da 1ª Verificação
                if pd.isna(df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]) or str(df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]).strip() == "":
                    data_ref = df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 1ª VERIFICAÇÃO"]
                    if pd.notna(data_ref):
                        dt_ref = datetime.strptime(str(data_ref), '%d/%m/%Y')
                        # Estima de forma inteligente para 1 dia após a sua entrega na pasta
                        df_oficial.at[idx, "INÍCIO DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"] = (dt_ref + timedelta(days=1)).strftime('%d/%m/%Y')

                # -----------------------------------------------------------------
                # FASE 2: RETORNO DO SETOR (2ª VERIFICAÇÃO)
                # -----------------------------------------------------------------
                # Filtra e-mails de entrada vindos do setor após o início do processo
                emails_posteriores_entrada = [
                    m for m in emails_do_doc 
                    if SEU_EMAIL not in str(m["from"]) and m != primeiro_email
                ]
                
                if emails_posteriores_entrada:
                    segundo_email_entrada = emails_posteriores_entrada[0]
                    
                    if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]).strip() == "":
                        data_2v = formatar_data_gmail(segundo_email_entrada["date"])
                        df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"] = data_2v
                        df_oficial.at[idx, "2ª VERIFICAÇÃO EZEQUIAS"] = "OK"
                    
                    # [TRATAMENTO DE GARGALO]: Se esqueceram o início da 2ª verificação
                    if pd.isna(df_oficial.at[idx, "INÍCIO DA 2ª VERIFICAÇÃO DO RESPONSÁVEL"]) or str(df_oficial.at[idx, "INÍCIO DA 2ª VERIFICAÇÃO DO RESPONSÁVEL"]).strip() == "":
                        data_ref_2v = df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]
                        if pd.notna(data_ref_2v):
                            dt_ref_2v = datetime.strptime(str(data_ref_2v), '%d/%m/%Y')
                            df_oficial.at[idx, "INÍCIO DA 2ª VERIFICAÇÃO DO RESPONSÁVEL"] = (dt_ref_2v + timedelta(days=1)).strftime('%d/%m/%Y')

                # -----------------------------------------------------------------
                # FASE 3: ENGENHARIA DE STATUS E FECHAMENTO (APROVAÇÃO VS GARGALO)
                # -----------------------------------------------------------------
                ultimo_email = emails_do_doc[-1]
                data_aprovacao_word = None
                
                # Tenta extrair a data real de dentro do último documento modificado
                if ultimo_email.is_multipart():
                    for part in ultimo_email.walk():
                        filename = part.get_filename()
                        if filename and filename.endswith(".docx") and codigo_doc in filename:
                            data_aprovacao_word = extrair_data_aprovacao_interna_memoria(part.get_payload(decode=True))
                
                # Se achou a data preenchida dentro do Word -> Ciclo Concluído!
                if data_aprovacao_word:
                    if pd.isna(df_oficial.at[idx, "DATA DE APROVAÇÃO"]) or str(df_oficial.at[idx, "DATA DE APROVAÇÃO"]).strip() == "":
                        df_oficial.at[idx, "DATA DE APROVAÇÃO"] = data_aprovacao_word
                        df_oficial.at[idx, "STATUS"] = "APROVADO"
                
                # Se NÃO achou a data no Word -> O processo caiu no gargalo do setor externo
                else:
                    # Se as meninas já terminaram a 1ª verificação mas a 2ª ainda não chegou
                    if not pd.isna(df_oficial.at[idx, "FIM DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]) and str(df_oficial.at[idx, "FIM DA 1ª VERIFICAÇÃO DO RESPONSÁVEL"]).strip() != "":
                        if pd.isna(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]) or str(df_oficial.at[idx, "DATA DE RECEBIMENTO PARA 2ª VERIFICAÇÃO"]).strip() == "":
                            df_oficial.at[idx, "APÓS 2ª VERIFICAÇÃO, ENVIADO PARA ALTERAÇÕES?"] = "AGUARDANDO SETOR"
                            df_oficial.at[idx, "STATUS"] = "VERIFICADO AGUARDA DEVOLUÇÃO SETOR"

    # Salva as alterações na planilha local
    df_oficial.to_excel(caminho_excel, index=False)
    print("Mapeamento de estados e gargalos executado com sucesso!")
    return df_oficial
