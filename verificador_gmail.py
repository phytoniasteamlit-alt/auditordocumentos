import streamlit as st
import hashlib
import pandas as pd
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

def get_gmail_service():
    # Puxa as credenciais diretamente do painel de Secrets do Streamlit Cloud
    creds = Credentials.from_authorized_user_info(st.secrets["gmail_creds"])
    return build('gmail', 'v1', credentials=creds)

def obter_ou_criar_marcador(service):
    nome_marcador = "🚨 ANEXO DUPLICADO"
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])
    for label in labels:
        if label['name'] == nome_marcador:
            return label['id']
    label_object = {
        "name": nome_marcador,
        "labelListVisibility": "labelShow",
        "messageListVisibility": "show"
    }
    created_label = service.users().labels().create(userId='me', body=label_object).execute()
    return created_label['id']

def marcar_no_gmail(service, message_id, label_id):
    service.users().messages().modify(
        userId='me', 
        id=message_id, 
        body={"addLabelIds": [label_id, "STARRED"]}
    ).execute()

def render_page():
    st.title("📊 Verificador de Anexos Duplicados")
    st.markdown("Busca anexos Word/PDF e marca duplicados com etiqueta vermelha e estrela direto no Gmail.")
    
    if st.button("Buscar e Etiquetar Duplicados"):
        with st.spinner("Analisando caixa de entrada..."):
            try:
                service = get_gmail_service()
                label_id = obter_ou_criar_marcador(service)
                
                # Busca mensagens recentes com anexos
                results = service.users().messages().list(userId='me', q="has:attachment", maxResults=20).execute()
                messages = results.get('messages', [])
                
                if not messages:
                    st.info("Nenhum e-mail com anexo encontrado recentemente.")
                    return
                
                registro_anexos = []
                hashes_vistos = set()
                
                for msg in messages:
                    msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
                    payload = msg_data.get('payload', {})
                    subject = next((header['value'] for header in payload.get('headers', []) if header['name'] == 'Subject'), "Sem Assunto")
                    
                    for part in payload.get('parts', []):
                        filename = part.get('filename')
                        if filename and (filename.endswith('.pdf') or filename.endswith('.docx') or filename.endswith('.doc')):
                            att_id = part['body'].get('attachmentId')
                            attachment = service.users().messages().attachments().get(userId='me', messageId=msg['id'], id=att_id).execute()
                            
                            data = attachment.get('data')
                            file_hash = hashlib.md5(data.encode('utf-8')).hexdigest()
                            
                            status = "Original"
                            if file_hash in hashes_vistos:
                                status = "🚨 DUPLICADO"
                                marcar_no_gmail(service, msg['id'], label_id)
                            else:
                                hashes_vistos.add(file_hash)
                                
                            registro_anexos.append({
                                "Assunto": subject,
                                "Arquivo": filename,
                                "Status": status
                            })
                
                if registro_anexos:
                    df = pd.DataFrame(registro_anexos)
                    st.dataframe(df.style.highlight_between(left="🚨 DUPLICADO", right="🚨 DUPLICADO", axis=1, color="#ffcccc"))
                    st.success("Verificação concluída! Os e-mails duplicados foram marcados no seu Gmail.")
                else:
                    st.info("Nenhum anexo PDF ou Word elegível foi encontrado.")
                    
            except Exception as e:
                st.error(f"Erro ao conectar com o Gmail: {e}")
                st.info("Verifique se as chaves nos Secrets do Streamlit Cloud foram preenchidas.")
