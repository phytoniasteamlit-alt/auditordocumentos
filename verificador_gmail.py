import hashlib
import os
import json
import base64
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = ['https://googleapis.com']

def get_gmail_service():
    # Puxa as credenciais das variáveis de ambiente do GitHub
    creds_json = os.environ.get("GMAIL_CREDS_JSON")
    token_json = os.environ.get("GMAIL_TOKEN_JSON")
    
    creds = None
    if token_json:
        creds = Credentials.from_authorized_user_info(json.loads(token_json), SCOPES)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if creds_json:
                # Usa as credenciais do aplicativo de computador (Desktop) que baixamos primeiro
                from google_auth_oauthlib.flow import InstalledAppFlow
                # Como o GitHub Actions roda sem tela, o token precisa ser gerado localmente uma vez.
                # Para simplificar na nuvem sem precisar abrir navegador, vamos usar o token direto se já existir.
                pass
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

def rodar_verificacao():
    print("Iniciando verificação de anexos duplicados...")
    try:
        service = get_gmail_service()
        if not service:
            print("Erro: Credenciais não encontradas ou inválidas.")
            return
            
        label_id = obter_ou_criar_marcador(service)
        
        # Busca e-mails recentes com anexos (últimas 48h para garantir)
        results = service.users().messages().list(userId='me', q="has:attachment", maxResults=30).execute()
        messages = results.get('messages', [])
        
        if not messages:
            print("Nenhum e-mail com anexo encontrado recentemente.")
            return
            
        hashes_vistos = set()
        
        for msg in messages:
            msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
            payload = msg_data.get('payload', {})
            
            # Navega pelas partes do e-mail para achar os anexos
            parts = [payload]
            if 'parts' in payload:
                parts = payload['parts']
                
            for part in parts:
                filename = part.get('filename')
                if filename and (filename.endswith('.pdf') or filename.endswith('.docx') or filename.endswith('.doc')):
                    att_id = part['body'].get('attachmentId')
                    if not att_id:
                        continue
                        
                    attachment = service.users().messages().attachments().get(
                        userId='me', messageId=msg['id'], id=att_id).execute()
                    
                    data = attachment.get('data')
                    # Decodifica o anexo para gerar o hash real do arquivo
                    file_bytes = base64.urlsafe_b64decode(data.encode('UTF-8'))
                    file_hash = hashlib.md5(file_bytes).hexdigest()
                    
                    if file_hash in hashes_vistos:
                        print(f"🚨 Duplicado detectado: {filename}. Marcando no Gmail...")
                        marcar_no_gmail(service, msg['id'], label_id)
                    else:
                        hashes_vistos.add(file_hash)
                        print(f"Arquivo original verificado: {filename}")
                        
        print("Verificação concluída com sucesso!")
    except Exception as e:
        print(f"Erro durante a execução: {e}")

if __name__ == "__main__":
    rodar_verificacao()
