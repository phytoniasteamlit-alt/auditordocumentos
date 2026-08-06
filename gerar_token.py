from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ['https://googleapis.com']

print("Iniciando fluxo de autenticação...")
flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
creds = flow.run_local_server(port=0)

with open('token.json', 'w') as token:
    token.write(creds.to_json())
    
print("Sucesso! O arquivo 'token.json' foi gerado na sua pasta.")
