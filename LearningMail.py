# -*- coding: utf-8 -*-
"""
Created on Sat Aug  9 14:33:52 2025

@author: K
"""
import os
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from email import message_from_bytes
import base64

# =============================================================================

# Configurações
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
DIRETORIO_SAIDA = '/caminho/para/salvar/arquivos'  # Altere para seu diretório
EXTENSOES_PERMITIDAS = ('.pptx', '.xlsx', 'word', 'png')  # Extensões permitidas
LABEL_ALVO = 'nome_da_sua_label'  # Ex: 'Financeiro', 'Relatórios', IMPORTANTE: usar nome EXATO da label
# Para buscar na Caixa de Entrada padrão, use: LABEL_ALVO = 'INBOX'

# =============================================================================

# =============================================================================

def autenticar_gmail():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return build('gmail', 'v1', credentials=creds)

# =============================================================================

# =============================================================================

def obter_id_label(service, nome_label):
    """Obtém o ID de uma label pelo nome"""
    resultados = service.users().labels().list(userId='me').execute()
    labels = resultados.get('labels', [])
    
    for label in labels:
        if label['name'].lower() == nome_label.lower():
            return label['id']
    
    # Se não encontrou, tenta nomes padrão do Gmail
    nomes_padrao = {
        'inbox': 'INBOX',
        'sent': 'SENT',
        'drafts': 'DRAFT',
        'spam': 'SPAM',
        'trash': 'TRASH'
    }
    
    return nomes_padrao.get(nome_label.lower(), nome_label)

# =============================================================================

# =============================================================================

def salvar_anexo(service, mensagem_id):
    try:
        mensagem = service.users().messages().get(
            userId='me', id=mensagem_id, format='raw').execute()
        
        msg_raw = base64.urlsafe_b64decode(mensagem['raw'])
        mime_msg = message_from_bytes(msg_raw)
        salvou_anexo = False
        
        for part in mime_msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if not part.get('Content-Disposition'):
                continue
            
            filename = part.get_filename()
            if filename:
                # Verifica a extensão do arquivo
                if filename.lower().endswith(EXTENSOES_PERMITIDAS):
                    filepath = os.path.join(DIRETORIO_SAIDA, filename)
                    with open(filepath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    print(f'Arquivo salvo: {filepath}')
                    salvou_anexo = True
        
        if not salvou_anexo:
            print("Nenhum anexo .pptx ou .xlsx encontrado neste e-mail")
    
    except Exception as e:
        print(f'Erro ao processar anexos: {str(e)}')

# =============================================================================

# =============================================================================

def main():
    # Criar diretório se não existir
    os.makedirs(DIRETORIO_SAIDA, exist_ok=True)
    
    service = autenticar_gmail()
    
    # Obter ID da label
    label_id = obter_id_label(service, LABEL_ALVO)
    print(f"Buscando e-mails na label: {LABEL_ALVO} (ID: {label_id})")
    
    # Buscar e-mails não lidos na label específica
    resultados = service.users().messages().list(
        userId='me', 
        labelIds=[label_id],
        q="is:unread", 
        maxResults=5  # Busca até 5 e-mails de uma vez
    ).execute()
    
    if 'messages' in resultados and resultados['messages']:
        print(f"Encontrados {len(resultados['messages'])} e-mail(s) não lido(s)")
        
        for msg in resultados['messages']:
            mensagem_id = msg['id']
            print(f"\nProcessando e-mail ID: {mensagem_id}")
            salvar_anexo(service, mensagem_id)
            
            # Marcar e-mail como lido (remover o rótulo "UNREAD")
            service.users().messages().modify(
                userId='me',
                id=mensagem_id,
                body={'removeLabelIds': ['UNREAD']}
            ).execute()
            print("E-mail marcado como lido")
    else:
        print("Nenhum e-mail não lido encontrado na label especificada")

# =============================================================================

if __name__ == '__main__':
    main()


# =============================================================================

