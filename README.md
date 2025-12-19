# LearningMail

Lê e-mails não lidos e baixa os arquivos escolhidos em um diretório

---

# O que faz...

Filtra por label (ou INBOX)

Baixa anexos específicos

Salva em uma pasta local

Marca os e-mails como lidos

---

# Pré-requisitos

Python 3.9+

Ativar Gmail API no Google Cloud

Criar credenciais OAuth e salvar como credentials.json

Instalar libs:

pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib

---

# Configurações importantes no código

Diretório:

DIRETORIO_SAIDA = r'C:\caminho\para\pasta'

---

Extensões válidas:

EXTENSOES_PERMITIDAS = ('.pptx', '.xlsx', '.docx', '.png', '.csv', '.edithere')

adc mais caso precise...

---

Label:

LABEL_ALVO = 'INBOX'  # ou nome exato da label

---

# Como executar

python learningmail.py

Na primeira vez, autorize o Gmail no navegador

O token fica salvo (token.pickle)

Resultado

Até X e-mails processados

Anexos salvos automaticamente

E-mails marcados como lidos após leitura

mit by k.

