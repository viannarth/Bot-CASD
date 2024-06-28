from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import logging

# Configuração do logging para imprimir mensagens de erro
logging.basicConfig(level=logging.ERROR)

try:
    # Autenticação do usuário para utilização da API do Google Drive
    gauth = GoogleAuth()
    # Carregamento de credenciais já salvas (se houver)
    gauth.LoadCredentialsFile("credentials.json")

    if gauth.credentials is None:
        # Autenticação do usuário pela web se não houver credenciais
        gauth.LocalWebserverAuth()
        gauth.SaveCredentialsFile("credentials.json")
    elif gauth.access_token_expired:
        # Autenticação do usuário pela web se os tokens estiverem expirados
        gauth.Refresh()
        gauth.SaveCredentialsFile("credentials.json")
    else:
        # Inicializar as credenciais salvas
        gauth.Authorize()
    
    # Acessar Google Drive
    drive = GoogleDrive(gauth)

    ID_PLANILHA = 'INSIRA AQUI O ID DA PLANILHA'
    # Download da planilha
    planilha = drive.CreateFile({'id': ID_PLANILHA})
    planilha.GetContentFile('Controle de Presença 2024.xlsx')

    print('O download foi realizado com sucesso')
except Exception as erro:
    # Caso houver erro no download, imprimir mensagem de erro no terminal
    logging.error(f"Ocorreu um erro: {erro}")