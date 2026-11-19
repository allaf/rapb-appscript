# coding=utf-8
from googleapiclient.discovery import build
from google.oauth2 import service_account
import os
import json

# Remplacez par vos identifiants
SERVICE_ACCOUNT_FILE = 'client-gsheet.json'
SCRIPT_ID = '1n9j8KyGBWelhXz5xtSilKTSnlchmDNrU94IKf_ebXcKMEN5-PYgTbYJ6'  # Trouvable dans l'URL du projet Apps Script
OUTPUT_DIR = 'src-appscript'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes = ['https://www.googleapis.com/auth/script.projects',
          'https://www.googleapis.com/auth/spreadsheets']

)

service = build('script', 'v1', credentials=credentials)

# Récupérer le code
request = service.projects().getContent(
    scriptId=SCRIPT_ID,
    versionNumber='4'  # ou un numéro de version spécifique
)
response = request.execute()



try:
    # Récupérer le contenu du projet Apps Script
    request = service.projects().getContent(
        scriptId=SCRIPT_ID,
        versionNumber=0
    )
    response = request.execute()

    print("📋 Réponse de l'API Apps Script reçue.")
    print(f" Nombre de fichiers trouvés : {len(response.get('files', []))}")
    if 'files' not in response:
        print("❌ Aucun fichier trouvé dans la réponse de l'API.")
        exit(1)

    # Sauvegarder chaque fichier
    for file in response['files']:
        file_name = file['name']
        file_type = file['type']
        file_content = file['source']

        # Déterminer l'extension en fonction du type
        if file_type == 'SERVER_JS':
            extension = '.gs'
        elif file_type == 'HTML':
            extension = '.html'
        elif file_type == 'JSON':
            extension = '.json'
        elif file_type == 'CSS':
            extension = '.css'
        else:
            extension = f'.{file_type.lower()}'  # Par défaut

        # Sauvegarder le fichier
        file_path = os.path.join(OUTPUT_DIR, f"{file_name}{extension}")

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(file_content)

        print(f"✅ {file_name}{extension} sauvegardé dans {OUTPUT_DIR}/")

except Exception as e:
    print(f"❌ Erreur : {e}")
    print(f"Détails : {str(e)}")

