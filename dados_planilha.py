from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from datetime import datetime

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

key = '1fo1cX43IeT3PE4VlKT3alY8r9muOftHJjsJtXbdHWqw'

class ObterTabela:
    def __init__(self):
        creds = None

        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secret.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        try:
            service = build('sheets', 'v4', credentials=creds)
        except:
            DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
            service = build('sheets', 'v4', credentials=creds, discoveryServiceUrl=DISCOVERY_SERVICE_URL)

        # ler infos sheets
        self.sheet = service.spreadsheets()

    def obter_tabela(self, operacao):

        result = self.sheet.values().get(spreadsheetId=key,range=f'{operacao}!A1:E600').execute()

        values = result.get('values', [])

        return values

    def obter_coluna(self, operacao):
        result = self.sheet.values().get(spreadsheetId=key,
                                         range=f'{operacao}!A1:A600').execute()
        values = result.get('values', [])

        return values

#
# tb = ObterTabela()
# lista =[]
# lista.append(tb.obter_coluna('CLARO'))
# print(len(lista[0]))