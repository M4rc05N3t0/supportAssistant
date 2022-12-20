import requests
import getKey

url = 'http://104.197.32.119'

token = getKey.get_key()

header = {'Content-Type':'application/json',
          'Authorization': 'Bearer ' + token}


def inputAPI(serial, usuario, modelo, data):
    endpoint = url + f'/serial/{serial}'

    body = {"serial": f"{serial}",
            "nome_colaborador": f"{usuario}",
            "tipo_equipamento": f"{modelo}",
            "acao": f"Equipamento em uso",
            "chamado": "",
            "analista_responsavel": "SupportAssistant",
            "conferencia": "",
            "data": f"{data}",
            "obs": "",
    }
    input= requests.request('PUT', json=body, headers=header, url=endpoint)

def get (serial):
    endpoint = url+f"/serial/{serial}"
    get = requests.request('GET',url=endpoint )
    return get.json()





