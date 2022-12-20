import requests

url = 'http://104.197.32.119/login'


header = {'Content-Type':'application/json'}

usuario = 'admin'
senha = 'crFZG+MW,(d;--EwVJaeFyU?g8?Mg2UD'

def get_key():
    body = {
    "login": f"{usuario}",
    "senha": f"{senha}"
    }

    get = requests.request('POST', json=body, headers=header, url=url)
    return_json = get.json()
    key = return_json['acess_token']
    return key
