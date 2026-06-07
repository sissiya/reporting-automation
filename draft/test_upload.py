import requests
url='http://127.0.0.1:8000/upload_json'
files={'file': open('Synthèse DC - Incident.json','rb')}
try:
    r = requests.post(url, files=files)
    print(r.status_code)
    print(r.text)
except Exception as e:
    print('ERROR', e)
