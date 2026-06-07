import json
import urllib.request

url = "http://127.0.0.1:8001/generate_ppt_file"
payload = {"json_path": "Synthèse DC - Incident.json", "filename": "test_report_from_api.pptx"}
req = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers={'Content-Type':'application/json'})
with urllib.request.urlopen(req) as resp:
    print(resp.read().decode('utf-8'))
