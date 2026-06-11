import urllib.request
import json

url = "http://127.0.0.1:8787/api/extracted-data/thirdPartyInvoice"
try:
    with urllib.request.urlopen(url) as response:
        payload = json.loads(response.read().decode('utf-8'))
        rows = payload.get("data", [])
        if rows:
            print("First row details:")
            for k, v in rows[0].items():
                print(f"  {k}: {v}")
        else:
            print("No rows found.")
except Exception as e:
    print(f"Error: {e}")
