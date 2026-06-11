import urllib.request
import json
import time

url = "http://127.0.0.1:8787/api/extracted-data/thirdPartyInvoice"
print(f"Requesting: {url}")
start = time.time()
try:
    with urllib.request.urlopen(url) as response:
        body = response.read()
        elapsed = time.time() - start
        print(f"  Status: {response.status}")
        print(f"  Time taken: {elapsed:.3f} seconds")
        print(f"  Response size: {len(body)} bytes")
        payload = json.loads(body.decode('utf-8'))
        rows = payload.get("data", [])
        print(f"  Rows returned in API payload: {len(rows)}")
        
        # Analyze vendor distribution in the API payload
        vendors = {}
        for r in rows:
            vendor = r.get("Third Party Vendor Name") or r.get("Broadcaster Name") or "Unknown"
            vendors[vendor] = vendors.get(vendor, 0) + 1
            
        print("\nVendor distribution in API payload:")
        for v, count in vendors.items():
            print(f"  - {v}: {count} rows")
            
except Exception as e:
    print(f"  Error: {e}")
