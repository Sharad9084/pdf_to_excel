import urllib.request
import json
import time

base_url = "http://127.0.0.1:8787"

endpoints = [
    "/api/health",
    "/api/extracted-data/po",
    "/api/extracted-data/agency",
]

for ep in endpoints:
    url = base_url + ep
    print(f"Testing URL: {url}")
    start = time.time()
    try:
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=10) as response:
            status = response.status
            body = response.read()
            elapsed = time.time() - start
            print(f"  Status: {status}")
            print(f"  Time taken: {elapsed:.3f} seconds")
            print(f"  Response length: {len(body)} bytes")
            if len(body) < 1000:
                print(f"  Body: {body.decode('utf-8')}")
            else:
                print(f"  Body (truncated): {body[:200].decode('utf-8')}...")
    except Exception as e:
        print(f"  Error: {e}")
