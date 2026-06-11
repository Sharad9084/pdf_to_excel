import urllib.request
import json
import time

render_url = "https://pdf-to-excel-5ota.onrender.com"

endpoints = [
    "/api/health",
    "/",
]

for ep in endpoints:
    url = render_url + ep
    print(f"Testing Render URL: {url}")
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
            print(f"  Body: {body.decode('utf-8')[:200]}")
    except Exception as e:
        print(f"  Error: {e}")
