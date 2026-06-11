import urllib.request
import json
import time

base_url = "http://127.0.0.1:8787"

# 1. Sign in
login_url = base_url + "/api/auth/signin"
login_data = json.dumps({
    "username": "auditor@mpro.com",
    "password": "Auditor@2026"
}).encode('utf-8')

print("Signing in...")
try:
    req = urllib.request.Request(login_url, data=login_data, headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req) as response:
        res = json.loads(response.read().decode('utf-8'))
        token = res["token"]
        print(f"  Sign in successful! Token: {token[:10]}...")
except Exception as e:
    print(f"  Sign in error: {e}")
    sys.exit(1)

# 2. Get cases
cases_url = base_url + "/api/cases"
print("\nFetching cases...")
try:
    req = urllib.request.Request(cases_url, headers={
        'Authorization': f'Bearer {token}'
    })
    start = time.time()
    with urllib.request.urlopen(req) as response:
        body = response.read()
        elapsed = time.time() - start
        print(f"  Status: {response.status}")
        print(f"  Time taken: {elapsed:.3f} seconds")
        print(f"  Response length: {len(body)} bytes")
        res_cases = json.loads(body.decode('utf-8'))
        cases = res_cases.get("cases", [])
        print(f"  Found {len(cases)} cases:")
        for c in cases:
            print(f"    - Name: '{c.get('name')}', ID: {c.get('id')}")
except Exception as e:
    print(f"  Fetch cases error: {e}")
