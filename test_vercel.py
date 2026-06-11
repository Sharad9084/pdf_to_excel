import urllib.request
import json
import time

vercel_url = "https://mpro-ai.vercel.app"

print("Testing Vercel home page...")
try:
    start = time.time()
    req = urllib.request.Request(vercel_url)
    with urllib.request.urlopen(req, timeout=10) as response:
        print(f"  Home page status: {response.status}")
        print(f"  Time taken: {time.time() - start:.3f} seconds")
except Exception as e:
    print(f"  Home page error: {e}")

print("\nTesting Vercel health API...")
try:
    start = time.time()
    req = urllib.request.Request(vercel_url + "/api/health")
    with urllib.request.urlopen(req, timeout=10) as response:
        body = response.read().decode('utf-8')
        print(f"  Health status: {response.status}")
        print(f"  Time taken: {time.time() - start:.3f} seconds")
        print(f"  Body: {body}")
except Exception as e:
    print(f"  Health API error: {e}")
