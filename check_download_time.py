import os
import datetime

download_dir = r"C:\Users\hp\Downloads"
files = [
    "invoice-reconciliation-thirdPartyInvoice.csv",
    "invoice-reconciliation-thirdPartyInvoice (1).csv",
    "invoice-reconciliation-thirdPartyInvoice (2).csv",
]

print("=== CHECKING DOWNLOADED CSV TIMESTAMPS ===")
for f in files:
    path = os.path.join(download_dir, f)
    if os.path.exists(path):
        mtime = os.path.getmtime(path)
        ctime = os.path.getctime(path)
        size = os.path.getsize(path)
        print(f"File: {f}")
        print(f"  Size: {size} bytes")
        print(f"  Created: {datetime.datetime.fromtimestamp(ctime)}")
        print(f"  Modified: {datetime.datetime.fromtimestamp(mtime)}")
    else:
        print(f"File: {f} - Not found")
