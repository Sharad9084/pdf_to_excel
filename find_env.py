import os

paths = [
    r"C:\Users\hp\OneDrive\Desktop\fronted",
    r"c:\Users\hp\OneDrive\Desktop\devine"
]

print("=== SEARCHING FOR .ENV FILES ===")
for p in paths:
    if os.path.exists(p):
        for root, dirs, files in os.walk(p):
            for file in files:
                if file.endswith('.env') or file == '.env':
                    full_path = os.path.join(root, file)
                    print(f"Found: {full_path}")
                    with open(full_path, 'r', encoding='utf-8') as f:
                        lines = f.readlines()
                    for line in lines:
                        if 'URL' in line or 'postgres' in line or 'db' in line or 'DATABASE' in line:
                            # print key, mask value
                            parts = line.split('=', 1)
                            if len(parts) == 2:
                                print(f"  {parts[0]} = {parts[1][:15]}...")
                            else:
                                print(f"  {line.strip()[:20]}...")
