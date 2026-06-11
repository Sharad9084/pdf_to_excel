with open(r"C:\Users\hp\OneDrive\Desktop\fronted\app.js", 'r', encoding='utf-8') as f:
    lines = f.readlines()
for idx, line in enumerate(lines):
    if "function exportCsv" in line:
        print(f"Line {idx+1}: {line.strip()}")
