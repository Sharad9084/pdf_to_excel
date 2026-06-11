with open(r"C:\Users\hp\OneDrive\Desktop\fronted\app.js", encoding="utf-8") as f:
    lines = f.readlines()

for idx in range(3730, min(3770, len(lines))):
    line_ascii = lines[idx].encode("ascii", errors="replace").decode("ascii")
    print(f"{idx+1}: {line_ascii.strip()}")
