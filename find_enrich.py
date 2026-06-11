with open(r"C:\Users\hp\OneDrive\Desktop\fronted\app.js", encoding="utf-8") as f:
    content = f.read()

idx = content.find("function enrichRow")
if idx != -1:
    print(content[idx:idx+1500])
else:
    print("Not found")
