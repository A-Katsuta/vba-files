import re
with open(r"build/unpacked/xl/vbaProject.bin","rb") as f:
    data = f.read()
for m in re.finditer(b"Mail", data):
    print(m.start(), data[m.start():m.start()+40])
