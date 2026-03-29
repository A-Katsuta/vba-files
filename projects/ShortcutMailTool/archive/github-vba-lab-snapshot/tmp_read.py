from pathlib import Path
path = Path(__file__).resolve().parent / 'ShortcutMailTool.bas'
text = path.read_text(encoding='utf-8', errors='ignore')
print(text[60:140].encode('utf-8'))
