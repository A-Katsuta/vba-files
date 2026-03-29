from pathlib import Path
text = Path(__file__).resolve().parent / 'ShortcutMailTool.bas'.read_bytes()
print(text[60:120])
