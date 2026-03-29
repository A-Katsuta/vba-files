from pathlib import Path
path = Path(r'C:\github-vba-lab\github-vba-lab\Rrecision_tilt_sensor.html')
text = path.read_text(encoding='utf-8')
marker = '\r\n        </div>\r\n    </div>\r\n\r\n    <script>'
if '<div class="sensor-label">JJ-07</div>' not in text and marker in text:
    text = text.replace(marker, '\r\n            <div class="sensor-label">JJ-07</div>' + marker)
path.write_text(text, encoding='utf-8')
