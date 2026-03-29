from pathlib import Path

HTML_PATH = Path(__file__).resolve().parents[2] / 'web-assets' / 'rrecision' / 'Rrecision_tilt_sensor.html'
MARKER = '\r\n        </div>\r\n    </div>\r\n\r\n    <script>'
LABEL_HTML = '\r\n            <div class="sensor-label">JJ-07</div>'

def main() -> None:
    text = HTML_PATH.read_text(encoding='utf-8')
    if '<div class="sensor-label">JJ-07</div>' not in text and MARKER in text:
        updated = text.replace(MARKER, LABEL_HTML + MARKER)
        HTML_PATH.write_text(updated, encoding='utf-8')


if __name__ == '__main__':
    main()
