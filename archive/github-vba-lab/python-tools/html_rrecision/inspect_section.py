from pathlib import Path

HTML_PATH = Path(__file__).resolve().parents[2] / 'web-assets' / 'rrecision' / 'Rrecision_tilt_sensor.html'


def main() -> None:
    text = HTML_PATH.read_text(encoding='utf-8')
    focus = '</div>\r\n    </div>'
    if focus in text:
        start = text.index(focus)
        snippet = text[start:start + 40]
        print(repr(snippet))
    else:
        print('target marker not found')


if __name__ == '__main__':
    main()
