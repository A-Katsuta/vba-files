from pathlib import Path

HTML_PATH = Path(__file__).resolve().parents[2] / 'web-assets' / 'rrecision' / 'Rrecision_tilt_sensor.html'


def main() -> None:
    text = HTML_PATH.read_text(encoding='utf-8')
    end_segment = text[-200:]
    print(repr(end_segment))


if __name__ == '__main__':
    main()
