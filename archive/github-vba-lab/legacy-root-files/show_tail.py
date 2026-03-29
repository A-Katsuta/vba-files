from pathlib import Path
text = Path(r'C:\github-vba-lab\github-vba-lab\Rrecision_tilt_sensor.html').read_text(encoding='utf-8')
end_segment = text[-200:]
print(repr(end_segment))
