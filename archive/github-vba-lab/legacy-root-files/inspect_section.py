from pathlib import Path
text = Path(r'C:\github-vba-lab\github-vba-lab\Rrecision_tilt_sensor.html').read_text(encoding='utf-8')
print(repr(text[text.index('</div>\r\n    </div>'):text.index('</div>\r\n    </div>')+40]))
