from pathlib import Path
root = Path(__file__).parent
text = (root/'printing/mod_PrintSelection.vba').read_text(encoding='utf-8')
text = ''.join(c if ord(c)<128 else '\\u%04x'%ord(c) for c in text)
(root/'src/mod_PrintSelection.txt').write_text(text, encoding='utf-8-sig')
