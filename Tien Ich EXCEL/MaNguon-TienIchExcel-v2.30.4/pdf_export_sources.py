from pathlib import Path
root = Path(__file__).parent
text = (root/'pdf_export/mod_ExportPDF.vba').read_text(encoding='utf-8')
text = ''.join(c if ord(c)<128 else '\\u%04x'%ord(c) for c in text)
(root/'src/mod_ExportPDF.txt').write_text(text, encoding='utf-8-sig')
