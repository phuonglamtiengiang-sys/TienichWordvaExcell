from pathlib import Path
root = Path(__file__).parent
for name in ('mod_Translate', 'mod_TranslateProviders', 'frmTranslate', 'frmTranslateProgress'):
    text = (root/f'translate/{name}.vba').read_text(encoding='utf-8')
    text = ''.join(c if ord(c)<128 else '\\u%04x'%ord(c) for c in text)
    (root/f'src/{name}.txt').write_text(text, encoding='utf-8-sig')
