from pathlib import Path

ROOT = Path(__file__).parent
for name in ('mod_CompareTables', 'frmCompareTables'):
    text = (ROOT/'compare'/f'{name}.vba').read_text(encoding='utf-8')
    text = ''.join(c if ord(c) < 128 else '\\u%04x' % ord(c) for c in text)
    (ROOT/'src'/f'{name}.txt').write_text(text, encoding='utf-8-sig')

