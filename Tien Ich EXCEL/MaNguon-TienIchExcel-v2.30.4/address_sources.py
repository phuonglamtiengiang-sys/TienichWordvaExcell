from pathlib import Path
import json
from address.history.build_history import build as build_history

ROOT = Path(__file__).parent
DIR = ROOT / 'address'
def load(name):
    return json.loads((DIR / (name + '.json')).read_text(encoding='utf-8-sig'))['data']
def indexed(name):
    return {x['code']: x for x in load(name)}
def fullname(x):
    return x['type'] + ' ' + x['name']

def generate():
    op, od, ow = indexed('provinces-legacy'), indexed('districts-legacy'), indexed('wards-legacy')
    np, nw = indexed('provinces'), indexed('wards')
    rows = []
    excluded = []
    for m in load('ward-merge-history'):
        if m['old_ward_code'] is None:
            excluded.append(m)
            continue
        old = ow.get(m['old_ward_code'])
        new = nw.get(m['new_ward_code'])
        if not old or not new:
            raise ValueError('Missing linked unit: ' + str(m))
        assert old['district_code'] == m['old_district_code']
        assert old['province_code'] == m['old_province_code']
        assert new['province_code'] == m['new_province_code']
        rows.append([old['code'], fullname(old), fullname(od[old['district_code']]),
                     fullname(op[old['province_code']]), fullname(new),
                     fullname(np[new['province_code']]), m['scope'], new['code'],
                     new.get('legal_ref') or m.get('legal_ref') or '', '2026.10.8'])
    rows.extend(build_history(rows))
    (DIR/'embedded-rows.json').write_text(json.dumps(rows, ensure_ascii=False),encoding='utf-8')
    (DIR/'excluded-rows.json').write_text(json.dumps(excluded,ensure_ascii=False,indent=2),encoding='utf-8')
    text = (DIR/'mod_Address.vba').read_text(encoding='utf-8')
    text = ''.join(c if ord(c)<128 else '\\u%04x'%ord(c) for c in text)
    (ROOT/'src/mod_Address.txt').write_text(text,encoding='utf-8')
    print('Address mappings:',len(rows),'old units mapped:',len(set(r[0] for r in rows)))

generate()
