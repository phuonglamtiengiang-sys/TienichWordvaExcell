"""Extract unchanged encoding tables from UniKey UVConverter 1.1.3b (GPL-2.0-or-later)."""
from pathlib import Path
import ast
import hashlib
import re

root = Path(__file__).resolve().parent
source = root / 'vendor/unikey/upstream/uvconv/vnconv/data.cpp'
raw = source.read_bytes()
text = raw.decode('latin1')
text = re.sub(r'/\*.*?\*/|//[^\r\n]*', '', text, flags=re.S)

def table(name, nested=False):
    start = text.index('{', text.index(name)) + 1
    if nested:
        start = text.index('{', start) + 1
    end = text.index('}', start)
    values = []
    for token in re.findall(r"'(?:\\.|[^'\\])*'|0[xX][0-9a-fA-F]+|\d+", text[start:end]):
        value = ord(ast.literal_eval(token)) if token.startswith("'") else int(token, 0)
        values.append(value)
    assert len(values) == 213, (name, len(values))
    return values

tables = [('Tcvn', table('SingleByteTables', True)), ('Vni', table('DoubleByteTables', True)), ('Unicode', table('UnicodeTable'))]
header = '''// SPDX-License-Identifier: GPL-2.0-or-later
// UniKey VnConv, Copyright (C) 1998-2002 Pham Kim Long.
// Generated without table changes from UVConverter 1.1.3b vnconv/data.cpp.
// Do not edit: run word/generate_unikey_tables.py. See vendor/unikey/NOTICE.md.
namespace TienIchWord {
    internal static class UniKeyTables {
'''
for name, values in tables:
    header += '        internal static readonly ushort[] ' + name + ' = {\n'
    for i in range(0, len(values), 12):
        header += '            ' + ','.join('0x%04X' % v for v in values[i:i+12]) + ',\n'
    header += '        };\n'
header += '    }\n}\n'
(root / 'UniKeyTables.generated.cs').write_text(header, encoding='utf-8')
print('Extracted 3 x 213 UniKey entries; data.cpp SHA256=' + hashlib.sha256(raw).hexdigest())
