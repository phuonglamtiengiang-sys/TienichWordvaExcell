from pathlib import Path
import sys
import zipfile
root=Path(__file__).resolve().parent.parent
out=root/'output/Word-2.30.4'; out.mkdir(parents=True,exist_ok=True)
(out/'HUONG-DAN.md').write_text((root/'word/HUONG-DAN.md').read_text(encoding='utf-8-sig'),encoding='utf-8-sig')
(out/'LICENSE-GPL-2.0.txt').write_bytes((root/'word/LICENSE-GPL-2.0.txt').read_bytes())
(out/'THIRD-PARTY-NOTICE.md').write_bytes((root/'word/vendor/unikey/NOTICE.md').read_bytes())
(out/'THAM-KHAO.md').write_bytes((root/'word/THAM-KHAO.md').read_bytes())
source=out/'MaNguon-Word-v2.30.4.zip'
with zipfile.ZipFile(source,'w',zipfile.ZIP_DEFLATED) as z:
    for f in (root/'word').rglob('*'):
        if f.is_file() and '__pycache__' not in f.parts and f.suffix!='.snk': z.write(f,f.relative_to(root).as_posix())
    for name in ['test_word_text_positions.ps1','test_word_ribbon_schema.ps1','test_word_editing.ps1','test_unikey_tables.py','test_unikey_tables.ps1','test_word_installed_load.ps1','test_word_icons.ps1']:
        f=root/'tests'/name
        if f.exists(): z.write(f,'tests/'+name)
    for name in ['branding/build_logo.py','output/branding/TienIchWord-Logo-Ngang.png','output/branding/TienIchWord-Logo.svg','output/branding/TienIchWord-AboutHeader.png']:
        f=root/name
        if f.exists(): z.write(f,name)
    z.write(root/'about/support-momo-original.png','about/support-momo-original.png')
refs=root/'output/TienIchWord-v2.30.4-References.zip'
with zipfile.ZipFile(refs,'w',zipfile.ZIP_DEFLATED) as z:
    for f in [out/'THAM-KHAO.md',out/'THIRD-PARTY-NOTICE.md',source]: z.write(f,f.name)
release=root/'output/TienIchWord-v2.30.4.zip'
required=['CaiDat-TienIchWord-2.30.4-Setup.exe','TienIchWord.ico','TienIchWord.dll','HUONG-DAN.md','LICENSE-GPL-2.0.txt','THIRD-PARTY-NOTICE.md','THAM-KHAO.md','MaNguon-Word-v2.30.4.zip']
if '--source-only' not in sys.argv and all((out/name).is_file() for name in required):
    temporary=release.with_name(release.stem+'.tmp.zip')
    try:
        with zipfile.ZipFile(temporary,'w',zipfile.ZIP_DEFLATED) as z:
            for name in required: z.write(out/name,'TienIchWord-v2.30.4/'+name)
        temporary.replace(release)
    finally:
        if temporary.exists(): temporary.unlink()
print(source,source.stat().st_size)
print(refs,refs.stat().st_size)
if release.exists(): print(release,release.stat().st_size)
