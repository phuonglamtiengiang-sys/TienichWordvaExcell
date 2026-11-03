from pathlib import Path
root = Path(__file__).parent
src = root / 'src'
text = (root/'undo/mod_Undo.vba').read_text(encoding='utf-8')
text = ''.join(c if ord(c)<128 else '\\u%04x'%ord(c) for c in text)
(src/'mod_Undo.txt').write_text(text, encoding='utf-8-sig')
for name, begin in [('mod_XuLyVanBan','UndoBegin inputRange, columns, "Xu ly van ban"'), ('mod_DoiSo','UndoBegin inputRange, 1, "Doi so thanh chu"'), ('mod_HanhChinhVN',None)]:
    path = src/(name+'.txt')
    text = path.read_text(encoding='utf-8-sig')
    marker = '    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents'
    if begin:
        text = text.replace(marker, '    '+begin+'\n'+marker)
    else:
        start=text.index('Public Sub Ribbon_CheckCCCD')
        prefix,text=text[:start],text[start:]
        text=text.replace(marker,'    UndoBegin selected, 3, "Phan tich CCCD"\n'+marker,1)
        start=text.index('Public Sub Ribbon_LookupMST')
        first,second=text[:start],text[start:]
        second=second.replace(marker,'    UndoBegin selected, 2, "Tra MST"\n'+marker,1)
        text=prefix+first+second
    # Only these four data-writing actions have Done labels in these modules.
    text=text.replace('    Exit Sub\nFailed:\n    reason = Err.Description: Resume Done','    UndoCommit\n    Exit Sub\nFailed:\n    reason = Err.Description: Resume Done')
    path.write_text(text,encoding='utf-8-sig')
