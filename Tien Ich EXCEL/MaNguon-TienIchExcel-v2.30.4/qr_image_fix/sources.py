from pathlib import Path
root=Path(__file__).resolve().parent.parent
p=root/'src/mod_Core.txt'
s=p.read_text(encoding='utf-8-sig')
if 'QRImageExtension(http.ResponseBody, kind)' not in s:
 s=s.replace('description As String, shapes As Object','description As String, shapes As Object, extension As String')
 s=s.replace('    Set fso = CreateObject("Scripting.FileSystemObject")\n    file = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName & ".png")','    extension = QRImageExtension(http.ResponseBody, kind)\n    Set fso = CreateObject("Scripting.FileSystemObject")\n    file = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName & extension)')
p.write_text(s,encoding='utf-8-sig')
(root/'src/mod_QRImageFormat.txt').write_text((root/'qr_image_fix/mod_QRImageFormat.vba').read_text(encoding='utf-8-sig'),encoding='utf-8-sig')
