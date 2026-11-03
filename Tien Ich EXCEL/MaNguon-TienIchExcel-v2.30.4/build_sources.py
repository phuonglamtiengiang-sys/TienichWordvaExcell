from pathlib import Path
import re, unicodedata, json

ROOT = Path(__file__).parent
SRC = ROOT / 'src'
def put(name, text):
    # VBA editor uses ANSI: encode non-ASCII literals as runtime Unicode escapes.
    text = ''.join(c if ord(c)<128 else '\\u%04x'%ord(c) for c in text)
    (SRC/(name+'.txt')).write_text(text.strip()+'\n',encoding='utf-8-sig')
def original(name):
    return (ROOT/'source-original'/(name+'.txt')).read_text(encoding='utf-8-sig')

common = r'''
Option Explicit
Public Const MAX_BATCH As Long = 1000
Public Const IMAGE_TAG As String = "TienIchExcel:2:"

Public Function DigitsOnly(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    DigitsOnly = Not (s Like "*[!0-9]*")
End Function

Public Function SelectedInput(ByRef result As Range, Optional ByVal outputCols As Long = 0, Optional ByVal oneColumn As Boolean = True) As Boolean
    Dim a As Range, c As Range, dest As Range
    On Error GoTo Failed
    If TypeName(Selection) <> "Range" Then Err.Raise 5, , "Hay chon vung o du lieu."
    Set result = Selection
    If result.CountLarge > MAX_BATCH Then Err.Raise 5, , "Moi lan xu ly toi da 1000 o. Hay chon vung du lieu cu the."
    If result.Parent.ProtectContents Then Err.Raise 5, , "Trang tinh dang duoc bao ve."
    For Each a In result.Areas
        If oneColumn And a.Columns.Count <> 1 Then Err.Raise 5, , "Hay chon mot cot du lieu nguon."
        If a.Column + a.Columns.Count - 1 + outputCols > result.Parent.Columns.Count Then Err.Raise 5, , "Khong du cot de ghi ket qua."
    Next a
    For Each c In result.Cells
        If c.MergeCells Then Err.Raise 5, , "Khong xu ly o gop."
        If outputCols > 0 Then
            Set dest = c.Offset(0, 1).Resize(1, outputCols)
            If Not Intersect(dest, result) Is Nothing Then Err.Raise 5, , "Vung ket qua trung voi vung nguon."
            For Each a In dest.Cells
                If a.MergeCells Then Err.Raise 5, , "Vung ket qua co o gop."
                If a.HasFormula Then Err.Raise 5, , "Cac cot ket qua ben phai co cong thuc. Khong ghi de."
                If IsError(a.Value2) Then Err.Raise 5, , "Cac cot ket qua ben phai co gia tri loi. Khong ghi de."
                If Len(CStr(a.Value2)) > 0 Then Err.Raise 5, , "Cac cot ket qua ben phai phai trong de tranh ghi de du lieu."
            Next a
        End If
    Next c
    SelectedInput = True
    Exit Function
Failed:
    ShowMsg Err.Description, "Kiem tra vung chon", 48
End Function

Public Sub WriteText(ByVal cell As Range, ByVal value As String)
    If Len(value) = 0 Then
        cell.ClearContents
        Exit Sub
    End If
    cell.NumberFormat = "@"
    cell.Value2 = value
End Sub

Public Function EncodeURL(ByVal value As String) As String
    EncodeURL = Application.WorksheetFunction.EncodeURL(value)
End Function

Public Function HttpGet(ByVal url As String) As Object
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 5000, 5000, 10000, 15000
    http.Open "GET", url, False
    http.Send
    If http.Status <> 200 Then Err.Raise vbObjectError + 101, , "Dich vu tra ve HTTP " & http.Status
    Set HttpGet = http
End Function

Public Function DecodeUTF8(ByVal bytes As Variant) As String
    Dim stream As Object, errorNumber As Long, description As String
    On Error GoTo Failed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    DecodeUTF8 = stream.ReadText(-1)
    stream.Close
    If Len(DecodeUTF8) > 0 Then
        If AscW(Left$(DecodeUTF8, 1)) = -257 Then DecodeUTF8 = Mid$(DecodeUTF8, 2)
    End If
    Exit Function
Failed:
    errorNumber = Err.Number: description = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise errorNumber, "DecodeUTF8", description
End Function

Private Function PNGDimension(ByVal bytes As Variant, ByVal offset As Long) As Double
    Dim i As Long, first As Long
    first = LBound(bytes) + offset
    For i = first To first + 3
        PNGDimension = PNGDimension * 256# + CDbl(bytes(i))
    Next i
End Function

Public Function InsertCodeImage(ByVal target As Range, ByVal url As String, ByVal height As Double, ByVal kind As String) As Boolean
    Dim http As Object, stream As Object, fso As Object, file As String
    Dim picture As Shape, old As Shape, i As Long, description As String, shapes As Object
    On Error GoTo Failed
    Set http = HttpGet(url)
    If InStr(1, http.GetResponseHeader("Content-Type"), "image/", vbTextCompare) <> 1 Then Err.Raise 5, , "Dich vu khong tra ve anh."
    Set fso = CreateObject("Scripting.FileSystemObject")
    file = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName & ".png")
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile file, 2
    stream.Close
    Set shapes = target.Parent.Shapes
    If kind = "code128" Or kind = "code39" Or kind = "ean13" Then
        ' Picture fills retain original pixels when Excel saves the workbook.
        ' AddPicture2(Compress:=False) still gets recompressed during SaveAs.
        Set picture = shapes.AddShape(msoShapeRectangle, target.Left + 2, target.Top + 2, PNGDimension(http.ResponseBody, 16) / 4#, PNGDimension(http.ResponseBody, 20) / 4#)
        picture.Fill.UserPicture file
        picture.Line.Visible = msoFalse
    Else
        On Error Resume Next
        Set picture = shapes.AddPicture2(file, msoFalse, msoTrue, target.Left + 2, target.Top + 2, -1, -1, 0)
        On Error GoTo Failed
        If picture Is Nothing Then Set picture = shapes.AddPicture(file, msoFalse, msoTrue, target.Left + 2, target.Top + 2, -1, -1)
    End If
    picture.LockAspectRatio = msoTrue
    If kind = "code128" Or kind = "code39" Or kind = "ean13" Then
        ' 6 source pixels per narrow module; 4 pixels per point = 288 dpi.
        ' Keep narrow modules at 1.5 points, never squeeze a long barcode.
        picture.Width = PNGDimension(http.ResponseBody, 16) / 4#
    Else
        picture.Height = height
        If picture.Width > 500 Then picture.Width = 500
    End If
    picture.Placement = xlMove
    picture.AlternativeText = IMAGE_TAG & kind
    For i = target.Parent.Shapes.Count To 1 Step -1
        Set old = target.Parent.Shapes(i)
        If old.Name <> picture.Name And Left$(old.AlternativeText, Len(IMAGE_TAG)) = IMAGE_TAG Then
            If Not Intersect(old.TopLeftCell, target) Is Nothing Then old.Delete
        End If
    Next i
    If target.RowHeight < picture.Height + 4 Then target.RowHeight = picture.Height + 4
    If target.Width < picture.Width + 4 Then target.EntireColumn.ColumnWidth = Application.Min(255, target.EntireColumn.ColumnWidth * (picture.Width + 4) / target.Width)
    fso.DeleteFile file
    InsertCodeImage = True
    Exit Function
Failed:
    description = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Len(file) > 0 Then fso.DeleteFile file
    On Error GoTo 0
    Err.Raise vbObjectError + 102, , description
End Function
'''
put('mod_Core',common)

text=original('mod_XuLyVanBan')
lower='aàáảãạăằắẳẵặâầấẩẫậeèéẻẽẹêềếểễệiìíỉĩịoòóỏõọôồốổỗộơờớởỡợuùúủũụưừứửữựyỳýỷỹỵdđ'
upper=lower.upper()
assert len(lower)==len(upper) and upper[lower.index('u')]=='U'
text=re.sub(r'Private Const UNI_LOWER As String = .*',lambda m:'Private Const UNI_LOWER As String = "'+lower+'"',text)
text=re.sub(r'Private Const UNI_UPPER As String = .*',lambda m:'Private Const UNI_UPPER As String = "'+upper+'"',text)
text=text.replace('    Dim sLower As String, sUpper As String, res As String', '    Static sLower As String, sUpper As String\n    Dim res As String')
text=text.replace('    sLower = UVanBan(UNI_LOWER)\n    sUpper = UVanBan(UNI_UPPER)', '    If Len(sLower) = 0 Then\n        sLower = UVanBan(UNI_LOWER)\n        sUpper = UVanBan(UNI_UPPER)\n    End If')
start=text.index('    Dim s As String',text.index('Public Function BoDau'))
end=text.index('\nEnd Function', start)
groups={}
for ch in lower+upper:
    base=unicodedata.normalize('NFD',ch)[0]
    if ch=='đ': base='d'
    if ch=='Đ': base='D'
    if ch!=base: groups[ch]=base
body='    Dim s As String, i As Long, ch As String, code As Long\n    If IsError(txt) Or IsNull(txt) Then Exit Function\n    s = CStr(txt)\n'
for ch, base in groups.items(): body+=f'    s = Replace(s, ChrW({ord(ch)}), "{base}")\n'
body+='    For i = 1 To Len(s)\n        ch = Mid$(s, i, 1): code = AscW(ch)\n        If code < 768 Or code > 879 Then BoDau = BoDau & ch\n    Next i'
text=text[:start]+body+text[end:]
text=text.replace('    s = CStr(txt)','    If IsError(txt) Or IsNull(txt) Then Exit Function\n    s = CStr(txt)')
text=text.replace('    s = Replace(s, ChrW(160), " ")','    s = Replace(s, ChrW(160), " ")\n    s = Replace(s, ChrW(8203), "")\n    s = Replace(s, ChrW(-257), "")')
start=text.index('Public Function ToSentenceCase')
end=text.index('\nEnd Function',start)+len('\nEnd Function')
text=text[:start]+'''Public Function ToSentenceCase(ByVal txt As String) As String
    Dim s As String, i As Long, ch As String, newSentence As Boolean
    s = ToLowerUnicode(LamSachText(txt)): newSentence = True
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If newSentence And ToUpperUnicode(ch) <> ToLowerUnicode(ch) Then
            ch = ToUpperUnicode(ch): newSentence = False
        End If
        ToSentenceCase = ToSentenceCase & ch
        If InStr(1, ".!?", ch, vbBinaryCompare) > 0 Then newSentence = True
    Next i
End Function'''+text[end:]
text=text[:text.index('Public Sub Ribbon_TachHoTen')]
for callback, mode in [('TachHoTen','split'),('XoaDauTV','accent'),('InHoa','upper'),('InThuong','lower'),('SentenceCase','sentence'),('DauMoiTu','title'),('CleanText','clean')]:
    text+=f'Public Sub Ribbon_{callback}(control As IRibbonControl)\n    RunTextAction "{mode}"\nEnd Sub\n\n'
text+='''Private Sub RunTextAction(ByVal mode As String)
    Dim inputRange As Range, cell As Range, result As String, oldScreen As Boolean, oldEvents As Boolean
    Dim columns As Long, reason As String
    If mode = "split" Then columns = 2
    If Not SelectedInput(inputRange, columns, mode = "split") Then Exit Sub
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False: Application.EnableEvents = False
    For Each cell In inputRange.Cells
        If Not IsError(cell.Value2) Then
            If VarType(cell.Value2) = vbString And Not cell.HasFormula Or mode = "split" Then
                Select Case mode
                    Case "split"
                        WriteText cell.Offset(0, 1), TachHoDem(cell.Value2)
                        WriteText cell.Offset(0, 2), TachTen(cell.Value2)
                    Case "accent": WriteText cell, BoDau(cell.Value2)
                    Case "upper": WriteText cell, ToUpperUnicode(cell.Value2)
                    Case "lower": WriteText cell, ToLowerUnicode(cell.Value2)
                    Case "sentence": WriteText cell, ToSentenceCase(cell.Value2)
                    Case "title": WriteText cell, ToTitleCase(cell.Value2)
                    Case "clean": WriteText cell, LamSachText(cell.Value2)
                End Select
            End If
        End If
    Next cell
Done:
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    If reason <> "" Then ShowMsg reason, "Loi xu ly", 48
    Exit Sub
Failed:
    reason = Err.Description: Resume Done
End Sub
'''
put('mod_XuLyVanBan',text)

text=original('mod_DoiSo')
text=text.replace('Optional ByVal donVi As String = " \\u0111\\u1ed3ng."','Optional ByVal donVi As String = "{setting}"')
text=text.replace('Optional ByVal unitName As String = " Dollars."','Optional ByVal unitName As String = "{setting}"')
text=text.replace('    DocSoEN = strRes & unitName', '    strRes = LCase$(strRes)\n    DocSoEN = UCase$(Left$(strRes, 1)) & Mid$(strRes, 2) & unitName')
text=text.replace('    Dim num As Double, isNegative As Boolean\n    If', '    Dim num As Double, isNegative As Boolean\n    If')
text=text.replace('    If IsEmpty(targetNum) Or Not IsNumeric(targetNum) Then DocSo = "": Exit Function','    If donVi = "{setting}" Then donVi = NumberUnit(False)\n    If IsEmpty(targetNum) Or Not IsNumeric(targetNum) Then DocSo = "": Exit Function')
text=text.replace('    If IsEmpty(targetNum) Or Not IsNumeric(targetNum) Then DocSoEN = "": Exit Function','    If unitName = "{setting}" Then unitName = NumberUnit(True)\n    If IsEmpty(targetNum) Or Not IsNumeric(targetNum) Then DocSoEN = "": Exit Function')
text=text.replace(') As String\n    Dim num As Double, isNegative As Boolean',') As Variant\n    Dim num As Double, isNegative As Boolean')
text=text.replace('    If IsEmpty(targetNum)', '    If IsError(targetNum) Then\n        '+ 'RESULT' +' = targetNum: Exit Function\n    End If\n    If IsNull(targetNum) Then Exit Function\n    If IsEmpty(targetNum)')
# Function-specific placeholders and strict limits: integers only, up to Excel 15 digits.
for fn in ['DocSo','DocSoEN']:
    start=text.index('Public Function '+fn+'('); end=text.index('\nEnd Function',start)
    part=text[start:end].replace('RESULT',fn)
    part=part.replace('    num = CDbl(targetNum)','    num = CDbl(targetNum)\n    If Abs(num) >= 1E+15 Or num <> Fix(num) Then\n        '+fn+' = CVErr(xlErrNum): Exit Function\n    End If')
    text=text[:start]+part+text[end:]
text=text[:text.index('Public Sub Ribbon_NumToWordsVN')]+'''Public Sub Ribbon_NumToWordsVN(control As IRibbonControl)
    RunNumberAction False
End Sub
Public Sub Ribbon_NumToWordsEN(control As IRibbonControl)
    RunNumberAction True
End Sub
Private Sub RunNumberAction(ByVal english As Boolean)
    Dim inputRange As Range, cell As Range, result As Variant, oldScreen As Boolean, oldEvents As Boolean, reason As String
    If Not SelectedInput(inputRange, 1) Then Exit Sub
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False: Application.EnableEvents = False
    For Each cell In inputRange.Cells
        If english Then result = DocSoEN(cell.Value2) Else result = DocSo(cell.Value2)
        If IsError(result) Then
            cell.Offset(0, 1).Value = result
        Else
            WriteText cell.Offset(0, 1), CStr(result)
        End If
    Next cell
Done:
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    If reason <> "" Then ShowMsg reason, "Loi doc so", 48
    Exit Sub
Failed:
    reason = Err.Description: Resume Done
End Sub
'''
put('mod_DoiSo',text)
put('mod_Settings',r'''
Option Explicit
Private ribbonUI As IRibbonUI
Public Sub Ribbon_Load(ByVal ribbon As IRibbonUI)
    Set ribbonUI = ribbon
End Sub
Public Function NumberUnit(Optional ByVal english As Boolean = False) As String
    NumberUnit = NumberUnitFor(GetSetting("TienIchExcel", "NumberWords", "Unit", "dong"), english)
End Function
Public Function NumberUnitFor(ByVal mode As String, Optional ByVal english As Boolean = False) As String
    Select Case mode
        Case "number": NumberUnitFor = ""
        Case "usd"
            If english Then NumberUnitFor = " US dollars." Else NumberUnitFor = u(" đô la Mỹ.")
        Case Else
            If english Then NumberUnitFor = " dong." Else NumberUnitFor = u(" đồng.")
    End Select
End Function
Public Sub Ribbon_UnitPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (GetSetting("TienIchExcel", "NumberWords", "Unit", "dong") = control.Tag)
End Sub
Public Sub Ribbon_SetUnit(control As IRibbonControl, ByVal pressed As Boolean)
    SaveSetting "TienIchExcel", "NumberWords", "Unit", control.Tag
    If Not ribbonUI Is Nothing Then ribbonUI.Invalidate
    Application.CalculateFull
End Sub
''')

text='Option Explicit\n'
for callback,kind in [('Code128','code128'),('Code39','code39'),('EAN13','ean13'),('QRCode','qr')]:
    text+=f'Public Sub Ribbon_{callback}(control As IRibbonControl)\n    RunCodes "{kind}"\nEnd Sub\n'
text+=r'''
Public Function ValidEAN13(ByVal value As String) As Boolean
    Dim i As Long, total As Long
    If Not DigitsOnly(value) Then Exit Function
    If Len(value) <> 12 And Len(value) <> 13 Then Exit Function
    For i = 1 To 12
        total = total + CLng(Mid$(value, i, 1)) * IIf(i Mod 2 = 0, 3, 1)
    Next i
    If Len(value) = 13 Then
        If CLng(Right$(value, 1)) <> (10 - total Mod 10) Mod 10 Then Exit Function
    End If
    ValidEAN13 = True
End Function

Public Function CodeURL(ByVal kind As String, ByVal value As String) As String
    Dim i As Long, n As Long
    If Len(value) = 0 Then Err.Raise 5, , "Du lieu rong."
    Select Case kind
        Case "ean13"
            If Not ValidEAN13(value) Then Err.Raise 5, , "EAN-13 can 12 chu so hoac 13 chu so dung checksum."
        Case "code39"
            value = UCase$(value)
            For i = 1 To Len(value)
                If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%", Mid$(value, i, 1), vbBinaryCompare) = 0 Then Err.Raise 5, , "Code39 co ky tu khong ho tro."
            Next i
        Case "code128"
            For i = 1 To Len(value)
                n = AscW(Mid$(value, i, 1))
                If n < 32 Or n > 126 Then Err.Raise 5, , "Code128 chi nhan ASCII 32-126. Hay dung QR cho tieng Viet."
            Next i
        Case "qr"
            CodeURL = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&margin=12&charset-source=UTF-8&data=" & EncodeURL(value)
            Exit Function
        Case Else: Err.Raise 5, , "Loai ma khong ho tro."
    End Select
    CodeURL = "https://bwipjs-api.metafloor.com/?bcid=" & kind & "&text=" & EncodeURL(value) & "&scale=6&height=18&includetext&textsize=12&monochrome&backgroundcolor=FFFFFF&paddingwidth=12&paddingheight=6"
End Function

Private Sub RunCodes(ByVal kind As String)
    Dim inputRange As Range, cell As Range, value As String, good As Long, bad As Long
    Dim reason As String, oldScreen As Boolean, oldEvents As Boolean
    If Not SelectedInput(inputRange, 1) Then Exit Sub
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False: Application.EnableEvents = False
    For Each cell In inputRange.Cells
        If IsError(cell.Value2) Then
            bad = bad + 1
        Else
            value = CStr(cell.Value2)
            If Len(value) > 0 Then
                On Error GoTo ItemFailed
                InsertCodeImage cell.Offset(0, 1), CodeURL(kind, value), IIf(kind = "qr", 120, 55), kind
                good = good + 1
            End If
        End If
NextCell:
    Next cell
Done:
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    ShowMsg "Da tao: " & good & ". Loi: " & bad & IIf(reason = "", "", vbCrLf & reason), "Tao ma"
    Exit Sub
ItemFailed:
    bad = bad + 1: reason = cell.Address(False, False) & ": " & Err.Description
    Resume NextCell
Failed:
    reason = Err.Description: Resume Done
End Sub

Public Sub Ribbon_Clear_Barcodes(control As IRibbonControl)
    Dim count As Long
    On Error GoTo Failed
    If TypeName(Selection) <> "Range" Then Exit Sub
    count = ClearCodeImages(Selection)
    ShowMsg "Da xoa " & count & " anh ma do phien ban nay tao.", "Xoa ma"
    Exit Sub
Failed:
    ShowMsg Err.Description, "Loi xoa ma", 48
End Sub

Public Function ClearCodeImages(ByVal selected As Range) As Long
    Dim item As Shape, i As Long, count As Long
    If selected.Parent.ProtectDrawingObjects Then Err.Raise 5, , "Trang tinh dang bao ve hinh anh."
    For i = selected.Parent.Shapes.Count To 1 Step -1
        Set item = selected.Parent.Shapes(i)
        If Left$(item.AlternativeText, Len(IMAGE_TAG)) = IMAGE_TAG Then
            If Not Intersect(item.TopLeftCell, selected) Is Nothing Then
                item.Delete: count = count + 1
            End If
        End If
    Next i
    ClearCodeImages = count
End Function
'''
put('mod_MaVachQR',text)

put('mod_VietQR',r'''
Option Explicit
Public Sub Ribbon_VietQR(control As IRibbonControl)
    ' Open the form first; validate its target only when creating the image.
    On Error GoTo Failed
    If TypeName(ActiveSheet) <> "Worksheet" Then Err.Raise 5, , "Hay mo mot trang tinh de tao VietQR."
    frmVietQR.Show
    Exit Sub
Failed:
    ShowMsg Err.Description, "VietQR", 48
End Sub
Public Function VietQRURL(ByVal bank As String, ByVal account As String, ByVal amount As String, ByVal memo As String) As String
    Dim i As Long, ch As String
    If Len(bank) = 0 Or bank Like "*[!A-Za-z0-9]*" Then Err.Raise 5, , "Ma ngan hang khong hop le."
    If Len(account) < 1 Or Len(account) > 19 Or account Like "*[!A-Za-z0-9]*" Then Err.Raise 5, , "Tai khoan can 1-19 ky tu chu/so."
    If amount <> "" Then
        If Not DigitsOnly(amount) Then Err.Raise 5, , "So tien phai la so nguyen VND, khong co dau phan cach."
        If Len(amount) > 13 Then Err.Raise 5, , "So tien qua lon."
        If CDbl(amount) <= 0 Then Err.Raise 5, , "So tien phai la so duong."
    End If
    If Len(memo) > 50 Then Err.Raise 5, , "Noi dung toi da 50 ky tu."
    memo = BoDau(memo)
    If memo Like "*[!A-Za-z0-9 ]*" Then Err.Raise 5, , "Noi dung chi gom chu, so va khoang trang."
    Dim template As String
    If amount = "" Then template = "compact" Else template = "compact2"
    VietQRURL = "https://img.vietqr.io/image/" & bank & "-" & account & "-" & template & ".png"
    If amount <> "" Then VietQRURL = VietQRURL & "?amount=" & amount
    If memo <> "" Then
        If amount <> "" Then VietQRURL = VietQRURL & "&" Else VietQRURL = VietQRURL & "?"
        VietQRURL = VietQRURL & "addInfo=" & EncodeURL(memo)
    End If
End Function
Public Sub Ribbon_OpenBankList(control As IRibbonControl)
    On Error GoTo Failed
    ThisWorkbook.FollowHyperlink "https://www.vietqr.io/api-bank"
    Exit Sub
Failed:
    ShowMsg Err.Description, "Tra ngan hang", 48
End Sub
''')
text=original('frmVietQR')
text=text.replace('Option Explicit','''Option Explicit
Private WithEvents saveButton As MSForms.CommandButton
Private targetCell As Range
Private statusLabel As MSForms.Label
Private firstActivation As Boolean''',1)
text=text.replace('    Dim savedBankIndex As String, savedAcc As String','''    Dim savedBankIndex As String, savedAcc As String
    Set targetCell = ActiveCell
    Me.Width = 500
    Me.Height = 248
    Me.StartUpPosition = 1
    ' UserForm title bars use the Windows ANSI code page; brand text is portable.
    Me.Caption = "VietQR"
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 10
    Set saveButton = Me.Controls.Add("Forms.CommandButton.1", "btnSaveAccount", True)
    With saveButton
        .Caption = u("Lưu tài khoản")
        .Left = 144: .Top = 162: .Width = 104: .Height = 28
        .ControlTipText = u("Ghi nhớ ngân hàng và số tài khoản trên máy này, không cần tạo QR trước.")
        .TabIndex = 4
    End With
    Set statusLabel = Me.Controls.Add("Forms.Label.1", "lblSaveStatus", True)
    With statusLabel
        .Left = 18: .Top = 126: .Width = 450: .Height = 28
        .WordWrap = True
        .Caption = u("Ngân hàng và tài khoản được nhớ cho lần sau. Bạn cũng có thể bấm Lưu tài khoản trước khi tạo QR.")
    End With
    With btnCreate
        .Left = 260: .Top = 162: .Width = 104: .Height = 28
        .TabIndex = 5: .Default = True
    End With
    With btnClose
        .Left = 376: .Top = 162: .Width = 80: .Height = 28
        .TabIndex = 6: .Cancel = True
    End With
    cboBank.Style = fmStyleDropDownList
    cboBank.MatchEntry = fmMatchEntryComplete
    cboBank.ListRows = 14
    cboBank.TabIndex = 0
    txtAcc.TabIndex = 1
    txtAmount.TabIndex = 2
    txtMemo.TabIndex = 3
    txtAcc.MaxLength = 19
    txtMemo.MaxLength = 50''')
text=text.replace('    Me.Caption = "Tao ma VietQR - Chuyen khoan ngan hang"','')
snapshot=json.loads((ROOT/'tests/banks-snapshot.json').read_text(encoding='utf-8-sig'))
banks=[b for b in snapshot['data'] if b.get('transferSupported')==1]
start=text.index('    With cboBank'); end=text.index('    txtAcc.Text = savedAcc',start)
bankcode='    With cboBank\n        .Clear\n        .ColumnCount = 2\n        .ColumnWidths = "280 pt;0 pt"\n'
for b in banks:
    bankcode+='        .AddItem "'+b['code']+' - '+b['shortName'].replace('"','""')+'"\n'
    bankcode+='        .List(.ListCount - 1, 1) = "'+b['bin']+'"\n'
bankcode+='''    End With
    savedBankIndex = GetSetting("ExcelAddinBarcode", "VietQRConfig", "DefaultBankBin", "")
    If savedBankIndex = "" Then savedBankIndex = GetSetting("AddInTienIchExcel", "VietQR", "LastBank", "MB")
    savedAcc = GetSetting("ExcelAddinBarcode", "VietQRConfig", "DefaultAccNo", "")
    If savedAcc = "" Then savedAcc = GetSetting("AddInTienIchExcel", "VietQR", "LastAccount", "")
    Dim bankIndex As Long
    cboBank.ListIndex = 0
    For bankIndex = 0 To cboBank.ListCount - 1
        If cboBank.List(bankIndex, 1) = savedBankIndex Or Split(cboBank.List(bankIndex, 0), " - ")(0) = savedBankIndex Then cboBank.ListIndex = bankIndex
    Next bankIndex
'''
text=text[:start]+bankcode+text[end:]
text=text.replace('If IsNumeric(Selection.Value) And Selection.Value > 0 Then','If Not IsError(ActiveCell.Value2) Then\n        If IsNumeric(ActiveCell.Value2) Then\n        If ActiveCell.Value2 > 0 Then')
text=text.replace('txtAmount.Text = CStr(Selection.Value)','txtAmount.Text = Format$(ActiveCell.Value2, "0")')
text=text.replace("    ' Tu dong dua con tro", "        End If\n        End If\n    ' Tu dong dua con tro")
text=text.replace('        txtAmount.SetFocus','        txtAmount.TabIndex = 2')
start=text.index('Private Sub btnCreate_Click()'); end=text.index('Private Sub btnClose_Click()',start)
text=text[:start]+r'''Private Sub btnCreate_Click()
    Dim bank As String, account As String, amount As String, memo As String
    Dim selected As Range, url As String
    On Error GoTo Failed
    If cboBank.ListIndex < 0 Then Err.Raise 5, , "Hay chon ngan hang."
    If targetCell Is Nothing Then Err.Raise 5, , "Hay chon o dat ma QR."
    targetCell.Parent.Activate
    targetCell.Select
    If Not SelectedInput(selected, 1) Then Exit Sub
    bank = cboBank.List(cboBank.ListIndex, 1)
    account = Trim$(txtAcc.Text): amount = Trim$(txtAmount.Text): memo = Trim$(txtMemo.Text)
    url = VietQRURL(bank, account, amount, memo)
    btnCreate.Enabled = False
    InsertCodeImage selected.Offset(0, 1), url, 120, "vietqr"
    If Not SaveDefaultAccount(False) Then Err.Raise 5, , "Da tao QR, nhung khong luu duoc tai khoan tren may."
    Unload Me
    Exit Sub
Failed:
    btnCreate.Enabled = True
    ShowMsg Err.Description, "Khong tao duoc VietQR", 48
End Sub

'''+text[end:]
text+=r'''

Private Sub UserForm_Activate()
    If firstActivation Then Exit Sub
    firstActivation = True
    If Trim$(txtAcc.Text) = "" Then
        cboBank.SetFocus
        cboBank.DropDown
    Else
        txtAmount.SetFocus
    End If
End Sub

Private Sub saveButton_Click()
    Dim saved As Boolean
    saved = SaveDefaultAccount(True)
End Sub

Public Function SaveDefaultAccount(Optional ByVal notify As Boolean = True) As Boolean
    Dim bank As String, account As String, checkedURL As String
    On Error GoTo Failed
    If cboBank.ListIndex < 0 Then Err.Raise 5, , "Hay chon ngan hang."
    bank = cboBank.List(cboBank.ListIndex, 1)
    account = Trim$(txtAcc.Text)
    checkedURL = VietQRURL(bank, account, "", "")
    SaveSetting "ExcelAddinBarcode", "VietQRConfig", "DefaultBankBin", bank
    SaveSetting "ExcelAddinBarcode", "VietQRConfig", "DefaultAccNo", account
    txtAcc.Text = account
    statusLabel.Caption = u("Đã lưu ngân hàng và số tài khoản. Lần mở sau sẽ tự điền lại.")
    statusLabel.ForeColor = RGB(0, 112, 70)
    SaveDefaultAccount = True
    Exit Function
Failed:
    statusLabel.Caption = Err.Description
    statusLabel.ForeColor = RGB(176, 0, 32)
End Function
'''
put('frmVietQR',text)

text=original('mod_HanhChinhVN')
start=text.index('Private Function ChuanHoaCCCD'); end=text.index('Private Function GetTenTinhCCCD',start)
text=text[:start]+'''Private Function ChuanHoaCCCD(ByVal cccd As Variant) As String
    If IsError(cccd) Or IsNull(cccd) Or IsEmpty(cccd) Then Exit Function
    ChuanHoaCCCD = Trim$(CStr(cccd))
End Function

'''+text[end:]
text=text.replace('Len(s) <> 12 Or Not IsNumeric(s)','Len(s) <> 12 Or Not DigitsOnly(s)')
start=text.index('Private Function ExtractJSONValue')
text=text[:start]+r'''
Public Function ExtractJSONValue(ByVal json As String, ByVal key As String) As String
    Dim i As Long, token As String, j As Long
    i = 1
    Do While i <= Len(json)
        If Mid$(json, i, 1) = """" Then
            token = ReadJSONString(json, i)
            j = i
            Do While j <= Len(json) And InStr(" " & vbCr & vbLf & vbTab, Mid$(json, j, 1)) > 0
                j = j + 1
            Loop
            If token = key And Mid$(json, j, 1) = ":" Then
                i = j + 1
                Do While i <= Len(json) And InStr(" " & vbCr & vbLf & vbTab, Mid$(json, i, 1)) > 0
                    i = i + 1
                Loop
                If Mid$(json, i, 1) = """" Then
                    ExtractJSONValue = ReadJSONString(json, i)
                Else
                    j = i
                    Do While j <= Len(json)
                        If InStr(",}] " & vbCr & vbLf & vbTab, Mid$(json, j, 1)) > 0 Then Exit Do
                        j = j + 1
                    Loop
                    token = Mid$(json, i, j - i)
                    If token <> "null" Then ExtractJSONValue = token
                End If
                Exit Function
            End If
        Else
            i = i + 1
        End If
    Loop
End Function

Private Function ReadJSONString(ByVal json As String, ByRef i As Long) As String
    Dim ch As String, code As Long
    i = i + 1
    Do While i <= Len(json)
        ch = Mid$(json, i, 1): i = i + 1
        If ch = """" Then Exit Function
        If ch = "\" Then
            ch = Mid$(json, i, 1): i = i + 1
            Select Case ch
                Case "u"
                    code = CLng("&H" & Mid$(json, i, 4))
                    If code > 32767 Then code = code - 65536
                    ch = ChrW(code): i = i + 4
                Case "n": ch = vbLf
                Case "r": ch = vbCr
                Case "t": ch = vbTab
                Case "b": ch = Chr$(8)
                Case "f": ch = Chr$(12)
                Case """", "\", "/"
                Case Else: Err.Raise 5, , "JSON escape khong hop le."
            End Select
        End If
        ReadJSONString = ReadJSONString & ch
    Loop
    Err.Raise 5, , "JSON khong ket thuc chuoi."
End Function

Public Sub Ribbon_CheckCCCD(control As IRibbonControl)
    Dim selected As Range, cell As Range, s As String, oldScreen As Boolean, oldEvents As Boolean, reason As String
    If Not SelectedInput(selected, 3) Then Exit Sub
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False: Application.EnableEvents = False
    For Each cell In selected.Cells
        s = ChuanHoaCCCD(cell.Value2)
        If s <> "" Then
            If KiemTraCCCD(s) Then
                WriteText cell.Offset(0, 1), CCCD_GioiTinh(s)
                cell.Offset(0, 2).Value2 = CCCD_NamSinh(s)
                WriteText cell.Offset(0, 3), CCCD_TinhThanh(s)
            Else
                WriteText cell.Offset(0, 1), "CCCD khong dung cau truc 12 chu so/ma tinh."
            End If
        End If
    Next cell
Done:
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    If reason <> "" Then ShowMsg reason, "Loi CCCD", 48
    Exit Sub
Failed:
    reason = Err.Description: Resume Done
End Sub

Public Sub Ribbon_LookupMST(control As IRibbonControl)
    Dim selected As Range, cell As Range, mst As String
    Dim company As String, address As String, oldScreen As Boolean, oldEvents As Boolean, reason As String, itemReason As String
    Dim usedSources As String
    Dim good As Long, bad As Long
    If Not SelectedInput(selected, 2) Then Exit Sub
    If selected.CountLarge > 100 Then
        ShowMsg u("M\u1ed7i l\u1ea7n tra t\u1ed1i \u0111a 100 MST. Ngu\u1ed3n: Thongtincongty.vn, d\u1ef1 ph\u00f2ng MaSoThue; ch\u1ecbu gi\u1edbi h\u1ea1n truy c\u1eadp c\u1ee7a ngu\u1ed3n."), "Tra MST", 48
        Exit Sub
    End If
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False: Application.EnableEvents = False
    For Each cell In selected.Cells
        company = "": address = "": itemReason = "": mst = ""
        If IsError(cell.Value2) Then
            itemReason = "O nguon dang co loi."
        Else
            mst = Trim$(CStr(cell.Value2))
            If mst <> "" Then
                mst = NormalizeMST(mst)
                If mst = "" Then
                    itemReason = "MST can 10 chu so hoac 10 chu so-3 chu so chi nhanh (luu dang Text)."
                Else
                    If LookupMSTData(mst, company, address, itemReason) Then
                        WriteText cell.Offset(0, 1), company
                        WriteText cell.Offset(0, 2), address
                        If InStr(1, usedSources, MSTLastSource, vbBinaryCompare) = 0 Then
                            If Len(usedSources) > 0 Then usedSources = usedSources & vbCrLf
                            usedSources = usedSources & MSTLastSource
                        End If
                        good = good + 1
                    End If
                End If
            End If
        End If
        If itemReason <> "" Then
            bad = bad + 1
            If bad <= 5 Then reason = reason & vbCrLf & cell.Address(False, False) & " " & mst & ": " & itemReason
        End If
    Next cell
Done:
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    If reason <> "" Then ShowMsg "Da tra: " & good & ". Chua tra duoc: " & bad & ". Cot ket qua cua cac dong loi duoc de trong de thu lai." & reason, "Ket qua tra MST", 48
    Exit Sub
Failed:
    reason = Err.Description: Resume Done
End Sub


Public Function HtmlDecodeText(ByVal text As String) As String
    text = Replace(text, "&quot;", Chr$(34))
    text = Replace(text, "&#039;", "'")
    text = Replace(text, "&#39;", "'")
    text = Replace(text, "&apos;", "'")
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&gt;", ">")
    text = Replace(text, "&nbsp;", " ")
    text = Replace(text, "&amp;", "&")
    HtmlDecodeText = LamSachText(text)
End Function

Public Function LookupMSTNameHint(ByVal mst As String, ByRef company As String, ByRef address As String, ByRef reason As String) As Boolean
    Dim http As Object, json As String, invoyReason As String
    company = "": address = "": reason = ""
    mst = NormalizeMST(mst)
    If mst = "" Then reason = "MST khong dung dinh dang.": Exit Function
    On Error GoTo InvoyFailed
    Set http = HttpGet("https://invoy.io.vn/api/v1/public/lookup-mst?mst=" & EncodeURL(mst))
    json = DecodeUTF8(http.ResponseBody)
    If NormalizeMST(ExtractJSONValue(json, "taxCode")) <> mst Then Err.Raise 5, , "Khong co du lieu dung MST yeu cau."
    company = HtmlDecodeText(ExtractJSONValue(json, "name"))
    address = HtmlDecodeText(ExtractJSONValue(json, "address"))
    If company = "" Then Err.Raise 5, , "Dich vu khong tra ve ten doanh nghiep."
    LookupMSTNameHint = True
    Exit Function
InvoyFailed:
    invoyReason = Err.Description
    Resume TryEsgoo
TryEsgoo:
    On Error GoTo Failed
    company = "": address = ""
    Set http = HttpGet("https://esgoo.net/api-mst/" & mst & ".htm")
    json = DecodeUTF8(http.ResponseBody)
    If ExtractJSONValue(json, "error") <> "0" Then Err.Raise 5, , ExtractJSONValue(json, "error_text")
    If NormalizeMST(ExtractJSONValue(json, "mst")) <> mst Then Err.Raise 5, , "Dich vu tra ve MST khac."
    company = HtmlDecodeText(ExtractJSONValue(json, "ten"))
    address = HtmlDecodeText(ExtractJSONValue(json, "dc"))
    If company = "" Then Err.Raise 5, , "Dich vu khong tra ve ten doanh nghiep."
    LookupMSTNameHint = True
    Exit Function
Failed:
    reason = "Invoy: " & invoyReason & "; Esgoo: " & Err.Description
    company = "": address = ""
End Function

Public Function NormalizeMST(ByVal value As String) As String
    NormalizeMST = ""
    value = Trim$(value)
    If Len(value) = 10 And DigitsOnly(value) Then
        NormalizeMST = value
    ElseIf Len(value) = 13 And DigitsOnly(value) Then
        NormalizeMST = Left$(value, 10) & "-" & Right$(value, 3)
    ElseIf Len(value) = 14 And Mid$(value, 11, 1) = "-" Then
        If DigitsOnly(Left$(value, 10) & Right$(value, 3)) Then NormalizeMST = value
    End If
End Function
'''
put('mod_HanhChinhVN',text)
text=original('mod_ThongTin')
text=text[:text.index('Public Sub Ribbon_About')] + '''Public Sub Ribbon_About(control As IRibbonControl)
    frmAbout.Show vbModeless
End Sub
'''
put('mod_ThongTin',text)
put('frmAbout',(ROOT/'about/frmAbout.vba').read_text(encoding='utf-8'))

import compare_sources
import address_sources
import undo_sources
import translate_sources
import pdf_export_sources
import printing_sources
import currency_sources
import qr_image_fix.sources

if __name__ == '__main__':
    print('Generated VBA sources:', ', '.join(p.stem for p in SRC.glob('*.txt')))
