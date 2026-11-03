Option Explicit
Private results As Object
Private blockedUntil As Date
Private translateBusy As Boolean
Private cancelRequested As Boolean
Public Const MAX_TRANSLATE_CHARS As Long = 5000

Public Function TranslationIsBusy() As Boolean
    TranslationIsBusy = translateBusy
End Function

Public Function TranslationSourceLanguage() As String
    TranslationSourceLanguage = GetSetting("TienIchExcel", "Translate", "Source", "vi")
End Function

Public Function TranslationTargetLanguage() As String
    TranslationTargetLanguage = GetSetting("TienIchExcel", "Translate", "Target", "en")
End Function

Public Sub SaveTranslationSettings(ByVal source As String, ByVal target As String, ByVal email As String)
    If translateBusy Then Err.Raise 5, , u("Đang dịch. Hãy chờ hoặc Dừng trước khi lưu cài đặt.")
    If LanguageCode(source) = "" Or LanguageCode(target) = "" Then Err.Raise 5, , u("Hãy chọn ngôn ngữ nguồn và đích hợp lệ.")
    If Not ValidTranslationEmail(email) Then Err.Raise 5, , u("Email chưa đúng định dạng; nhập email đang dùng hoặc để trống.")
    SaveTranslationEmail email
    SaveSetting "TienIchExcel", "Translate", "Source", Trim$(source)
    SaveSetting "TienIchExcel", "Translate", "Target", Trim$(target)
End Sub

Public Function TranslateWithDefaults(ByVal selected As Range, Optional ByVal progress As Object = Nothing) As String
    TranslateWithDefaults = TranslateSelection(selected, TranslationSourceLanguage(), TranslationTargetLanguage(), progress)
End Function

Public Function TranslationEmail() As String
    TranslationEmail = Trim$(GetSetting("TienIchExcel", "Translate", "Email", ""))
End Function

Public Function ValidTranslationEmail(ByVal email As String) As Boolean
    Dim expression As Object
    email = Trim$(email)
    If email = "" Then ValidTranslationEmail = True: Exit Function
    If Len(email) > 254 Or InStr(email, "..") > 0 Then Exit Function
    Set expression = CreateObject("VBScript.RegExp")
    expression.Pattern = "^[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+([.][A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*@[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?([.][A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?)+$"
    ValidTranslationEmail = expression.Test(email)
End Function

Public Sub SaveTranslationEmail(ByVal email As String)
    email = Trim$(email)
    If Not ValidTranslationEmail(email) Then Err.Raise 5, , u("Email chưa đúng định dạng. Hãy nhập email đang dùng của bạn hoặc để trống.")
    If StrComp(email, TranslationEmail(), vbBinaryCompare) <> 0 Then
        SaveSetting "TienIchExcel", "Translate", "Email", email
        blockedUntil = 0
    End If
End Sub

Public Function TranslationRequestURL(ByVal text As String, ByVal source As String, ByVal target As String, Optional ByVal contact As Variant) As String
    Dim email As String
    If IsMissing(contact) Then email = TranslationEmail() Else email = Trim$(CStr(contact))
    If Not ValidTranslationEmail(email) Then Err.Raise 5, , u("Email MyMemory đã lưu không hợp lệ; hãy sửa trong form Dịch ngôn ngữ.")
    TranslationRequestURL = "https://api.mymemory.translated.net/get?q=" & EncodeURL(text) & "&langpair=" & EncodeURL(source & "|" & target)
    If Len(email) > 0 Then TranslationRequestURL = TranslationRequestURL & "&de=" & EncodeURL(email)
End Function

' Each item is Array(payload, gap). Gaps preserve paragraph/word separators.
Public Function TranslationParts(ByVal text As String) As Collection
    Dim parts As New Collection, p As Long, start As Long, q As Long, units As Long
    Dim bytes As Long, width As Long, code As Long, low As Long, lastSpace As Long, finish As Long
    Dim ch As String
    p = 1
    Do While p <= Len(text)
        start = p
        Do While p <= Len(text)
            If InStr(1, " " & vbTab & vbCr & vbLf, Mid$(text, p, 1), vbBinaryCompare) = 0 Then Exit Do
            p = p + 1
        Loop
        If p > start Then parts.Add Array("", Mid$(text, start, p - start))
        If p > Len(text) Then Exit Do
        start = p: q = p: bytes = 0: lastSpace = 0
        Do While q <= Len(text)
            ch = Mid$(text, q, 1)
            If ch = vbCr Or ch = vbLf Then Exit Do
            code = AscW(ch) And &HFFFF&: units = 1
            Select Case code
                Case 0 To 127: width = 1
                Case 128 To 2047: width = 2
                Case &HD800& To &HDBFF&
                    If q < Len(text) Then
                        low = AscW(Mid$(text, q + 1, 1)) And &HFFFF&
                        If low >= &HDC00& And low <= &HDFFF& Then units = 2
                    End If
                    If units = 2 Then width = 4 Else width = 3
                Case Else: width = 3
            End Select
            If bytes + width > 500 Then Exit Do
            bytes = bytes + width
            If ch = " " Or ch = vbTab Then lastSpace = q
            q = q + units
        Loop
        finish = q - 1
        If q <= Len(text) And bytes >= 497 And lastSpace > start Then finish = lastSpace - 1
        Do While finish >= start
            ch = Mid$(text, finish, 1)
            If ch <> " " And ch <> vbTab Then Exit Do
            finish = finish - 1
        Loop
        If finish < start Then Err.Raise 5, , "Cannot split translation text."
        parts.Add Array(Mid$(text, start, finish - start + 1), "")
        p = finish + 1
    Loop
    Set TranslationParts = parts
End Function

Private Function TranslationResult(ByVal value As Variant, ByVal source As String, ByVal target As String, Optional ByVal interactive As Boolean = False) As Variant
    Dim raw As Variant, text As String, parts As Collection, part As Variant, entry As Variant, joined As String
    On Error GoTo Failed
    If IsObject(value) Then
        If TypeName(value) <> "Range" Then
            TranslationResult = ShortTranslationResult(value, source, target): Exit Function
        End If
        If value.CountLarge <> 1 Then
            TranslationResult = ShortTranslationResult(value, source, target): Exit Function
        End If
        raw = value.Value2
    Else
        raw = value
    End If
    If IsError(raw) Or IsArray(raw) Or IsNull(raw) Then
        TranslationResult = ShortTranslationResult(raw, source, target): Exit Function
    End If
    text = CStr(raw)
    If Len(text) > MAX_TRANSLATE_CHARS Then
        TranslationResult = Array(False, u("Tối đa 5.000 ký tự mỗi ô. Hãy chia thành nhiều ô."), Now + 1): Exit Function
    End If
    If UTF8Length(text) <= 500 Or LCase$(Trim$(source)) = LCase$(Trim$(target)) Then
        TranslationResult = ShortTranslationResult(text, source, target): Exit Function
    End If
    Set parts = TranslationParts(text)
    For Each part In parts
        If interactive Then
            DoEvents
            If cancelRequested Then
                TranslationResult = Array(False, u("Đã dừng. Chưa ghi kết quả vào trang tính."), Now + 1): Exit Function
            End If
        End If
        If Len(part(0)) > 0 Then
            entry = ShortTranslationResult(part(0), source, target)
            If Not entry(0) Then TranslationResult = entry: Exit Function
            joined = joined & entry(1)
        End If
        joined = joined & part(1)
        If Len(joined) > 32767 Then
            TranslationResult = Array(False, u("Bản dịch vượt giới hạn 32.767 ký tự của ô Excel."), Now + 1): Exit Function
        End If
    Next part
    TranslationResult = Array(True, joined, Now + 1)
    Exit Function
Failed:
    TranslationResult = Array(False, u("Không xử lý được văn bản: ") & Err.Description, Now + 1)
End Function

Public Sub CancelDirectTranslation()
    cancelRequested = True
End Sub

Public Sub ValidateTranslationRange(ByVal selected As Range)
    Dim cell As Range, dest As Range, total As Long
    If selected Is Nothing Then Err.Raise 5, , u("Hãy chọn các ô văn bản cần dịch.")
    If selected.CountLarge > 100 Then Err.Raise 5, , u("Mỗi lần chọn tối đa 100 ô trong một cột.")
    If selected.Parent.ProtectContents Then Err.Raise 5, , u("Trang tính đang được bảo vệ.")
    If selected.Parent.Parent.ReadOnly Then Err.Raise 5, , u("Tệp đang mở chỉ đọc.")
    For Each dest In selected.Areas
        If dest.Columns.Count <> 1 Then Err.Raise 5, , u("Chỉ chọn một cột văn bản nguồn.")
    Next dest
    For Each cell In selected.Cells
        If cell.Column <> selected.Cells(1, 1).Column Then Err.Raise 5, , u("Các ô nguồn phải nằm trong cùng một cột.")
        If cell.MergeCells Then Err.Raise 5, , u("Không hỗ trợ ô gộp.")
        If cell.Column = cell.Parent.Columns.Count Then Err.Raise 5, , u("Không còn cột bên phải để ghi bản dịch.")
        Set dest = cell.Offset(0, 1)
        If Not Intersect(dest, selected) Is Nothing Then Err.Raise 5, , u("Cột kết quả trùng vùng nguồn.")
        If dest.MergeCells Or dest.HasFormula Then Err.Raise 5, , u("Cột kết quả bên phải phải trống, không gộp và không có công thức.")
        If IsError(dest.Value2) Then Err.Raise 5, , u("Cột kết quả đang có dữ liệu lỗi.")
        If Len(CStr(dest.Value2)) > 0 Then Err.Raise 5, , u("Cột kết quả bên phải đang có dữ liệu. Không ghi đè.")
        If IsError(cell.Value2) Then Err.Raise 5, , u("Ô nguồn đang chứa lỗi Excel: ") & cell.Address(False, False)
        If VarType(cell.Value2) = vbString Then total = total + Len(CStr(cell.Value2))
    Next cell
    If total > MAX_TRANSLATE_CHARS Then Err.Raise 5, , u("Tổng văn bản mỗi lần dịch tối đa 5.000 ký tự.")
End Sub

Public Function TranslateSelection(ByVal selected As Range, ByVal source As String, ByVal target As String, Optional ByVal progress As Object = Nothing) As String
    Dim staged As New Collection, record As Object, cell As Range, entry As Variant
    Dim count As Long, events As Boolean, screen As Boolean, writing As Boolean, reason As String, ownsRun As Boolean, stateChanged As Boolean
    On Error GoTo Failed
    If translateBusy Then Err.Raise 5, , u("Đang có một lượt dịch. Hãy chờ hoặc bấm Dừng.")
    ValidateTranslationRange selected
    If LanguageCode(source) = "" Or LanguageCode(target) = "" Then Err.Raise 5, , u("Hãy chọn ngôn ngữ nguồn và đích.")
    translateBusy = True: ownsRun = True: cancelRequested = False
    ' Read source snapshots before allowing UI events during network requests.
    For Each cell In selected.Cells
        Set record = CreateObject("Scripting.Dictionary")
        record.Add "cell", cell
        record.Add "value", cell.Value2
        record.Add "formula", cell.Formula
        record.Add "output", ""
        staged.Add record
    Next cell
    For Each record In staged
        Set cell = record("cell")
        If Not progress Is Nothing Then progress.SetProgress u("Đang dịch ô ") & cell.Address(False, False) & u(". Có thể bấm Dừng; kết quả chưa được ghi.")
        DoEvents
        If cancelRequested Then Err.Raise 18, , u("Đã dừng. Chưa ghi kết quả vào trang tính.")
        If VarType(record("value")) = vbString And Len(Trim$(CStr(record("value")))) > 0 Then
            entry = TranslationResult(record("value"), source, target, True)
            If Not entry(0) Then Err.Raise 5, , cell.Address(False, False) & ": " & entry(1)
            record("output") = entry(1)
            count = count + 1
        End If
    Next record
    If count = 0 Then Err.Raise 5, , u("Không có ô văn bản để dịch. Ô trống và ô số được bỏ qua.")
    DoEvents
    If cancelRequested Then Err.Raise 18, , u("Đã dừng. Chưa ghi kết quả vào trang tính.")
    ' Recheck after DoEvents: never overwrite edits made while translation ran.
    ValidateTranslationRange selected
    For Each record In staged
        Set cell = record("cell")
        If VarType(cell.Value2) <> VarType(record("value")) Then Err.Raise 5, , u("Dữ liệu nguồn đã thay đổi. Chưa ghi kết quả; hãy chọn lại vùng dịch.")
        If CStr(cell.Value2) <> CStr(record("value")) Or CStr(cell.Formula) <> CStr(record("formula")) Then Err.Raise 5, , u("Dữ liệu nguồn đã thay đổi. Chưa ghi kết quả; hãy chọn lại vùng dịch.")
    Next record
    events = Application.EnableEvents: screen = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    stateChanged = True
    UndoBegin selected, 1, u("Dịch ngôn ngữ")
    writing = True
    For Each record In staged
        If Len(record("output")) > 0 Then WriteText record("cell").Offset(0, 1), record("output")
    Next record
    UndoCommit
    writing = False
    Application.EnableEvents = events: Application.ScreenUpdating = screen
    TranslateSelection = "OK:" & count
    translateBusy = False
    Exit Function
Failed:
    reason = Err.Description
    If writing Then
        On Error Resume Next
        UndoCommit
        reason = reason & u(" Nếu có kết quả đã ghi, dùng nút Hoàn tác để khôi phục.")
        On Error GoTo 0
    End If
    If stateChanged Then
        Application.EnableEvents = events: Application.ScreenUpdating = screen
    End If
    If ownsRun Then translateBusy = False
    TranslateSelection = reason
End Function

Public Sub Ribbon_TranslateSelection(control As IRibbonControl)
    Dim selected As Range, result As String
    On Error GoTo Failed
    If translateBusy Then Exit Sub
    If TypeName(Selection) <> "Range" Then Err.Raise 5, , u("Hãy chọn các ô văn bản cần dịch.")
    Set selected = Selection
    ValidateTranslationRange selected
    Load frmTranslateProgress
    frmTranslateProgress.BeginRun
    If Not frmTranslateProgress.Visible Then frmTranslateProgress.Show vbModeless
    result = TranslateWithDefaults(selected, frmTranslateProgress)
    frmTranslateProgress.Finish result
    Exit Sub
Failed:
    ShowMsg Err.Description, u("Dịch ngôn ngữ"), 48
End Sub

Public Sub Ribbon_TranslationSettings(control As IRibbonControl)
    On Error GoTo Failed
    If Not frmTranslate.Visible Then frmTranslate.Show vbModeless
    Exit Sub
Failed:
    ShowMsg Err.Description, u("Cài đặt dịch"), 48
End Sub

Private Function LanguageCode(ByVal code As String) As String
    Dim expression As Object
    code = LCase$(Trim$(code))
    Set expression = CreateObject("VBScript.RegExp")
    expression.Pattern = "^[a-z]{2,3}(-[a-z]{2,4})?$"
    If expression.Test(code) Then LanguageCode = code
End Function

Private Function UTF8Length(ByVal text As String) As Long
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText text
    UTF8Length = stream.Size - 3
    stream.Close
End Function

Private Function BestTranslation(ByVal json As String, ByVal fallback As String, ByVal sourceText As String) As String
    Dim i As Long, start As Long, depth As Long, quoted As Boolean, escaped As Boolean
    Dim ch As String, item As String, candidate As String, bestScore As Double, score As Double
    BestTranslation = fallback
    If sourceText = "" Then Exit Function
    bestScore = Val(ExtractJSONValue(json, "match"))
    i = InStr(1, json, """matches""", vbBinaryCompare)
    If i = 0 Then Exit Function
    i = InStr(i, json, "[")
    If i = 0 Then Exit Function
    For i = i + 1 To Len(json)
        ch = Mid$(json, i, 1)
        If quoted Then
            If escaped Then
                escaped = False
            ElseIf ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                quoted = False
            End If
        Else
            Select Case ch
                Case """": quoted = True
                Case "{"
                    If depth = 0 Then start = i
                    depth = depth + 1
                Case "}"
                    depth = depth - 1
                    If depth = 0 And start > 0 Then
                        item = Mid$(json, start, i - start + 1)
                        score = Val(ExtractJSONValue(item, "match"))
                        candidate = ExtractJSONValue(item, "translation")
                        If score > bestScore And Len(candidate) > 0 Then
                            If StrComp(ExtractJSONValue(item, "segment"), sourceText, vbBinaryCompare) = 0 Then
                                bestScore = score: BestTranslation = candidate
                            End If
                        End If
                    End If
                Case "]"
                    If depth = 0 Then Exit Function
            End Select
        End If
    Next i
End Function

Public Function ParseTranslation(ByVal json As String, ByRef translated As String, ByRef reason As String, Optional ByVal sourceText As String = "") As Boolean
    Dim status As String
    translated = "": reason = ""
    On Error GoTo Failed
    status = ExtractJSONValue(json, "responseStatus")
    If ExtractJSONValue(json, "quotaFinished") = "true" Or status = "429" Then
        reason = u("Dịch vụ đã hết hạn mức hoặc đang giới hạn lượt gọi. Thử lại sau.")
        Exit Function
    End If
    If status <> "200" Then
        reason = u("Dịch vụ từ chối yêu cầu: ") & Left$(ExtractJSONValue(json, "responseDetails"), 250)
        Exit Function
    End If
    translated = ExtractJSONValue(json, "translatedText")
    If Len(translated) = 0 Then
        reason = u("Dịch vụ không trả về bản dịch.")
        Exit Function
    End If
    translated = BestTranslation(json, translated, sourceText)
    ParseTranslation = True
    Exit Function
Failed:
    reason = u("Không đọc được phản hồi của dịch vụ dịch.")
End Function

' A cache entry contains success, text/reason, expiry. Errors expire after 60s.
Private Function ShortTranslationResult(ByVal value As Variant, ByVal source As String, ByVal target As String) As Variant
    Dim text As String, key As String, entry As Variant, http As Object
    Dim translated As String, reason As String, json As String, raw As Variant
    On Error GoTo Failed
    If IsObject(value) Then
        If TypeName(value) <> "Range" Then
            ShortTranslationResult = Array(False, u("Chọn một ô hoặc một chuỗi văn bản."), Now + 1): Exit Function
        End If
        If value.CountLarge <> 1 Then
            ShortTranslationResult = Array(False, u("Mỗi công thức chỉ nhận một ô nguồn."), Now + 1): Exit Function
        End If
        raw = value.Value2
    Else
        raw = value
    End If
    If IsError(raw) Then
        ShortTranslationResult = Array(False, u("Ô nguồn đang chứa lỗi Excel."), Now + 1): Exit Function
    End If
    If IsArray(raw) Or IsNull(raw) Then
        ShortTranslationResult = Array(False, u("Chọn một ô hoặc một chuỗi văn bản."), Now + 1): Exit Function
    End If
    text = CStr(raw)
    If Len(Trim$(text)) = 0 Then
        ShortTranslationResult = Array(True, "", Now + 1): Exit Function
    End If
    source = LanguageCode(source): target = LanguageCode(target)
    If source = "" Or target = "" Then
        ShortTranslationResult = Array(False, u("Mã ngôn ngữ không hợp lệ. Ví dụ: vi, en, ja, ko, zh-CN, fr, de. Không dùng auto."), Now + 1): Exit Function
    End If
    If source = target Then
        ShortTranslationResult = Array(True, text, Now + 1): Exit Function
    End If
    If UTF8Length(text) > 500 Then
        ShortTranslationResult = Array(False, u("Mỗi lần dịch tối đa 500 byte UTF-8. Hãy chia văn bản dài thành nhiều ô."), Now + 1): Exit Function
    End If
    key = ProviderConfigurationKey() & vbLf & source & "|" & target & "|" & TranslationEmail() & "|" & text
    If results Is Nothing Then Set results = CreateObject("Scripting.Dictionary")
    If results.Exists(key) Then
        entry = results(key)
        If Now < entry(2) Then ShortTranslationResult = entry: Exit Function
        results.Remove key
    End If
    entry = ProviderTranslate(text, source, target)
    If results.Count >= 1000 Then results.RemoveAll
    results.Add key, entry
    ShortTranslationResult = entry
    Exit Function
Failed:
    blockedUntil = DateAdd("s", 60, Now)
    ShortTranslationResult = Array(False, u("Không kết nối được dịch vụ dịch. Kiểm tra Internet rồi chờ 60 giây và tính lại ô. ") & Left$(Err.Description, 200), blockedUntil)
End Function

Public Function Dich(ByVal text As Variant, Optional ByVal nguon As String = "{setting}", Optional ByVal dichDen As String = "{setting}") As Variant
    Dim entry As Variant
    If nguon = "{setting}" Then nguon = TranslationSourceLanguage()
    If dichDen = "{setting}" Then dichDen = TranslationTargetLanguage()
    If TypeName(text) = "Range" Then
        If text.CountLarge = 1 Then
            If IsError(text.Value2) Then Dich = text.Value2: Exit Function
        End If
    End If
    If IsError(text) Then Dich = text: Exit Function
    entry = TranslationResult(text, nguon, dichDen)
    If entry(0) Then Dich = entry(1) Else Dich = CVErr(xlErrValue)
End Function

Public Function DichTrangThai(ByVal text As Variant, Optional ByVal nguon As String = "{setting}", Optional ByVal dichDen As String = "{setting}") As String
    Dim entry As Variant
    If nguon = "{setting}" Then nguon = TranslationSourceLanguage()
    If dichDen = "{setting}" Then dichDen = TranslationTargetLanguage()
    entry = TranslationResult(text, nguon, dichDen)
    If entry(0) Then
        If entry(1) = "" Then DichTrangThai = u("Ô nguồn trống") Else DichTrangThai = "OK"
    Else
        DichTrangThai = entry(1)
    End If
End Function

Public Sub Ribbon_TranslateHelp(control As IRibbonControl)
    ShowMsg u("1. Nhập văn bản cần dịch vào ô A2.") & vbCrLf & _
        u("2. Chọn ô khác, ví dụ B2, rồi nhập một công thức bên dưới và nhấn Enter:") & vbCrLf & vbCrLf & _
        u("=Dich(A2) — dùng ngôn ngữ đã lưu trong Cài đặt dịch.") & vbCrLf & _
        u("=Dich(A2,""vi"",""en"") — Việt sang Anh.") & vbCrLf & _
        u("=Dich(A2,""en"",""vi"") — Anh sang Việt.") & vbCrLf & _
        u("=Dich(A2,""vi"",""ja"") — Việt sang Nhật.") & vbCrLf & vbCrLf & _
        u("3. Kéo công thức xuống để dịch các dòng tiếp theo.") & vbCrLf & _
        u("Nếu có lỗi: =DichTrangThai(A2,""vi"",""en"") — dùng cùng mã ngôn ngữ với công thức đang lỗi.") & vbCrLf & vbCrLf & _
        u("Mã: vi = Việt; en = Anh; ja = Nhật; ko = Hàn; zh-CN = Trung giản thể; fr = Pháp; de = Đức; th = Thái.") & vbCrLf & _
        u("Nếu Excel dùng dấu ; để ngăn đối số, thay dấu phẩy bằng ;. Không nhập công thức vào chính ô A2.") & vbCrLf & _
        u("Cần Internet; tối đa 5.000 ký tự/ô. Dịch vụ và email dùng theo Cài đặt dịch. Google web có thể bị giới hạn; cần kiểm tra lại bản dịch."), u("Cách dùng công thức dịch")
End Sub
