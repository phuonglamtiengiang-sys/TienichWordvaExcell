Option Explicit
Public Const FX_BANK_URL As String = "https://www.vietcombank.com.vn/KHCN/Cong-cu-tien-ich/Ty-gia"
Public Const FX_XML_URL As String = "https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx"
Private fxCached As Object
Private fxLastAttempt As Date
#If VBA7 Then
Private Declare PtrSafe Function FXOpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function FXCloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
Private Declare PtrSafe Function FXEmptyClipboard Lib "user32" Alias "EmptyClipboard" () As Long
Private Declare PtrSafe Function FXSetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal format As Long, ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function FXGlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal flags As Long, ByVal bytes As LongPtr) As LongPtr
Private Declare PtrSafe Function FXGlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function FXGlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal memory As LongPtr) As Long
Private Declare PtrSafe Function FXGlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Sub FXCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal target As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
#Else
Private Declare Function FXOpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hwnd As Long) As Long
Private Declare Function FXCloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
Private Declare Function FXEmptyClipboard Lib "user32" Alias "EmptyClipboard" () As Long
Private Declare Function FXSetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal format As Long, ByVal memory As Long) As Long
Private Declare Function FXGlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal flags As Long, ByVal bytes As Long) As Long
Private Declare Function FXGlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal memory As Long) As Long
Private Declare Function FXGlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal memory As Long) As Long
Private Declare Function FXGlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal memory As Long) As Long
Private Declare Sub FXCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal target As Long, ByVal source As Long, ByVal bytes As Long)
#End If

Public Function FXGetRates() As Object
    Set FXGetRates = FXBankTable()
End Function

Private Function FXDisplayRate(ByVal fields As Object, ByVal field As String, ByVal forCopy As Boolean) As String
    Dim text As String, number As Double
    If Not fields.Exists(field) Then FXDisplayRate = "-": Exit Function
    text = Trim$(CStr(fields(field)))
    If text = "" Or text = "-" Then FXDisplayRate = "-": Exit Function
    number = FXNumber(text)
    If forCopy Then
        FXDisplayRate = Replace(Trim$(Str$(number)), ".", Application.DecimalSeparator)
    Else
        FXDisplayRate = text
    End If
End Function

Public Function FXViewerRows(ByVal table As Object) As Variant
    Dim result() As Variant, key As Variant, row As Long, fields As Object
    If table.Count <= 2 Then Err.Raise 5, , u("Bảng tỷ giá không có ngoại tệ.")
    ReDim result(0 To table.Count - 3, 0 To 4)
    For Each key In table.Keys
        If key <> "Updated" And key <> "Retrieved" Then
            Set fields = table(key)
            result(row, 0) = key: result(row, 1) = fields("Name")
            result(row, 2) = FXDisplayRate(fields, "Buy", False)
            result(row, 3) = FXDisplayRate(fields, "Transfer", False)
            result(row, 4) = FXDisplayRate(fields, "Sell", False)
            row = row + 1
        End If
    Next key
    FXViewerRows = result
End Function

Private Function FXCopyCell(ByVal text As String) As String
    text = Replace(Replace(Replace(text, vbTab, " "), vbCr, " "), vbLf, " ")
    If Len(text) > 0 Then
        If InStr("=+@-", Left$(text, 1)) > 0 Then text = "'" & text
    End If
    FXCopyCell = text
End Function

Public Function FXRatesText(ByVal table As Object, Optional ByVal selectedCodes As String = "", Optional ByVal includeSource As Boolean = True) As String
    Dim text As String, key As Variant, fields As Object, count As Long
    If includeSource Then
        text = u("Nguồn") & vbTab & "Vietcombank" & vbTab & FX_BANK_URL & vbCrLf
        text = text & u("Công bố") & vbTab & Format$(table("Updated"), "dd/mm/yyyy hh:nn:ss") & vbCrLf
        text = text & u("Tải về") & vbTab & Format$(table("Retrieved"), "dd/mm/yyyy hh:nn:ss") & vbCrLf
        text = text & u("Đơn vị") & vbTab & u("VND cho 1 đơn vị ngoại tệ; tỷ giá tham khảo") & vbCrLf & vbCrLf
    End If
    text = text & u("Mã ngoại tệ") & vbTab & u("Tên ngoại tệ") & vbTab & u("Mua tiền mặt") & vbTab & u("Mua chuyển khoản") & vbTab & u("Bán") & vbCrLf
    For Each key In table.Keys
        If key <> "Updated" And key <> "Retrieved" Then
            If selectedCodes = "" Or InStr(1, "|" & selectedCodes & "|", "|" & key & "|", vbBinaryCompare) > 0 Then
                Set fields = table(key)
                text = text & FXCopyCell(CStr(key)) & vbTab & FXCopyCell(CStr(fields("Name"))) & vbTab
                text = text & FXDisplayRate(fields, "Buy", True) & vbTab & FXDisplayRate(fields, "Transfer", True) & vbTab & FXDisplayRate(fields, "Sell", True) & vbCrLf
                count = count + 1
            End If
        End If
    Next key
    If count = 0 Then Err.Raise 5, , u("Chưa chọn dòng tỷ giá để sao chép.")
    FXRatesText = text
End Function

Public Sub FXCopyToClipboard(ByVal text As String)
#If VBA7 Then
    Dim memory As LongPtr, pointer As LongPtr
#Else
    Dim memory As Long, pointer As Long
#End If
    Dim opened As Boolean, reason As String
    On Error GoTo Failed
    If Len(text) = 0 Then Err.Raise 5, , u("Không có dữ liệu để sao chép.")
    memory = FXGlobalAlloc(66, LenB(text) + 2)
    If memory = 0 Then Err.Raise 7, , u("Không đủ bộ nhớ để sao chép.")
    pointer = FXGlobalLock(memory)
    If pointer = 0 Then Err.Raise 5, , u("Không chuẩn bị được dữ liệu sao chép.")
    FXCopyMemory pointer, StrPtr(text), LenB(text)
    FXGlobalUnlock memory: pointer = 0
    If FXOpenClipboard(Application.hwnd) = 0 Then Err.Raise 5, , u("Clipboard đang bận. Hãy thử sao chép lại.")
    opened = True
    If FXEmptyClipboard() = 0 Then Err.Raise 5, , u("Không mở được vùng sao chép.")
    If FXSetClipboardData(13, memory) = 0 Then Err.Raise 5, , u("Không đưa được dữ liệu vào clipboard.")
    memory = 0
    FXCloseClipboard
    Exit Sub
Failed:
    reason = Err.Description
    On Error Resume Next
    If pointer <> 0 Then FXGlobalUnlock memory
    If memory <> 0 Then FXGlobalFree memory
    If opened Then FXCloseClipboard
    On Error GoTo 0
    Err.Raise 5, "FXCopyToClipboard", reason
End Sub

Public Sub Ribbon_ViewRates(control As IRibbonControl)
    On Error GoTo Failed
    frmCurrency.Show
    Exit Sub
Failed:
    ShowMsg Err.Description, u("Xem tỷ giá"), 48
End Sub

Private Function FXNumber(ByVal text As String) As Double
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^([0-9]+|[0-9]{1,3}(,[0-9]{3})+)(\.[0-9]+)?$"
    If Not regex.Test(text) Then Err.Raise 5, , u("Tỷ giá nguồn bị thiếu hoặc không hợp lệ: ") & text
    FXNumber = Val(Replace(text, ",", ""))
    If FXNumber <= 0 Or FXNumber > 1000000000000# Then Err.Raise 5, , u("Tỷ giá ngoài phạm vi cho phép.")
End Function

Public Function FXParseBankXML(ByVal text As String) As Object
    Dim doc As Object, table As Object, rows As Object, row As Object, fields As Object, stamp As Object, code As String, field As Variant
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False: doc.resolveExternals = False: doc.setProperty "ProhibitDTD", True
    If Not doc.LoadXML(text) Then Err.Raise 5, , u("Không đọc được XML tỷ giá Vietcombank.")
    If doc.documentElement.nodeName <> "ExrateList" Then Err.Raise 5, , u("Nguồn trả về không phải bảng tỷ giá Vietcombank.")
    Set stamp = doc.SelectSingleNode("/ExrateList/DateTime")
    If stamp Is Nothing Then Err.Raise 5, , u("Bảng tỷ giá thiếu thời điểm công bố.")
    Set table = CreateObject("Scripting.Dictionary")
    table.Add "Updated", FXBankTime(stamp.Text)
    table.Add "Retrieved", Now
    Set rows = doc.SelectNodes("/ExrateList/Exrate")
    For Each row In rows
        code = Trim$(CStr(row.getAttribute("CurrencyCode")))
        If table.Exists(code) Then Err.Raise 5, , u("Bảng tỷ giá có mã trùng.")
        Set fields = CreateObject("Scripting.Dictionary")
        If IsNull(row.getAttribute("CurrencyName")) Then
            fields.Add "Name", code
        Else
            fields.Add "Name", Trim$(CStr(row.getAttribute("CurrencyName")))
        End If
        For Each field In Array("Buy", "Transfer", "Sell")
            If Not IsNull(row.getAttribute(field)) Then fields.Add CStr(field), Trim$(CStr(row.getAttribute(field)))
        Next field
        table.Add code, fields
    Next row
    If Not table.Exists("USD") Then Err.Raise 5, , u("Bảng tỷ giá thiếu dữ liệu USD để xác minh nguồn.")
    Set FXParseBankXML = table
End Function

Private Function FXBankTime(ByVal text As String) As Date
    Dim regex As Object, matches As Object, m As Object, month As Long, day As Long, year As Long, hour As Long, minute As Long, second As Long
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^([0-9]{1,2})/([0-9]{1,2})/([0-9]{4}) ([0-9]{1,2}):([0-9]{2}):([0-9]{2}) (AM|PM)$"
    Set matches = regex.Execute(Trim$(text))
    If matches.Count <> 1 Then Err.Raise 5, , u("Định dạng thời gian tỷ giá đã thay đổi; chưa thể dùng dữ liệu.")
    Set m = matches(0)
    month = CLng(m.SubMatches(0)): day = CLng(m.SubMatches(1)): year = CLng(m.SubMatches(2))
    hour = CLng(m.SubMatches(3)): minute = CLng(m.SubMatches(4)): second = CLng(m.SubMatches(5))
    If month < 1 Or month > 12 Or day < 1 Or day > 31 Or year < 2000 Or hour < 1 Or hour > 12 Or minute > 59 Or second > 59 Then Err.Raise 5, , u("Thời gian nguồn không hợp lệ.")
    If hour = 12 Then hour = 0
    If m.SubMatches(6) = "PM" Then hour = hour + 12
    FXBankTime = DateSerial(year, month, day) + TimeSerial(hour, minute, second)
    If VBA.Month(FXBankTime) <> month Or VBA.Day(FXBankTime) <> day Then Err.Raise 5, , u("Ngày nguồn không hợp lệ.")
End Function

Private Function FXBankTable() As Object
    Dim request As Object
    If fxLastAttempt <> 0 Then
        If DateDiff("s", fxLastAttempt, Now) < 300 Then
            If fxCached Is Nothing Then Err.Raise 5, , u("Lần lấy tỷ giá trước thất bại. Hãy thử sau 5 phút để thử lại.")
            Set FXBankTable = fxCached: Exit Function
        End If
    End If
    fxLastAttempt = Now: Set fxCached = Nothing
    Set request = HttpGet(FX_XML_URL)
    Set fxCached = FXParseBankXML(DecodeUTF8(request.ResponseBody))
    Set FXBankTable = fxCached
End Function
