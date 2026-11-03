Option Explicit
Private mRows As Variant
Private mIndex As Object
Private Const DATA_SHEET As String = "AddressData"

Private Function AddressPart(ByVal value As String) As String
    Dim s As String, prefix As Variant
    s = LCase$(BoDau(value))
    s = Replace(Replace(s, ChrW(160), " "), ".", " ")
    s = Application.Trim(s)
    For Each prefix In Array("thanh pho ", "thi xa ", "thi tran ", "dac khu ", "phuong ", "huyen ", "quan ", "tinh ", "xa ", "tp ", "tx ", "tt ", "p ", "q ", "h ", "t ", "x ")
        If Left$(s, Len(prefix)) = prefix Then
            s = Trim$(Mid$(s, Len(prefix) + 1))
            Exit For
        End If
    Next
    If s = "hcm" Or s = "tphcm" Or s = "ho chi minh city" Then s = "ho chi minh"
    If s = "thua thien hue" Then s = "hue"
    s = Replace(s, "nt moc chau", "nong truong moc chau")
    AddressPart = s
End Function

Private Sub AddAddressIndex(ByVal key As String, ByVal row As Long)
    Dim items As Collection
    If mIndex.Exists(key) Then
        Set items = mIndex(key)
    Else
        Set items = New Collection
        mIndex.Add key, items
    End If
    items.Add row
End Sub

Private Sub LoadAddressData()
    Dim i As Long, p As String, d As String, w As String
    If Not mIndex Is Nothing Then Exit Sub
    mRows = ThisWorkbook.Worksheets(DATA_SHEET).UsedRange.Value2
    Set mIndex = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(mRows, 1)
        w = AddressPart(CStr(mRows(i, 2)))
        d = AddressPart(CStr(mRows(i, 3)))
        p = AddressPart(CStr(mRows(i, 4)))
        AddAddressIndex "F|" & p & "|" & d & "|" & w, i
        AddAddressIndex "S|" & p & "|" & w, i
        AddAddressIndex "C|" & CStr(mRows(i, 1)), i
    Next
End Sub

' Only complete comma-separated administrative segments are matched.
' Street and house-number segments are copied verbatim; partial splits never auto-convert.
Public Function AddressResolve(ByVal value As Variant, ByRef result As String, ByRef status As String, ByRef candidates As String, ByRef legal As String) As Boolean
    Dim parts As Variant, n As Long, start As Long, i As Long, r As Variant
    Dim key As String, matches As Collection, oldCodes As Object, newCodes As Object
    Dim allFull As Boolean, hasUnknown As Boolean, historical As Boolean, prefix As String, p As String, d As String, w As String
    result = "": status = "": candidates = "": legal = ""
    On Error GoTo Failed
    If IsError(value) Then
        status = u("Ô nguồn có lỗi Excel"): Exit Function
    End If
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    If Len(Trim$(CStr(value))) = 0 Then Exit Function
    LoadAddressData
    If Len(Trim$(CStr(value))) = 5 And DigitsOnly(Trim$(CStr(value))) Then
        key = "C|" & Trim$(CStr(value))
        If mIndex.Exists(key) Then Set matches = mIndex(key)
        If matches Is Nothing Then GoTo Missing
        start = 0
        GoTo EvaluateMatches
    End If
    parts = Split(CStr(value), ",")
    n = UBound(parts)
    Do While n >= 0
        If Len(Trim$(parts(n))) > 0 Then Exit Do
        n = n - 1
    Loop
    If n < 1 Then GoTo Missing
    If AddressPart(CStr(parts(n))) = "viet nam" Or AddressPart(CStr(parts(n))) = "vietnam" Then n = n - 1
    If n < 1 Then GoTo Missing
    p = AddressPart(CStr(parts(n)))
    If n >= 2 Then
        d = AddressPart(CStr(parts(n - 1)))
        w = AddressPart(CStr(parts(n - 2)))
        key = "F|" & p & "|" & d & "|" & w
        If mIndex.Exists(key) Then
            Set matches = mIndex(key)
            start = n - 2
        End If
    End If
    If matches Is Nothing Then
        key = "S|" & p & "|" & AddressPart(CStr(parts(n - 1)))
        If mIndex.Exists(key) Then
            Set matches = mIndex(key)
            start = n - 1
        End If
    End If
    If matches Is Nothing Then GoTo Missing
EvaluateMatches:
    Set oldCodes = CreateObject("Scripting.Dictionary")
    Set newCodes = CreateObject("Scripting.Dictionary")
    allFull = True
    For Each r In matches
        oldCodes(CStr(mRows(r, 1))) = True
        key = CStr(mRows(r, 8))
        If Len(key) > 0 And Not newCodes.Exists(key) Then
            newCodes.Add key, True
            If Len(candidates) > 0 Then candidates = candidates & vbLf
            candidates = candidates & CStr(mRows(r, 5)) & ", " & CStr(mRows(r, 6))
        End If
        If LCase$(CStr(mRows(r, 7))) <> "full" Then allFull = False
        If LCase$(CStr(mRows(r, 7))) = "unknown" Then hasUnknown = True
        If Left$(CStr(mRows(r, 10)), 3) = "NSO" Then historical = True
        If Len(CStr(mRows(r, 9))) > 0 Then
            If InStr(1, ";" & legal & ";", ";" & CStr(mRows(r, 9)) & ";", vbBinaryCompare) = 0 Then
                If Len(legal) > 0 Then legal = legal & ";"
                legal = legal & CStr(mRows(r, 9))
            End If
        End If
    Next
    If hasUnknown Then
        status = u("Cần kiểm tra: chưa đủ căn cứ nối lịch sử 2024–2025; gợi ý (nếu có) chưa phải kết quả xác nhận")
    ElseIf oldCodes.Count > 1 Then
        status = u("Cần kiểm tra: trùng tên hoặc khác thời kỳ; bổ sung quận/huyện, đối chiếu mã xã cũ và ngày hiệu lực")
    ElseIf newCodes.Count <> 1 Or Not allFull Then
        status = u("Cần kiểm tra: xã/phường cũ bị chia hoặc sáp nhập một phần")
    Else
        For i = 0 To start - 1
            If Len(Trim$(parts(i))) > 0 Then prefix = prefix & parts(i) & ","
        Next
        If Len(prefix) > 0 Then prefix = prefix & " "
        result = prefix & candidates
        status = u("Đã đối chiếu: sáp nhập toàn bộ")
        If historical Then status = u("Đã đối chiếu qua lịch sử sáp nhập 2024–2025")
        AddressResolve = True
    End If
    Exit Function
Missing:
    status = u("Chưa khớp: cần xã/phường, quận/huyện cũ, tỉnh/thành cũ, cách nhau bằng dấu phẩy; kiểm tra địa chỉ đã mới hoặc ngoài danh mục")
    Exit Function
Failed:
    status = u("Không xử lý được: ") & Err.Description
End Function

Public Function DiaChiMoi(ByVal value As Variant) As Variant
    Dim result As String, status As String, candidates As String, legal As String
    If IsError(value) Then DiaChiMoi = value: Exit Function
    AddressResolve value, result, status, candidates, legal
    DiaChiMoi = result
End Function

Public Function DiaChiMoiTrangThai(ByVal value As Variant) As String
    Dim result As String, status As String, candidates As String, legal As String
    AddressResolve value, result, status, candidates, legal
    DiaChiMoiTrangThai = status
End Function

Public Function DiaChiMoiGoiY(ByVal value As Variant) As String
    Dim result As String, status As String, candidates As String, legal As String
    AddressResolve value, result, status, candidates, legal
    DiaChiMoiGoiY = candidates
End Function

Public Sub Ribbon_ConvertAddress(control As IRibbonControl)
    Dim source As Range, report As Workbook
    If Not SelectedInput(source, 0) Then Exit Sub
    On Error GoTo Failed
    Set report = AddressReport(source)
    report.Activate
    Exit Sub
Failed:
    ShowMsg Err.Description, "Chuyen dia chi", 48
End Sub

Public Function AddressReport(ByVal source As Range) As Workbook
    Dim book As Workbook, sheet As Worksheet, c As Range, data() As Variant
    Dim i As Long, result As String, status As String, candidates As String, legal As String
    Dim oldScreen As Boolean, errorNumber As Long, errorText As String
    If source Is Nothing Then Err.Raise 5, , "Hay chon cot dia chi."
    If source.CountLarge > 1000 Then Err.Raise 5, , "Toi da 1000 dia chi moi lan."
    If source.Columns.Count <> 1 Or source.Areas.Count <> 1 Then Err.Raise 5, , "Hay chon mot cot lien tuc."
    For Each c In source.Cells
        If c.MergeCells Then Err.Raise 5, , "Khong xu ly o gop."
    Next
    LoadAddressData
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    ReDim data(1 To source.Cells.Count, 1 To 6)
    For Each c In source.Cells
        i = i + 1
        AddressResolve c.Value2, result, status, candidates, legal
        If IsError(c.Value2) Then data(i, 1) = c.Text Else data(i, 1) = CStr(c.Value2)
        data(i, 2) = result: data(i, 3) = status: data(i, 4) = candidates
        data(i, 5) = legal: data(i, 6) = c.Address(External:=True)
    Next
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1)
    sheet.Name = u("Địa chỉ mới")
    sheet.Range("A1:F1").Merge
    sheet.Range("A1").Value2 = u("CHUYỂN ĐỊA CHỈ HÀNH CHÍNH CŨ → MỚI")
    sheet.Range("A2:F2").Merge
    sheet.Range("A2").Value2 = u("Nguồn: Cục Thống kê (lịch sử 2024–2025), NQ1191 và tinhthanhpho.com (CC BY 4.0, bản 2026.10.8) | Chỉ tự đổi khi đủ căn cứ")
    sheet.Hyperlinks.Add sheet.Range("A2"), "https://danhmuchanhchinh.nso.gov.vn/Doi_Chieu_Moi.aspx"
    sheet.Range("A3:F3").Merge
    sheet.Range("A3").Value2 = u("Chỉ điền địa chỉ mới khi khớp một xã cũ và sáp nhập toàn bộ. Dòng cần kiểm tra để trống kết quả; xem gợi ý và nghị quyết.")
    sheet.Range("A5").Value2 = u("Địa chỉ cũ")
    sheet.Range("B5").Value2 = u("Địa chỉ mới")
    sheet.Range("C5").Value2 = u("Trạng thái")
    sheet.Range("D5").Value2 = u("Xã/phường, tỉnh/thành mới có thể tương ứng")
    sheet.Range("E5").Value2 = u("Nghị quyết theo nguồn dữ liệu")
    sheet.Range("F5").Value2 = u("Ô nguồn")
    With sheet.Range("A6").Resize(i, 6)
        .NumberFormat = "@"
        .Value2 = data
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    sheet.Columns("A:B").ColumnWidth = 58
    sheet.Columns("C:D").ColumnWidth = 45
    sheet.Columns("E:F").ColumnWidth = 30
    sheet.Range("A1:F3").WrapText = True
    sheet.Rows("1:3").RowHeight = 30
    sheet.Range("A1").Font.Size = 15
    sheet.Range("A5:F5").Font.Bold = True
    sheet.Range("A5:F5").Interior.Color = RGB(26, 66, 106)
    sheet.Range("A5:F5").Font.Color = vbWhite
    sheet.Range("A5:F5").WrapText = True
    sheet.Rows(5).RowHeight = 32
    sheet.Range("A5").Resize(i + 1, 6).AutoFilter
    sheet.Range("A6").Resize(i, 6).Rows.AutoFit
    If Application.Visible Then
        sheet.Activate
        sheet.Range("A6").Select
        ActiveWindow.FreezePanes = True
    End If
    Application.ScreenUpdating = oldScreen
    Set AddressReport = book
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    Err.Raise errorNumber, , errorText
End Function
