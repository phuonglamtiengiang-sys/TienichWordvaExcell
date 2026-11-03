Option Explicit
Private highlightSequence As Long
Private Const MAX_INPUT_CELLS As Long = 2000000
Private Const MAX_DETAIL_ROWS As Long = 250000

Public Sub Ribbon_CompareTables(control As IRibbonControl)
    On Error GoTo Failed
    frmCompareTables.Show vbModeless
    Exit Sub
Failed:
    ShowMsg Err.Description, u("Đối chiếu hai bảng"), 48
End Sub

Public Sub ValidateHighlightRange(ByVal source As Range)
    Dim merged As Variant
    If source Is Nothing Then Err.Raise 5, , u("Hãy chọn vùng dữ liệu trên sheet cần đối chiếu.")
    If source.Areas.Count <> 1 Then Err.Raise 5, , u("Mỗi bảng cần là một khối liền nhau.")
    If source.CountLarge > 100000 Then Err.Raise 5, , u("Mỗi vùng tô màu tối đa 100.000 ô; hãy quét đúng khối dữ liệu.")
    merged = source.MergeCells
    If IsNull(merged) Then Err.Raise 5, , u("Vùng đang có ô gộp; hãy chọn vùng không gộp.")
    If merged Then Err.Raise 5, , u("Vùng đang có ô gộp; hãy chọn vùng không gộp.")
    If source.Parent.ProtectContents Or source.Parent.Parent.ReadOnly Then Err.Raise 5, , u("Trang tính bị khóa hoặc file chỉ đọc; không thể tô màu.")
End Sub

Private Function HighlightData(ByVal source As Range) As Variant
    Dim data(1 To 1, 1 To 1) As Variant
    If source.CountLarge = 1 Then
        data(1, 1) = source.Value2: HighlightData = data
    Else
        HighlightData = source.Value2
    End If
End Function

Private Function FillTheme(ByVal target As Range, ByVal pattern As Boolean) As Variant
    On Error GoTo NoTheme
    If pattern Then FillTheme = target.Interior.PatternThemeColor Else FillTheme = target.Interior.ThemeColor
    Exit Function
NoTheme:
    FillTheme = 0
End Function

Private Sub SnapshotFills(ByVal target As Range, ByVal records As Collection)
    Dim pattern As Variant, color As Variant, colorIndex As Variant, patternIndex As Variant
    Dim tint As Variant, patternTint As Variant, theme As Variant, patternTheme As Variant
    Dim record As Object, half As Long
    pattern = target.Interior.Pattern
    If Not IsNull(pattern) Then
        If pattern = xlNone Then Exit Sub
        color = target.Interior.Color: colorIndex = target.Interior.ColorIndex
        patternIndex = target.Interior.PatternColorIndex
        tint = target.Interior.TintAndShade: patternTint = target.Interior.PatternTintAndShade
        theme = FillTheme(target, False): patternTheme = FillTheme(target, True)
        If Not IsNull(colorIndex) And Not IsNull(patternIndex) And Not IsNull(tint) And Not IsNull(patternTint) And Not IsNull(theme) And Not IsNull(patternTheme) Then
            Set record = CreateObject("Scripting.Dictionary")
            record.Add "sheet", target.Parent
            record.Add "range", target
            record.Add "fill", Array(pattern, color, target.Interior.PatternColor, tint, patternTint, theme, patternTheme)
            records.Add record
            Exit Sub
        End If
    End If
    If target.CountLarge = 1 Then Err.Raise 5, , "Cannot capture cell fill."
    If target.Rows.Count >= target.Columns.Count And target.Rows.Count > 1 Then
        half = target.Rows.Count \ 2
        SnapshotFills target.Resize(half, target.Columns.Count), records
        SnapshotFills target.Offset(half, 0).Resize(target.Rows.Count - half, target.Columns.Count), records
    Else
        half = target.Columns.Count \ 2
        SnapshotFills target.Resize(target.Rows.Count, half), records
        SnapshotFills target.Offset(0, half).Resize(target.Rows.Count, target.Columns.Count - half), records
    End If
End Sub

Private Sub RestoreCompareFill(ByVal record As Object)
    Dim state As Variant, target As Range
    state = record("fill"): Set target = record("range")
    With target.Interior
        If state(5) > 0 Then .ThemeColor = state(5) Else .Color = state(1)
        .TintAndShade = state(3)
        If state(6) > 0 Then .PatternThemeColor = state(6) Else .PatternColor = state(2)
        .PatternTintAndShade = state(4)
        .Pattern = state(0)
    End With
End Sub

Private Function MaskRectangles(ByRef mask() As Byte, ByVal rows As Long, ByVal cols As Long) As Collection
    Dim active As Object, current As Object, result As New Collection
    Dim row As Long, col As Long, start As Long, kind As Byte, key As Variant, box As Variant
    Set active = CreateObject("Scripting.Dictionary")
    For row = 1 To rows
        Set current = CreateObject("Scripting.Dictionary")
        col = 1
        Do While col <= cols
            kind = mask(row, col): start = col: col = col + 1
            Do While col <= cols
                If mask(row, col) <> kind Then Exit Do
                col = col + 1
            Loop
            If kind > 0 Then
                key = kind & ":" & start & ":" & col
                If active.Exists(key) Then
                    box = active(key): box(2) = box(2) + 1
                Else
                    box = Array(row, start, 1, col - start, kind)
                End If
                current.Add key, box
            End If
        Loop
        For Each key In active.Keys
            If Not current.Exists(key) Then result.Add active(key)
        Next key
        Set active = current
    Next row
    For Each key In active.Keys
        result.Add active(key)
    Next key
    Set MaskRectangles = result
End Function

Private Sub CreateCompareColor(ByVal target As Range, ByVal color As Long, ByVal tag As String, ByVal rules As Collection)
    Dim rule As FormatCondition, record As Object
    If target Is Nothing Then Exit Sub
    Set rule = target.FormatConditions.Add(Type:=xlExpression, Formula1:="=N(""" & tag & """)=0")
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "sheet", target.Parent: record.Add "rule", rule: record.Add "tag", tag: record.Add "applies", target
    rules.Add record
    If color < 0 Then rule.Interior.Pattern = xlNone Else rule.Interior.Color = color
    rule.StopIfTrue = True
    rule.SetFirstPriority
End Sub

Private Sub PaintMask(ByVal source As Range, ByVal boxes As Collection, ByVal tag As String, ByVal rules As Collection)
    Dim kind As Long, box As Variant, batch As Range, rect As Range, count As Long, sequence As Long
    For kind = 1 To 2
        count = 0: Set batch = Nothing
        For Each box In boxes
            If box(4) = kind Then
                Set rect = source.Cells(box(0), box(1)).Resize(box(2), box(3))
                If batch Is Nothing Then Set batch = rect Else Set batch = Application.Union(batch, rect)
                count = count + 1
                If count >= 128 Then
                    sequence = sequence + 1
                    CreateCompareColor batch, IIf(kind = 1, RGB(255, 235, 156), RGB(255, 199, 156)), tag & ":" & sequence, rules
                    count = 0: Set batch = Nothing
                End If
            End If
        Next box
        If Not batch Is Nothing Then
            sequence = sequence + 1
            CreateCompareColor batch, IIf(kind = 1, RGB(255, 235, 156), RGB(255, 199, 156)), tag & ":" & sequence, rules
        End If
    Next kind
End Sub

Public Function HighlightCompareRanges(ByVal tableA As Range, ByVal tableB As Range, Optional ByVal ignoreCase As Boolean = True, Optional ByVal trimSpaces As Boolean = True, Optional ByVal tolerance As Double = 0) As Long
    Dim a As Variant, b As Variant, row As Long, col As Long, count As Long
    Dim maskA() As Byte, maskB() As Byte, boxesA As Collection, boxesB As Collection
    Dim rules As New Collection, tag As String, events As Boolean, screen As Boolean, changedState As Boolean
    Dim errorNumber As Long, reason As String, record As Object, paintStarted As Boolean
    On Error GoTo Failed
    ValidateHighlightRange tableA: ValidateHighlightRange tableB
    If tolerance < 0 Then Err.Raise 5, , u("Sai số phải là số không âm.")
    If tableA.Parent Is tableB.Parent Then
        If Not Intersect(tableA, tableB) Is Nothing Then Err.Raise 5, , u("Hai khối dữ liệu không được chồng lên nhau trên cùng sheet.")
    End If
    a = HighlightData(tableA): b = HighlightData(tableB)
    ReDim maskA(1 To tableA.Rows.Count, 1 To tableA.Columns.Count)
    ReDim maskB(1 To tableB.Rows.Count, 1 To tableB.Columns.Count)
    For row = 1 To tableA.Rows.Count
        For col = 1 To tableA.Columns.Count
            If row <= tableB.Rows.Count And col <= tableB.Columns.Count Then
                If Not SameValue(a(row, col), b(row, col), ignoreCase, trimSpaces, tolerance) Then
                    maskA(row, col) = 1: maskB(row, col) = 1: count = count + 2
                End If
            Else
                maskA(row, col) = 2: count = count + 1
            End If
        Next col
    Next row
    For row = 1 To tableB.Rows.Count
        For col = 1 To tableB.Columns.Count
            If row > tableA.Rows.Count Or col > tableA.Columns.Count Then maskB(row, col) = 2: count = count + 1
        Next col
    Next row
    Set boxesA = MaskRectangles(maskA, tableA.Rows.Count, tableA.Columns.Count)
    Set boxesB = MaskRectangles(maskB, tableB.Rows.Count, tableB.Columns.Count)
    events = Application.EnableEvents: screen = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False: changedState = True
    SnapshotFills tableA, rules: SnapshotFills tableB, rules
    paintStarted = True
    tableA.Interior.Pattern = xlNone: tableB.Interior.Pattern = xlNone
    highlightSequence = highlightSequence + 1
    tag = "TienIchExcel.Compare:" & Format$(Now, "yyyymmddhhnnss") & ":" & highlightSequence
    CreateCompareColor tableA, -1, tag & ":baseA", rules
    CreateCompareColor tableB, -1, tag & ":baseB", rules
    PaintMask tableA, boxesA, tag & ":A", rules
    PaintMask tableB, boxesB, tag & ":B", rules
    UndoRecordCompareColors rules
    Application.EnableEvents = events: Application.ScreenUpdating = screen
    HighlightCompareRanges = count
    Exit Function
Failed:
    errorNumber = Err.Number: reason = Err.Description
    On Error Resume Next
    For Each record In rules
        If record.Exists("rule") Then DeleteCompareRule record
    Next record
    If paintStarted Then
        For Each record In rules
            If record.Exists("fill") Then RestoreCompareFill record
        Next record
    End If
    If changedState Then Application.EnableEvents = events: Application.ScreenUpdating = screen
    On Error GoTo 0
    Err.Raise errorNumber, "HighlightCompareRanges", reason
End Function

Private Function CompareSheetIsOpen(ByVal sheet As Object) As Boolean
    Dim book As Workbook, candidate As Worksheet
    For Each book In Application.Workbooks
        For Each candidate In book.Worksheets
            If candidate Is sheet Then CompareSheetIsOpen = True: Exit Function
        Next candidate
    Next book
End Function

Private Sub DeleteCompareRule(ByVal record As Object)
    Dim rule As Object
    For Each rule In record("applies").FormatConditions
        If rule.Type = xlExpression Then
            If InStr(1, rule.Formula1, record("tag"), vbBinaryCompare) > 0 Then
                rule.Delete
                Exit Sub
            End If
        End If
    Next rule
End Sub

Public Function UndoCompareColors(ByVal rules As Collection) As String
    Dim record As Object, rule As Object, remove As New Collection, exists As Boolean, formula As String, pattern As Variant
    On Error GoTo Failed
    For Each record In rules
        If Not CompareSheetIsOpen(record("sheet")) Then Err.Raise 5, , u("Hãy giữ file gốc đang mở để hoàn tác màu.")
        If record("sheet").ProtectContents Then Err.Raise 5, , u("Trang tính đang bị khóa.")
        If record.Exists("fill") Then
            pattern = record("range").Interior.Pattern
            If IsNull(pattern) Then Err.Raise 5, , u("Màu nền đã được sửa sau đối chiếu; không ghi đè thay đổi mới.")
            If pattern <> xlNone Then Err.Raise 5, , u("Màu nền đã được sửa sau đối chiếu; không ghi đè thay đổi mới.")
        End If
    Next record
    For Each record In rules
        If record.Exists("rule") Then DeleteCompareRule record
    Next record
    For Each record In rules
        If record.Exists("fill") Then RestoreCompareFill record
    Next record
    UndoCompareColors = "OK"
    Exit Function
Failed:
    UndoCompareColors = u("Chưa gỡ được màu đối chiếu: ") & Err.Description
End Function

Public Sub ValidateCompareRange(ByVal source As Range)
    Dim merged As Variant
    If source Is Nothing Then Err.Raise 5, , u("Chưa chọn bảng dữ liệu.")
    If source.Areas.Count <> 1 Then Err.Raise 5, , u("Mỗi bảng phải là một vùng liền nhau.")
    If source.Rows.Count < 2 Then Err.Raise 5, , u("Chọn cả dòng tiêu đề và ít nhất một dòng dữ liệu.")
    If source.CountLarge > MAX_INPUT_CELLS Then Err.Raise 5, , u("Bảng quá lớn. Mỗi bảng tối đa 2 triệu ô; hãy chọn đúng vùng dữ liệu.")
    merged = source.MergeCells
    If IsNull(merged) Then Err.Raise 5, , u("Bảng có ô gộp. Hãy chọn vùng dữ liệu không có ô gộp.")
    If merged Then Err.Raise 5, , u("Bảng có ô gộp. Hãy chọn vùng dữ liệu không có ô gộp.")
End Sub

Private Function CompareText(ByVal value As String, ByVal ignoreCase As Boolean, ByVal trimSpaces As Boolean) As String
    If trimSpaces Then value = Trim$(Replace(value, ChrW(160), " "))
    If ignoreCase Then value = ToLowerUnicode(value)
    CompareText = value
End Function

Private Function ReadHeaders(ByVal data As Variant) As Variant
    Dim headers() As String, j As Long, used As Object, name As String
    ReDim headers(1 To UBound(data, 2))
    Set used = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(data, 2)
        If IsError(data(1, j)) Then Err.Raise 5, , u("Dòng tiêu đề có ô lỗi.")
        name = Trim$(CStr(data(1, j)))
        If name = "" Then name = u("Cột ") & j
        If used.Exists(ToLowerUnicode(name)) Then Err.Raise 5, , u("Tiêu đề cột bị trùng: ") & name
        used.Add ToLowerUnicode(name), True
        headers(j) = name
    Next j
    ReadHeaders = headers
End Function

Public Function CompareAutoPairs(ByVal sourceA As Range, ByVal sourceB As Range, ByVal keyA As Long, ByVal keyB As Long) As String
    Dim dataA As Variant, dataB As Variant, headersA As Variant, headersB As Variant
    Dim i As Long, j As Long, result As String
    ValidateCompareRange sourceA: ValidateCompareRange sourceB
    If sourceA.Columns.Count = 1 Then
        ReDim dataA(1 To 1, 1 To 1): dataA(1, 1) = sourceA.Cells(1, 1).Value2
    Else
        dataA = sourceA.Rows(1).Value2
    End If
    If sourceB.Columns.Count = 1 Then
        ReDim dataB(1 To 1, 1 To 1): dataB(1, 1) = sourceB.Cells(1, 1).Value2
    Else
        dataB = sourceB.Rows(1).Value2
    End If
    headersA = ReadHeaders(dataA): headersB = ReadHeaders(dataB)
    For i = 1 To UBound(headersA)
        If i <> keyA Then
            For j = 1 To UBound(headersB)
                If j <> keyB Then
                    If ToLowerUnicode(headersA(i)) = ToLowerUnicode(headersB(j)) Then
                        If result <> "" Then result = result & ","
                        result = result & i & ":" & j
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
    CompareAutoPairs = result
End Function

Private Function ParsePairs(ByVal spec As String, ByVal columnsA As Long, ByVal columnsB As Long, ByVal keyA As Long, ByVal keyB As Long) As Collection
    Dim result As New Collection, usedA As Object, usedB As Object, entry As Variant, pair As Variant
    Dim a As Long, b As Long
    Set usedA = CreateObject("Scripting.Dictionary"): Set usedB = CreateObject("Scripting.Dictionary")
    If spec <> "" And spec <> "-" Then
        For Each entry In Split(spec, ",")
            pair = Split(Trim$(CStr(entry)), ":")
            If UBound(pair) <> 1 Then Err.Raise 5, , u("Cặp cột so sánh không hợp lệ.")
            If Not DigitsOnly(CStr(pair(0))) Or Not DigitsOnly(CStr(pair(1))) Then Err.Raise 5, , u("Số thứ tự cột không hợp lệ.")
            a = CLng(pair(0)): b = CLng(pair(1))
            If a < 1 Or a > columnsA Or b < 1 Or b > columnsB Then Err.Raise 5, , u("Cột so sánh nằm ngoài bảng.")
            If a = keyA Or b = keyB Then Err.Raise 5, , u("Không ghép cột mã vào danh sách cột thông tin.")
            If usedA.Exists(CStr(a)) Or usedB.Exists(CStr(b)) Then Err.Raise 5, , u("Mỗi cột chỉ được ghép một lần.")
            usedA.Add CStr(a), True: usedB.Add CStr(b), True
            result.Add Array(a, b)
        Next entry
    End If
    Set ParsePairs = result
End Function

Private Function RowHasData(ByVal data As Variant, ByVal row As Long) As Boolean
    Dim j As Long
    For j = 1 To UBound(data, 2)
        If IsError(data(row, j)) Then RowHasData = True: Exit Function
        If Len(CStr(data(row, j))) > 0 Then RowHasData = True: Exit Function
    Next j
End Function

Private Sub IndexTable(ByVal data As Variant, ByVal keyColumn As Long, ByVal source As Range, ByVal side As String, ByVal ignoreCase As Boolean, ByVal trimSpaces As Boolean, ByVal index As Object, ByVal keys As Object, ByVal issues As Collection, ByRef blankRows As Long)
    Dim row As Long, key As String, rows As Collection
    For row = 2 To UBound(data, 1)
        If Not RowHasData(data, row) Then
            blankRows = blankRows + 1
        ElseIf IsError(data(row, keyColumn)) Then
            issues.Add Array(side, source.Row + row - 1, data(row, keyColumn), u("Ô mã đang có lỗi; chưa ghép dòng."))
        Else
            key = CompareText(CStr(data(row, keyColumn)), ignoreCase, trimSpaces)
            If key = "" Then
                issues.Add Array(side, source.Row + row - 1, "", u("Mã trống; chưa ghép dòng."))
            Else
                If Not index.Exists(key) Then
                    Set rows = New Collection
                    index.Add key, rows
                End If
                Set rows = index(key): rows.Add row
                If Not keys.Exists(key) Then keys.Add key, True
            End If
        End If
    Next row
End Sub

Private Function NumberValue(ByVal value As Variant) As Boolean
    Select Case VarType(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, 14: NumberValue = True
    End Select
End Function

Private Function SameValue(ByVal a As Variant, ByVal b As Variant, ByVal ignoreCase As Boolean, ByVal trimSpaces As Boolean, ByVal tolerance As Double) As Boolean
    If IsError(a) Or IsError(b) Then
        If IsError(a) And IsError(b) Then SameValue = (CStr(a) = CStr(b))
        Exit Function
    End If
    If NumberValue(a) And NumberValue(b) Then
        SameValue = (Abs(CDbl(a) - CDbl(b)) <= tolerance)
    ElseIf NumberValue(a) Or NumberValue(b) Then
        ' Numeric zero is different from an empty cell; numeric text is distinct.
        SameValue = False
    ElseIf VarType(a) = vbBoolean Or VarType(b) = vbBoolean Then
        If VarType(a) = VarType(b) Then SameValue = (a = b)
    Else
        SameValue = (CompareText(CStr(a), ignoreCase, trimSpaces) = CompareText(CStr(b), ignoreCase, trimSpaces))
    End If
End Function

Private Sub AddAmbiguousRows(ByVal rows As Collection, ByVal data As Variant, ByVal keyColumn As Long, ByVal source As Range, ByVal side As String, ByVal issues As Collection)
    Dim row As Variant
    For Each row In rows
        issues.Add Array(side, source.Row + CLng(row) - 1, data(CLng(row), keyColumn), u("Mã trùng ở một hoặc cả hai bảng; không tự ghép."))
    Next row
End Sub

Private Sub BeginReportSheet(ByVal ws As Worksheet, ByVal title As String, ByVal headers As Variant)
    Dim j As Long
    ws.Cells.Font.Name = "Calibri": ws.Cells.Font.Size = 11
    ws.Range("A1").Value2 = title
    ws.Range("A1").Font.Size = 18: ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(23, 65, 85)
    For j = LBound(headers) To UBound(headers)
        WriteText ws.Cells(4, j - LBound(headers) + 1), CStr(headers(j))
    Next j
    With ws.Cells(4, 1).Resize(1, UBound(headers) - LBound(headers) + 1)
        .Interior.Color = RGB(23, 65, 85): .Font.Color = vbWhite: .Font.Bold = True
        .WrapText = True: .RowHeight = 32
    End With
    ws.Columns.ColumnWidth = 20
End Sub

Private Sub FinishReportSheet(ByVal ws As Worksheet, ByVal rows As Long, ByVal columns As Long)
    Dim table As ListObject
    If rows = 0 Then
        WriteText ws.Cells(5, 1), u("Không có bản ghi trong nhóm này.")
    Else
        Set table = ws.ListObjects.Add(xlSrcRange, ws.Cells(4, 1).Resize(rows + 1, columns), , xlYes)
        table.TableStyle = "TableStyleMedium2"
        ws.Cells(5, 1).Resize(rows, columns).VerticalAlignment = xlTop
        ws.Cells(5, 1).Resize(rows, columns).WrapText = True
        If rows <= 1000 Then ws.Cells(5, 1).Resize(rows, columns).Rows.AutoFit
    End If
    ws.Activate
    If Application.Visible Then
        ActiveWindow.SplitRow = 4: ActiveWindow.FreezePanes = True
    End If
    ws.Range("A1").Select
End Sub

Private Sub WriteSourceRows(ByVal ws As Worksheet, ByVal title As String, ByVal data As Variant, ByVal headers As Variant, ByVal source As Range, ByVal rows As Collection)
    Dim output() As Variant, labels() As Variant, i As Long, j As Long, columns As Long, format As Variant
    columns = UBound(data, 2)
    ReDim labels(0 To columns)
    For j = 1 To columns: labels(j - 1) = headers(j): Next j
    labels(columns) = u("Dòng nguồn")
    BeginReportSheet ws, title, labels
    If rows.Count > 0 Then
        ReDim output(1 To rows.Count, 1 To columns + 1)
        For i = 1 To rows.Count
            For j = 1 To columns: output(i, j) = data(CLng(rows(i)), j): Next j
            output(i, columns + 1) = source.Row + CLng(rows(i)) - 1
        Next i
        With ws.Cells(5, 1).Resize(rows.Count, columns + 1)
            .NumberFormat = "@": .Value2 = output
        End With
        For j = 1 To columns
            format = source.Columns(j).NumberFormat
            If IsNull(format) Then format = source.Cells(2, j).NumberFormat
            ws.Cells(5, j).Resize(rows.Count, 1).NumberFormat = format
        Next j
        ws.Cells(5, columns + 1).Resize(rows.Count, 1).NumberFormat = "0"
    End If
    FinishReportSheet ws, rows.Count, columns + 1
End Sub

Private Sub WriteRecords(ByVal ws As Worksheet, ByVal title As String, ByVal headers As Variant, ByVal records As Collection)
    Dim output() As Variant, i As Long, j As Long, record As Variant, columns As Long
    columns = UBound(headers) - LBound(headers) + 1
    BeginReportSheet ws, title, headers
    If records.Count > 0 Then
        ReDim output(1 To records.Count, 1 To columns)
        For i = 1 To records.Count
            record = records(i)
            For j = 1 To columns: output(i, j) = record(j - 1): Next j
        Next i
        With ws.Cells(5, 1).Resize(records.Count, columns)
            .NumberFormat = "@": .Value2 = output
        End With
    End If
    FinishReportSheet ws, records.Count, columns
End Sub

Public Function CompareTables(ByVal tableA As Range, ByVal tableB As Range, ByVal keyA As Long, ByVal keyB As Long, Optional ByVal pairSpec As String = "", Optional ByVal ignoreCase As Boolean = True, Optional ByVal trimSpaces As Boolean = True, Optional ByVal numberTolerance As Double = 0) As Workbook
    Dim dataA As Variant, dataB As Variant, headersA As Variant, headersB As Variant, pairs As Collection
    Dim indexA As Object, indexB As Object, keys As Object, key As Variant, rowsA As Collection, rowsB As Collection
    Dim onlyA As New Collection, onlyB As New Collection, differences As New Collection, issues As New Collection
    Dim blankA As Long, blankB As Long, duplicateKeys As Long, sameRows As Long, differentRows As Long
    Dim rowA As Long, rowB As Long, pair As Variant, different As Boolean, checked As Long
    Dim report As Workbook, summary As Worksheet, wsA As Worksheet, wsB As Worksheet, wsDiff As Worksheet, wsIssues As Worksheet
    Dim oldScreen As Boolean, oldEvents As Boolean, oldCancel As Long, oldStatus As Variant, captured As Boolean
    Dim errorNumber As Long, description As String, mapping As String, record As Variant, i As Long
    On Error GoTo Failed
    ValidateCompareRange tableA: ValidateCompareRange tableB
    If tableA.CountLarge + tableB.CountLarge > MAX_INPUT_CELLS Then Err.Raise 5, , u("Tổng hai bảng tối đa 2 triệu ô. Hãy thu gọn vùng chọn.")
    If keyA < 1 Or keyA > tableA.Columns.Count Or keyB < 1 Or keyB > tableB.Columns.Count Then Err.Raise 5, , u("Cột mã nằm ngoài bảng.")
    If numberTolerance < 0 Then Err.Raise 5, , u("Sai số phải lớn hơn hoặc bằng 0.")
    dataA = tableA.Value2: dataB = tableB.Value2
    headersA = ReadHeaders(dataA): headersB = ReadHeaders(dataB)
    If pairSpec = "" Then pairSpec = CompareAutoPairs(tableA, tableB, keyA, keyB)
    Set pairs = ParsePairs(pairSpec, UBound(dataA, 2), UBound(dataB, 2), keyA, keyB)
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    oldStatus = Application.StatusBar: oldCancel = Application.EnableCancelKey: captured = True
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.EnableCancelKey = xlErrorHandler
    Application.StatusBar = u("Đang đọc và lập chỉ mục hai bảng...")
    Set indexA = CreateObject("Scripting.Dictionary"): Set indexB = CreateObject("Scripting.Dictionary")
    Set keys = CreateObject("Scripting.Dictionary")
    IndexTable dataA, keyA, tableA, "A", ignoreCase, trimSpaces, indexA, keys, issues, blankA
    IndexTable dataB, keyB, tableB, "B", ignoreCase, trimSpaces, indexB, keys, issues, blankB
    For Each key In keys.Keys
        Set rowsA = Nothing: Set rowsB = Nothing
        If indexA.Exists(key) Then Set rowsA = indexA(key)
        If indexB.Exists(key) Then Set rowsB = indexB(key)
        different = False
        If Not rowsA Is Nothing Then
            If rowsA.Count > 1 Then different = True
        End If
        If Not rowsB Is Nothing Then
            If rowsB.Count > 1 Then different = True
        End If
        If different Then
            duplicateKeys = duplicateKeys + 1
            If Not rowsA Is Nothing Then AddAmbiguousRows rowsA, dataA, keyA, tableA, "A", issues
            If Not rowsB Is Nothing Then AddAmbiguousRows rowsB, dataB, keyB, tableB, "B", issues
        ElseIf rowsA Is Nothing Then
            onlyB.Add CLng(rowsB(1))
        ElseIf rowsB Is Nothing Then
            onlyA.Add CLng(rowsA(1))
        Else
            rowA = CLng(rowsA(1)): rowB = CLng(rowsB(1)): different = False
            For Each pair In pairs
                If Not SameValue(dataA(rowA, pair(0)), dataB(rowB, pair(1)), ignoreCase, trimSpaces, numberTolerance) Then
                    different = True
                    If differences.Count >= MAX_DETAIL_ROWS Then Err.Raise 5, , u("Có hơn 250.000 ô khác nhau. Hãy đối chiếu ít cột hơn hoặc chia nhỏ dữ liệu.")
                    differences.Add Array(CStr(dataA(rowA, keyA)), headersA(pair(0)), headersB(pair(1)), dataA(rowA, pair(0)), dataB(rowB, pair(1)), tableA.Row + rowA - 1, tableB.Row + rowB - 1)
                End If
            Next pair
            If different Then differentRows = differentRows + 1 Else sameRows = sameRows + 1
        End If
        checked = checked + 1
        If checked Mod 1000 = 0 Then
            Application.StatusBar = u("Đang đối chiếu mã: ") & checked & " / " & keys.Count
            DoEvents
        End If
    Next key
    Application.StatusBar = u("Đang tạo báo cáo đối chiếu...")
    Set report = Application.Workbooks.Add(xlWBATWorksheet)
    Set summary = report.Worksheets(1): summary.Name = u("Tổng quan")
    Set wsA = report.Worksheets.Add(After:=summary): wsA.Name = u("Chỉ có ở A")
    Set wsB = report.Worksheets.Add(After:=wsA): wsB.Name = u("Chỉ có ở B")
    Set wsDiff = report.Worksheets.Add(After:=wsB): wsDiff.Name = u("Khác thông tin")
    Set wsIssues = report.Worksheets.Add(After:=wsDiff): wsIssues.Name = u("Cần kiểm tra")
    WriteSourceRows wsA, u("Các mã chỉ có ở bảng A"), dataA, headersA, tableA, onlyA
    WriteSourceRows wsB, u("Các mã chỉ có ở bảng B"), dataB, headersB, tableB, onlyB
    WriteRecords wsDiff, u("Cùng mã nhưng khác thông tin"), Array(u("Mã"), u("Cột bảng A"), u("Cột bảng B"), u("Giá trị A"), u("Giá trị B"), u("Dòng A"), u("Dòng B")), differences
    wsDiff.Columns("B:E").ColumnWidth = 30
    WriteRecords wsIssues, u("Dòng có mã trống, ô mã lỗi hoặc mã trùng"), Array(u("Bảng"), u("Dòng nguồn"), u("Mã"), u("Lý do")), issues
    wsIssues.Columns("D").ColumnWidth = 65
    With summary
        .Cells.Font.Name = "Calibri": .Cells.Font.Size = 11
        .Columns("A").ColumnWidth = 37: .Columns("B").ColumnWidth = 70
        WriteText .Range("A1"), u("ĐỐI CHIẾU HAI BẢNG THEO MÃ")
        .Range("A1").Font.Size = 18: .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(23, 65, 85)
        WriteText .Range("A3"), u("Bảng A")
        WriteText .Range("B3"), tableA.Parent.Parent.FullName & " | " & tableA.Parent.Name & "!" & tableA.Address(False, False)
        WriteText .Range("A4"), u("Bảng B")
        WriteText .Range("B4"), tableB.Parent.Parent.FullName & " | " & tableB.Parent.Name & "!" & tableB.Address(False, False)
        WriteText .Range("A5"), u("Cột mã A / B")
        WriteText .Range("B5"), headersA(keyA) & " / " & headersB(keyB)
        WriteText .Range("A7"), u("Chỉ có ở A"): .Range("B7").Value2 = onlyA.Count
        WriteText .Range("A8"), u("Chỉ có ở B"): .Range("B8").Value2 = onlyB.Count
        WriteText .Range("A9"), u("Cùng mã, khác thông tin (dòng)"): .Range("B9").Value2 = differentRows
        WriteText .Range("A10"), u("Số ô khác thông tin"): .Range("B10").Value2 = differences.Count
        WriteText .Range("A11"), u("Cùng mã, thông tin giống nhau"): .Range("B11").Value2 = sameRows
        WriteText .Range("A12"), u("Nhóm mã trùng, chưa ghép"): .Range("B12").Value2 = duplicateKeys
        WriteText .Range("A13"), u("Dòng cần kiểm tra"): .Range("B13").Value2 = issues.Count
        WriteText .Range("A14"), u("Dòng hoàn toàn trống A / B"): WriteText .Range("B14"), blankA & " / " & blankB
        WriteText .Range("A16"), u("Bỏ qua hoa/thường"): WriteText .Range("B16"), IIf(ignoreCase, u("Có"), u("Không"))
        WriteText .Range("A17"), u("Bỏ khoảng trắng đầu/cuối"): WriteText .Range("B17"), IIf(trimSpaces, u("Có"), u("Không"))
        WriteText .Range("A18"), u("Sai số tuyệt đối cho giá trị số"): .Range("B18").Value2 = numberTolerance
        WriteText .Range("A20"), u("Cặp cột đã so sánh")
        i = 21
        For Each pair In pairs
            WriteText .Cells(i, 1), headersA(pair(0)): WriteText .Cells(i, 2), headersB(pair(1))
            i = i + 1
        Next pair
        If pairs.Count = 0 Then WriteText .Cells(i, 1), u("Chỉ kiểm tra sự có mặt của mã; chưa so sánh cột thông tin."): i = i + 1
        WriteText .Cells(i + 1, 1), u("Cách đọc báo cáo")
        WriteText .Cells(i + 2, 1), u("Chỉ có ở A/B: toàn bộ giá trị của dòng nguồn; cột cuối là dòng nguồn.")
        WriteText .Cells(i + 3, 1), u("Khác thông tin: mỗi dòng là một ô khác nhau; hai cột cuối chỉ dòng A/B.")
        WriteText .Cells(i + 4, 1), u("Mã trùng không được tự ghép, kể cả khi chỉ một bên bị trùng.")
        WriteText .Cells(i + 5, 1), u("Mã lấy theo giá trị ô: 00123 dạng Text khác 123. Số dạng Text khác số ở cột thông tin.")
        WriteText .Cells(i + 6, 1), u("Báo cáo là ảnh chụp giá trị tại lúc chạy; không tự cập nhật khi sửa bảng nguồn.")
        For checked = i + 2 To i + 6
            .Range(.Cells(checked, 1), .Cells(checked, 2)).Merge
            .Cells(checked, 1).WrapText = True
            .Rows(checked).RowHeight = 32
        Next checked
        .Range("A3:B5").WrapText = True: .Range("A3:B5").Rows.AutoFit
        .Range("A7:B13").Borders.Color = RGB(220, 230, 235)
        .Range("B7:B13").NumberFormat = "#,##0": .Range("B7:B13").Font.Bold = True
        .Range("A7:B8").Interior.Color = RGB(230, 241, 247)
        .Range("A9:B10").Interior.Color = RGB(255, 243, 219)
    End With
    summary.Activate: summary.Range("A1").Select
    Set CompareTables = report
Done:
    If captured Then
        Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
        Application.StatusBar = oldStatus: Application.EnableCancelKey = oldCancel
    End If
    Exit Function
Failed:
    errorNumber = Err.Number: description = Err.Description
    On Error Resume Next
    If Not report Is Nothing Then report.Close False
    If captured Then
        Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
        Application.StatusBar = oldStatus: Application.EnableCancelKey = oldCancel
    End If
    On Error GoTo 0
    If errorNumber = 18 Then description = u("Đã hủy đối chiếu. Hai bảng nguồn được giữ nguyên.")
    Err.Raise errorNumber, "CompareTables", description
End Function
