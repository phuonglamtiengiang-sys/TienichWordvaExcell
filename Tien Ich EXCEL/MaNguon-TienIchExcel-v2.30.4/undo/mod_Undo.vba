Option Explicit
Private history As Collection
Private pending As Collection
Private pendingName As String

Private Function CellState(ByVal cell As Range) As Variant
    CellState = Array(cell.HasFormula, cell.Formula, cell.Value2, cell.NumberFormat)
End Function

Private Function SameState(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim i As Long
    For i = 0 To 3
        If VarType(a(i)) <> VarType(b(i)) Then Exit Function
        If IsError(a(i)) Then
            If CStr(a(i)) <> CStr(b(i)) Then Exit Function
        ElseIf IsNull(a(i)) Or IsNull(b(i)) Then
            Exit Function
        ElseIf CStr(a(i)) <> CStr(b(i)) Then
            Exit Function
        End If
    Next i
    SameState = True
End Function

Public Sub UndoBegin(ByVal selected As Range, ByVal outputColumns As Long, ByVal title As String)
    Dim cell As Range, target As Range, j As Long, record As Object
    Set pending = Nothing
    Set pending = New Collection
    pendingName = title
    For Each cell In selected.Cells
        For j = IIf(outputColumns = 0, 0, 1) To outputColumns
            Set target = cell.Offset(0, j)
            Set record = CreateObject("Scripting.Dictionary")
            record.Add "cell", target
            record.Add "before", CellState(target)
            pending.Add record
        Next j
    Next cell
End Sub

Public Sub UndoCommit()
    Dim record As Object, changed As Collection, entry As Object
    If pending Is Nothing Then Exit Sub
    Set changed = New Collection
    For Each record In pending
        record.Add "after", CellState(record("cell"))
        If Not SameState(record("before"), record("after")) Then changed.Add record
    Next record
    Set pending = Nothing
    If changed.Count = 0 Then Exit Sub
    If history Is Nothing Then Set history = New Collection
    Set entry = CreateObject("Scripting.Dictionary")
    entry.Add "items", changed
    entry.Add "title", pendingName
    history.Add entry
    If history.Count > 20 Then history.Remove 1
End Sub

Public Function UndoCount() As Long
    If Not history Is Nothing Then UndoCount = history.Count
End Function

Public Sub UndoRecordCompareColors(ByVal rules As Collection)
    Dim entry As Object
    If rules.Count = 0 Then Exit Sub
    If history Is Nothing Then Set history = New Collection
    Set entry = CreateObject("Scripting.Dictionary")
    entry.Add "highlightRules", rules
    history.Add entry
    If history.Count > 20 Then history.Remove 1
End Sub

' Public result permits automated tests without displaying a modal dialog.
Public Function UndoRestore() As String
    Dim entry As Object, record As Object, cell As Range, state As Variant
    Dim events As Boolean, screen As Boolean, restoring As Boolean, reason As String
    On Error GoTo Failed
    If UndoCount() = 0 Then
        UndoRestore = u("Không có thao tác được hỗ trợ để hoàn tác.")
        Exit Function
    End If
    Set entry = history(history.Count)
    If entry.Exists("highlightRules") Then
        UndoRestore = UndoCompareColors(entry("highlightRules"))
        If UndoRestore = "OK" Then history.Remove history.Count
        Exit Function
    End If
    For Each record In entry("items")
        Set cell = record("cell")
        ' Refer to the original worksheet object, never resolve by file/sheet name.
        If cell.Parent.Parent Is Nothing Then Err.Raise 5
        If cell.Parent.ProtectContents Or cell.MergeCells Then Err.Raise 5, , u("Trang tính đang khóa hoặc ô đã được gộp.")
        If Not SameState(CellState(cell), record("after")) Then Err.Raise 5, , u("Ô kết quả đã thay đổi sau thao tác. Không hoàn tác để tránh ghi đè dữ liệu mới.")
    Next record
    events = Application.EnableEvents: screen = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    restoring = True
    For Each record In entry("items")
        Set cell = record("cell"): state = record("before")
        RestoreState cell, state
    Next record
    history.Remove history.Count
    Application.EnableEvents = events: Application.ScreenUpdating = screen
    UndoRestore = "OK"
    Exit Function
Failed:
    reason = Err.Description
    If restoring Then
        On Error Resume Next
        For Each record In entry("items")
            RestoreState record("cell"), record("after")
        Next record
        Application.EnableEvents = events: Application.ScreenUpdating = screen
        On Error GoTo 0
    End If
    UndoRestore = u("Chưa thể hoàn tác. Hãy kiểm tra tệp/trang tính còn mở. ") & reason
End Function

Private Sub RestoreState(ByVal cell As Range, ByVal state As Variant)
    If state(0) Then
        cell.NumberFormat = "General"
        cell.Formula = state(1)
    Else
        ' Prevent literal strings beginning with = from becoming formulas.
        cell.NumberFormat = "@"
        cell.Value2 = state(2)
    End If
    cell.NumberFormat = state(3)
End Sub

Public Sub UndoLastAction()
    Dim result As String
    result = UndoRestore()
    If result <> "OK" Then ShowMsg result, u("Hoàn tác"), 48
End Sub

Public Sub Ribbon_Undo(control As IRibbonControl)
    UndoLastAction
End Sub

Public Sub UndoClearHistory()
    Set history = Nothing
    Set pending = Nothing
End Sub
