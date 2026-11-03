Option Explicit
Private printSelectionBusy As Boolean

Private Function RestoreSelectionPrintSetup(ByVal sheet As Worksheet, ByVal area As String, ByVal zoom As Variant, ByVal wide As Variant, ByVal tall As Variant, ByVal titleRows As String, ByVal titleCols As String) As Boolean
    On Error Resume Next
    Err.Clear
    With sheet.PageSetup
        .PrintArea = area
        .FitToPagesWide = wide: .FitToPagesTall = tall: .Zoom = zoom
        .PrintTitleRows = titleRows: .PrintTitleColumns = titleCols
    End With
    RestoreSelectionPrintSetup = (Err.Number = 0)
End Function

Public Sub PreviewSelectedRange(ByVal target As Range)
    Dim sheet As Worksheet, view As Window, selected As Object
    Dim area As String, zoom As Variant, wide As Variant, tall As Variant, titleRows As String, titleCols As String
    Dim oldStatus As Variant, oldCancel As Long, events As Boolean, screen As Boolean, wasSaved As Boolean
    Dim started As Boolean, setupCaptured As Boolean, restored As Boolean, number As Long, reason As String
    If printSelectionBusy Then Err.Raise 5, , u("Đang xem trước bản in. Hãy đóng bản xem trước để tiếp tục.")
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 5, , u("Hãy quét chọn vùng ô cần in trước.")
    If target.Areas.Count <> 1 Then Err.Raise 5, , u("Chỉ chọn một vùng ô liên tục để xem trước và in.")
    Set sheet = target.Parent
    If sheet.Visible <> xlSheetVisible Then Err.Raise 5, , u("Hãy mở sheet cần in trước khi xem trước.")
    For Each view In sheet.Parent.Windows
        If view.SelectedSheets.Count > 1 Then
            For Each selected In view.SelectedSheets
                If selected.Name = sheet.Name Then Err.Raise 5, , u("Hãy chọn riêng một sheet để tránh in kèm sheet khác.")
            Next selected
        End If
    Next view
    printSelectionBusy = True
    events = Application.EnableEvents: screen = Application.ScreenUpdating
    oldStatus = Application.StatusBar: oldCancel = Application.EnableCancelKey
    wasSaved = sheet.Parent.Saved: started = True
    Application.EnableEvents = False: Application.EnableCancelKey = xlErrorHandler
    Application.StatusBar = u("Xem trước và in vùng ") & target.Address(False, False)
    With sheet.PageSetup
        area = .PrintArea: zoom = .Zoom: wide = .FitToPagesWide: tall = .FitToPagesTall
        titleRows = .PrintTitleRows: titleCols = .PrintTitleColumns: setupCaptured = True
        .PrintArea = target.Address
        .PrintTitleRows = "": .PrintTitleColumns = ""
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
    End With
    'Preview is modal: only the user's Print action sends anything to a printer.
    'Disable layout editing so temporary changes cannot leak into the workbook.
    Application.ScreenUpdating = True
    sheet.PrintPreview EnableChanges:=False
    restored = RestoreSelectionPrintSetup(sheet, area, zoom, wide, tall, titleRows, titleCols)
    If Not restored Then Err.Raise 5, , u("Không khôi phục được đầy đủ thiết lập in. Hãy kiểm tra Page Layout.")
    setupCaptured = False
    If wasSaved Then sheet.Parent.Saved = True
    Application.EnableEvents = events: Application.ScreenUpdating = screen
    Application.StatusBar = oldStatus: Application.EnableCancelKey = oldCancel: printSelectionBusy = False
    Exit Sub
Failed:
    number = Err.Number: reason = Err.Description
    On Error Resume Next
    If started Then Application.EnableCancelKey = xlDisabled
    restored = True
    If setupCaptured Then restored = RestoreSelectionPrintSetup(sheet, area, zoom, wide, tall, titleRows, titleCols)
    If started Then
        If wasSaved And restored Then sheet.Parent.Saved = True
        Application.EnableEvents = events: Application.ScreenUpdating = screen
        Application.StatusBar = oldStatus: Application.EnableCancelKey = oldCancel
    End If
    printSelectionBusy = False
    If number = 18 Then reason = u("Đã đóng hoặc dừng xem trước bản in.")
    If Not restored Then reason = reason & u(" Không khôi phục được đầy đủ thiết lập in; hãy kiểm tra Page Layout.")
    On Error GoTo 0
    Err.Raise number, "PreviewSelectedRange", reason
End Sub

Public Sub Ribbon_PrintSelection(control As IRibbonControl)
    On Error GoTo Failed
    If TypeName(Selection) <> "Range" Then Err.Raise 5, , u("Hãy quét chọn vùng ô cần in trước khi bấm nút.")
    PreviewSelectedRange Selection
    Exit Sub
Failed:
    If Err.Number <> 18 Then ShowMsg Err.Description & vbCrLf & u("Kiểm tra máy in mặc định và thiết lập in của Excel nếu không mở được bản xem trước."), u("In vùng chọn"), 48
End Sub
