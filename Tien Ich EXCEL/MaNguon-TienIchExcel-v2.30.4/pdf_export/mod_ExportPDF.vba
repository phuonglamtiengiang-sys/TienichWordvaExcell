Option Explicit
Private pdfBusy As Boolean

Private Function RestorePDFSetup(ByVal sheet As Worksheet, ByVal area As String, ByVal zoom As Variant, ByVal wide As Variant, ByVal tall As Variant, ByVal titleRows As String, ByVal titleCols As String) As Boolean
    On Error Resume Next
    Err.Clear
    With sheet.PageSetup
        .PrintArea = area
        .FitToPagesWide = wide: .FitToPagesTall = tall: .Zoom = zoom
        .PrintTitleRows = titleRows: .PrintTitleColumns = titleCols
    End With
    RestorePDFSetup = (Err.Number = 0)
End Function

Public Function ExportSelectedRangePDF(ByVal target As Range, ByVal filePath As String, Optional ByVal overwrite As Boolean = False, Optional ByVal openAfter As Boolean = False) As Boolean
    Dim sheet As Worksheet, fileSystem As Object, parentPath As String, temporary As String, view As Window, selected As Object
    Dim area As String, zoom As Variant, wide As Variant, tall As Variant, titleRows As String, titleCols As String
    Dim oldStatus As Variant, events As Boolean, screen As Boolean, oldCancel As Long, wasSaved As Boolean
    Dim started As Boolean, setupCaptured As Boolean, restored As Boolean, number As Long, reason As String
    If pdfBusy Then Err.Raise 5, , u("Đang xuất PDF. Hãy chờ hoàn tất.")
    On Error GoTo Failed
    If target Is Nothing Then Err.Raise 5, , u("Hãy chọn vùng ô cần xuất PDF.")
    If target.Areas.Count <> 1 Then Err.Raise 5, , u("Hãy chọn một vùng ô liên tục để xuất PDF.")
    Set sheet = target.Parent
    For Each view In sheet.Parent.Windows
        If view.SelectedSheets.Count > 1 Then
            For Each selected In view.SelectedSheets
                If selected.Name = sheet.Name Then Err.Raise 5, , u("Đang chọn nhóm nhiều sheet. Hãy chọn riêng một sheet rồi xuất PDF để tránh xuất kèm sheet khác.")
            Next selected
        End If
    Next view
    If LCase$(Right$(filePath, 4)) <> ".pdf" Then Err.Raise 5, , u("Tên tệp phải có đuôi .pdf.")
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    filePath = fileSystem.GetAbsolutePathName(filePath)
    parentPath = fileSystem.GetParentFolderName(filePath)
    If Not fileSystem.FolderExists(parentPath) Then Err.Raise 76, , u("Thư mục lưu không tồn tại.")
    If fileSystem.FileExists(filePath) And Not overwrite Then Err.Raise 58, , u("Tệp PDF đã tồn tại. Chọn tên khác hoặc xác nhận ghi đè.")
    pdfBusy = True
    events = Application.EnableEvents: screen = Application.ScreenUpdating
    oldStatus = Application.StatusBar: oldCancel = Application.EnableCancelKey
    wasSaved = sheet.Parent.Saved: started = True
    Application.EnableEvents = False: Application.ScreenUpdating = False: Application.EnableCancelKey = xlErrorHandler
    Application.StatusBar = u("Đang xuất PDF vùng ") & target.Address(False, False) & "..."
    With sheet.PageSetup
        area = .PrintArea: zoom = .Zoom: wide = .FitToPagesWide: tall = .FitToPagesTall
        titleRows = .PrintTitleRows: titleCols = .PrintTitleColumns: setupCaptured = True
        .PrintArea = target.Address
        .PrintTitleRows = "": .PrintTitleColumns = ""
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
    End With
    temporary = fileSystem.BuildPath(parentPath, fileSystem.GetTempName & ".pdf")
    sheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=temporary, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Not fileSystem.FileExists(temporary) Then Err.Raise 5, , u("Excel chưa tạo được tệp PDF.")
    restored = RestorePDFSetup(sheet, area, zoom, wide, tall, titleRows, titleCols)
    If Not restored Then Err.Raise 5, , u("Không khôi phục được đầy đủ thiết lập in. Hãy kiểm tra Page Layout.")
    setupCaptured = False
    'Recheck before committing so an existing file is never silently replaced.
    If fileSystem.FileExists(filePath) And Not overwrite Then Err.Raise 58, , u("Tệp PDF đã tồn tại. Hãy chọn tên khác.")
    fileSystem.CopyFile temporary, filePath, overwrite
    fileSystem.DeleteFile temporary: temporary = ""
    If wasSaved Then sheet.Parent.Saved = True
    Application.EnableEvents = events: Application.ScreenUpdating = screen
    Application.StatusBar = oldStatus: Application.EnableCancelKey = oldCancel: pdfBusy = False
    ExportSelectedRangePDF = True
    If openAfter Then
        On Error Resume Next
        sheet.Parent.FollowHyperlink Address:=filePath
        On Error GoTo 0
    End If
    Exit Function
Failed:
    number = Err.Number: reason = Err.Description
    On Error Resume Next
    If started Then Application.EnableCancelKey = xlDisabled
    restored = True
    If setupCaptured Then restored = RestorePDFSetup(sheet, area, zoom, wide, tall, titleRows, titleCols)
    If Len(temporary) > 0 Then
        If fileSystem.FileExists(temporary) Then fileSystem.DeleteFile temporary
    End If
    If started Then
        If wasSaved And restored Then sheet.Parent.Saved = True
        Application.EnableEvents = events: Application.ScreenUpdating = screen
        Application.StatusBar = oldStatus: Application.EnableCancelKey = oldCancel
    End If
    pdfBusy = False
    If number = 18 Then reason = u("Đã dừng xuất PDF.")
    If Not restored Then reason = reason & u(" Không khôi phục được đầy đủ thiết lập in; hãy kiểm tra Page Layout.")
    On Error GoTo 0
    Err.Raise number, "ExportSelectedRangePDF", reason
End Function

Public Sub Ribbon_ExportSelectionPDF(control As IRibbonControl)
    Dim target As Range, chosen As Variant, folder As String, name As String, bad As Variant, overwrite As Boolean
    On Error GoTo Failed
    If TypeName(Selection) <> "Range" Then Err.Raise 5, , u("Hãy quét chọn vùng ô cần xuất PDF trước khi bấm nút.")
    Set target = Selection
    If target.Areas.Count <> 1 Then Err.Raise 5, , u("Hãy chọn một vùng ô liên tục.")
    folder = target.Parent.Parent.Path
    If folder = "" Or InStr(folder, "://") > 0 Then folder = Application.DefaultFilePath
    name = target.Parent.Name & "_" & Format$(Now, "yyyymmdd_hhnnss")
    For Each bad In Array(":", "\", "/", "?", "*", "[", "]", """", "<", ">", "|")
        name = Replace(name, CStr(bad), "_")
    Next bad
    chosen = Application.GetSaveAsFilename(InitialFileName:=folder & Application.PathSeparator & name & ".pdf", FileFilter:="PDF (*.pdf), *.pdf", Title:=u("Lưu PDF vùng đã chọn"))
    If VarType(chosen) = vbBoolean Then Exit Sub
    If LCase$(Right$(CStr(chosen), 4)) <> ".pdf" Then chosen = CStr(chosen) & ".pdf"
    If CreateObject("Scripting.FileSystemObject").FileExists(CStr(chosen)) Then
        If MsgBox(u("Tệp đã tồn tại. Bạn có muốn ghi đè?") & vbCrLf & CStr(chosen), vbYesNo Or vbQuestion Or vbDefaultButton2, u("Xuất PDF vùng chọn")) <> vbYes Then Exit Sub
        overwrite = True
    End If
    ExportSelectedRangePDF target, CStr(chosen), overwrite, True
    Exit Sub
Failed:
    ShowMsg Err.Description & vbCrLf & u("Nếu không xuất được, hãy kiểm tra quyền lưu, tệp PDF đang mở và thiết lập máy in của Excel."), u("Xuất PDF vùng chọn"), 48
End Sub
