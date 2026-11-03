Option Explicit
Private WithEvents ratesList As MSForms.ListBox
Private WithEvents refreshButton As MSForms.CommandButton
Private WithEvents copyAllButton As MSForms.CommandButton
Private WithEvents copyRowsButton As MSForms.CommandButton
Private WithEvents closeButton As MSForms.CommandButton
Private includeSourceBox As MSForms.CheckBox
Private timeLabel As MSForms.Label, statusLabel As MSForms.Label
Private table As Object
Private activated As Boolean, busy As Boolean

Private Function AddLabel(ByVal name As String, ByVal text As String, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As MSForms.Label
    Dim item As MSForms.Label
    Set item = Me.Controls.Add("Forms.Label.1", name, True)
    item.Caption = u(text): item.Left = x: item.Top = y: item.Width = width: item.Height = height: item.WordWrap = True
    Set AddLabel = item
End Function

Private Function Button(ByVal name As String, ByVal text As String, ByVal x As Single, ByVal width As Single) As MSForms.CommandButton
    Dim item As MSForms.CommandButton
    Set item = Me.Controls.Add("Forms.CommandButton.1", name, True)
    item.Caption = u(text): item.Left = x: item.Top = 480: item.Width = width: item.Height = 30
    Set Button = item
End Function

Private Sub UserForm_Initialize()
    Dim item As MSForms.Label
    Me.Caption = "Xem ty gia - Tien ich Excel"
    Me.Width = 740: Me.Height = 550: Me.StartUpPosition = 1
    Me.Font.Name = "Segoe UI": Me.Font.Size = 10
    Set item = AddLabel("lblTitle", "BẢNG TỶ GIÁ VIETCOMBANK", 16, 12, 692, 24)
    item.Font.Bold = True
    Set item = AddLabel("lblSource", "Nguồn: Vietcombank | " & FX_BANK_URL, 16, 43, 692, 22)
    Set timeLabel = AddLabel("lblTime", "Đang chuẩn bị tải bảng tỷ giá...", 16, 72, 692, 36)
    Set item = AddLabel("lblCode", "Mã", 20, 113, 48, 20): item.Font.Bold = True
    Set item = AddLabel("lblName", "Tên ngoại tệ", 75, 113, 220, 20): item.Font.Bold = True
    Set item = AddLabel("lblCash", "Mua tiền mặt", 305, 113, 115, 20): item.Font.Bold = True
    Set item = AddLabel("lblTransfer", "Mua chuyển khoản", 434, 113, 125, 20): item.Font.Bold = True
    Set item = AddLabel("lblSell", "Bán", 563, 113, 125, 20): item.Font.Bold = True
    Set ratesList = Me.Controls.Add("Forms.ListBox.1", "lstRates", True)
    ratesList.Left = 16: ratesList.Top = 135: ratesList.Width = 692: ratesList.Height = 250
    ratesList.ColumnCount = 5: ratesList.ColumnWidths = "55 pt;230 pt;129 pt;129 pt;129 pt"
    ratesList.MultiSelect = fmMultiSelectExtended
    Set item = AddLabel("lblHelp", "Chọn dòng để chép; giữ Ctrl/Shift để chọn nhiều dòng. Ctrl+C chép dòng chọn. Dán vào Excel bằng Ctrl+V.", 16, 391, 692, 22)
    Set includeSourceBox = Me.Controls.Add("Forms.CheckBox.1", "chkSource", True)
    includeSourceBox.Left = 16: includeSourceBox.Top = 417: includeSourceBox.Width = 692: includeSourceBox.Height = 20
    includeSourceBox.Caption = u("Kèm nguồn và thời điểm khi sao chép (bỏ chọn nếu chỉ cần bảng tỷ giá)")
    includeSourceBox.Value = True
    Set statusLabel = AddLabel("lblStatus", "Mở form sẽ tự tải. Tối đa một lần tải mỗi 5 phút; trong thời gian này dùng bảng đã tải.", 16, 444, 692, 30)
    Set refreshButton = Button("btnRefresh", "Cập nhật", 16, 110)
    Set copyAllButton = Button("btnCopyAll", "Sao chép toàn bảng", 136, 180)
    Set copyRowsButton = Button("btnCopyRows", "Sao chép dòng chọn", 326, 232)
    Set closeButton = Button("btnClose", "Đóng", 568, 140)
    copyAllButton.Enabled = False: copyRowsButton.Enabled = False
End Sub

Private Sub UserForm_Activate()
    If activated Then Exit Sub
    activated = True
    Me.Repaint
    LoadRates
End Sub

Public Sub ShowRates(ByVal data As Object)
    Dim values As Variant
    Set table = Nothing: ratesList.Clear: copyAllButton.Enabled = False: copyRowsButton.Enabled = False
    values = FXViewerRows(data)
    ratesList.List = values
    Set table = data
    timeLabel.Caption = u("Công bố: ") & Format$(data("Updated"), "dd/mm/yyyy hh:nn:ss") & u(" | Tải về: ") & Format$(data("Retrieved"), "dd/mm/yyyy hh:nn:ss") & vbCrLf & u("Đơn vị: VND cho 1 đơn vị ngoại tệ. Tỷ giá tham khảo; dấu '-' là không có giá.")
    copyAllButton.Enabled = True
    statusLabel.Caption = ratesList.ListCount & u(" ngoại tệ. Sao chép sẽ tách thành 5 cột khi dán vào Excel.")
    If DateDiff("d", data("Updated"), Now) >= 3 Then statusLabel.Caption = statusLabel.Caption & u(" Lưu ý: ngày công bố đã từ 3 ngày trước trở lên.")
End Sub

Private Sub LoadRates()
    Dim data As Object
    If busy Then Exit Sub
    On Error GoTo Failed
    busy = True: refreshButton.Enabled = False: closeButton.Enabled = False
    Set table = Nothing: ratesList.Clear: copyAllButton.Enabled = False: copyRowsButton.Enabled = False
    statusLabel.Caption = u("Đang tải tỷ giá từ Vietcombank..."): Me.Repaint
    Set data = FXGetRates()
    ShowRates data
Finished:
    refreshButton.Enabled = True: closeButton.Enabled = True: busy = False
    Exit Sub
Failed:
    timeLabel.Caption = u("Chưa tải được tỷ giá. Không có bảng giá để sao chép.")
    statusLabel.Caption = Err.Description
    Resume Finished
End Sub

Private Sub refreshButton_Click()
    LoadRates
End Sub

Private Sub ratesList_Change()
    Dim i As Long, hasSelection As Boolean
    If copyRowsButton Is Nothing Then Exit Sub
    For i = 0 To ratesList.ListCount - 1
        If ratesList.Selected(i) Then hasSelection = True: Exit For
    Next i
    copyRowsButton.Enabled = hasSelection And (Not table Is Nothing)
End Sub

Private Sub ratesList_Click()
    ratesList_Change
End Sub

Private Sub ratesList_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ratesList_Change
End Sub

Private Sub ratesList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyC And (Shift And 2) <> 0 Then
        If copyRowsButton.Enabled Then copyRowsButton_Click
        KeyCode = 0
    End If
End Sub

Private Sub copyAllButton_Click()
    On Error GoTo Failed
    If table Is Nothing Then Exit Sub
    FXCopyToClipboard FXRatesText(table, "", CBool(includeSourceBox.Value))
    statusLabel.Caption = u("Đã sao chép toàn bảng. Đóng form, chọn ô đích trong Excel rồi nhấn Ctrl+V.")
    Exit Sub
Failed:
    statusLabel.Caption = Err.Description
End Sub

Private Sub copyRowsButton_Click()
    Dim codes As String, i As Long, count As Long
    On Error GoTo Failed
    If table Is Nothing Then Exit Sub
    For i = 0 To ratesList.ListCount - 1
        If ratesList.Selected(i) Then
            If codes <> "" Then codes = codes & "|"
            codes = codes & CStr(ratesList.List(i, 0)): count = count + 1
        End If
    Next i
    If count = 0 Then Err.Raise 5, , u("Hãy chọn ít nhất một dòng tỷ giá.")
    FXCopyToClipboard FXRatesText(table, codes, CBool(includeSourceBox.Value))
    statusLabel.Caption = u("Đã sao chép ") & count & u(" dòng. Đóng form và nhấn Ctrl+V tại ô muốn dán.")
    Exit Sub
Failed:
    statusLabel.Caption = Err.Description
End Sub

Private Sub closeButton_Click()
    Unload Me
End Sub
