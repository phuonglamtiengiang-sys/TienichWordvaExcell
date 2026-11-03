Option Explicit
Private WithEvents pickA As MSForms.CommandButton
Private WithEvents pickB As MSForms.CommandButton
Private WithEvents addPair As MSForms.CommandButton
Private WithEvents removePair As MSForms.CommandButton
Private WithEvents runCompare As MSForms.CommandButton
Private WithEvents closeForm As MSForms.CommandButton
Private WithEvents keyBoxA As MSForms.ComboBox
Private WithEvents keyBoxB As MSForms.ComboBox
Private WithEvents modeBox As MSForms.ComboBox
Private columnBoxA As MSForms.ComboBox, columnBoxB As MSForms.ComboBox
Private addressA As MSForms.TextBox, addressB As MSForms.TextBox, toleranceBox As MSForms.TextBox
Private pairsList As MSForms.ListBox, ignoreCaseBox As MSForms.CheckBox, trimBox As MSForms.CheckBox
Private status As MSForms.Label
Private sourceA As Range, sourceB As Range, loading As Boolean, manualMappings As Boolean

Private Function MakeButton(ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal width As Single) As MSForms.CommandButton
    Dim button As MSForms.CommandButton
    Set button = Me.Controls.Add("Forms.CommandButton.1", name, True)
    button.Caption = u(caption): button.Left = x: button.Top = y: button.Width = width: button.Height = 26
    Set MakeButton = button
End Function

Private Sub MakeLabel(ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal width As Single, Optional ByVal height As Single = 20)
    Dim label As MSForms.Label
    Set label = Me.Controls.Add("Forms.Label.1", name, True)
    label.Caption = u(caption): label.Left = x: label.Top = y: label.Width = width: label.Height = height
    label.WordWrap = True
End Sub

Private Function MakeCombo(ByVal name As String, ByVal x As Single, ByVal y As Single, ByVal width As Single) As MSForms.ComboBox
    Dim combo As MSForms.ComboBox
    Set combo = Me.Controls.Add("Forms.ComboBox.1", name, True)
    combo.Left = x: combo.Top = y: combo.Width = width: combo.Height = 22
    combo.Style = fmStyleDropDownList: combo.ListRows = 12
    Set MakeCombo = combo
End Function

Private Sub UserForm_Initialize()
    loading = True
    Me.Caption = "Compare Tables"
    Me.Width = 610: Me.Height = 526: Me.StartUpPosition = 1
    Me.Font.Name = "Segoe UI": Me.Font.Size = 9
    MakeLabel "lblIntro", "ĐỐI CHIẾU HAI BẢNG EXCEL", 16, 10, 570
    Me.Controls("lblIntro").Font.Bold = True
    MakeLabel "lblHelp", "Chọn vùng gồm một dòng tiêu đề và các dòng dữ liệu. Có thể chọn ở hai file khác nhau.", 16, 30, 570
    Set pickA = MakeButton("btnPickA", "Chọn bảng A", 16, 58, 108)
    Set pickB = MakeButton("btnPickB", "Chọn bảng B", 16, 94, 108)
    Set addressA = Me.Controls.Add("Forms.TextBox.1", "txtAddressA", True)
    addressA.Left = 136: addressA.Top = 59: addressA.Width = 448: addressA.Height = 24: addressA.Locked = True: addressA.TabStop = False
    Set addressB = Me.Controls.Add("Forms.TextBox.1", "txtAddressB", True)
    addressB.Left = 136: addressB.Top = 95: addressB.Width = 448: addressB.Height = 24: addressB.Locked = True: addressB.TabStop = False
    MakeLabel "lblKeyA", "Cột mã A", 16, 136, 76
    MakeLabel "lblKeyB", "Cột mã B", 310, 136, 76
    Set keyBoxA = MakeCombo("cboKeyA", 96, 132, 190)
    Set keyBoxB = MakeCombo("cboKeyB", 390, 132, 194)
    MakeLabel "lblPairs", "Ghép cột thông tin của bảng A với bảng B (tự ghép các tiêu đề giống nhau)", 16, 169, 570
    Set columnBoxA = MakeCombo("cboColumnA", 16, 192, 220)
    Set columnBoxB = MakeCombo("cboColumnB", 248, 192, 220)
    Set addPair = MakeButton("btnAddPair", "Thêm cặp", 480, 190, 104)
    Set pairsList = Me.Controls.Add("Forms.ListBox.1", "lstPairs", True)
    pairsList.Left = 16: pairsList.Top = 224: pairsList.Width = 568: pairsList.Height = 88
    pairsList.ColumnCount = 4: pairsList.ColumnWidths = "0 pt;0 pt;272 pt;272 pt"
    Set removePair = MakeButton("btnRemovePair", "Bỏ cặp đã chọn", 456, 318, 128)
    MakeLabel "lblPairsHint", "Danh sách trống: chỉ kiểm tra mã có/không có, chưa so sánh thông tin.", 16, 319, 425, 29
    Set ignoreCaseBox = Me.Controls.Add("Forms.CheckBox.1", "chkIgnoreCase", True)
    ignoreCaseBox.Caption = u("Bỏ qua hoa/thường"): ignoreCaseBox.Value = True
    ignoreCaseBox.Left = 16: ignoreCaseBox.Top = 355: ignoreCaseBox.Width = 250: ignoreCaseBox.Height = 22
    Set trimBox = Me.Controls.Add("Forms.CheckBox.1", "chkTrim", True)
    trimBox.Caption = u("Bỏ khoảng trắng đầu/cuối"): trimBox.Value = True
    trimBox.Left = 310: trimBox.Top = 355: trimBox.Width = 274: trimBox.Height = 22
    MakeLabel "lblTolerance", "Sai số tuyệt đối khi so sánh số", 16, 389, 220
    Set toleranceBox = Me.Controls.Add("Forms.TextBox.1", "txtTolerance", True)
    toleranceBox.Left = 240: toleranceBox.Top = 383: toleranceBox.Width = 85: toleranceBox.Height = 24: toleranceBox.Text = "0"
    Set runCompare = MakeButton("btnCompare", "Xuất báo cáo", 350, 420, 126)
    Set closeForm = MakeButton("btnClose", "Đóng", 488, 420, 96): closeForm.Cancel = True
    Set status = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    status.Left = 16: status.Top = 419: status.Width = 318: status.Height = 30: status.WordWrap = True
    status.Caption = u("Báo cáo mở trong file mới; hai bảng nguồn được giữ nguyên.")
    Dim item As Object
    For Each item In Me.Controls
        If item.Name <> "lblIntro" Then item.Top = item.Top + 44
    Next item
    MakeLabel "lblMode", "Cách đối chiếu", 16, 39, 116
    Set modeBox = MakeCombo("cboMode", 136, 35, 448)
    modeBox.AddItem u("So từng ô theo vị trí — tô màu hai bảng")
    modeBox.AddItem u("Ghép theo cột mã — xuất báo cáo riêng")
    modeBox.ListIndex = 0
    loading = False
    UpdateMode
End Sub

Private Sub UpdateMode()
    Dim name As Variant, byKey As Boolean
    byKey = (modeBox.ListIndex = 1)
    For Each name In Array("cboKeyA", "cboKeyB", "cboColumnA", "cboColumnB", "btnAddPair", "btnRemovePair", "lstPairs", "lblKeyA", "lblKeyB", "lblPairs", "lblPairsHint")
        Me.Controls(CStr(name)).Enabled = byKey
        Me.Controls(CStr(name)).Visible = byKey
    Next name
    Me.Controls("lblHelp").Height = 24
    If byKey Then
        Me.Height = 526
        ignoreCaseBox.Top = 399: trimBox.Top = 399
        Me.Controls("lblTolerance").Top = 433: toleranceBox.Top = 427
        status.Top = 463: status.Height = 30: runCompare.Top = 464: closeForm.Top = 464
        runCompare.Caption = u("Xuất báo cáo")
        Me.Controls("lblHelp").Caption = u("Chọn vùng có dòng tiêu đề. Bấm Chọn bảng A/B, chọn sheet rồi quét khối dữ liệu.")
        status.Caption = u("Báo cáo mở trong file mới; giữ nguyên hai bảng nguồn.")
    Else
        Me.Height = 314
        ignoreCaseBox.Top = 182: trimBox.Top = 182
        Me.Controls("lblTolerance").Top = 216: toleranceBox.Top = 210
        status.Top = 244: status.Height = 40: runCompare.Top = 246: closeForm.Top = 246
        runCompare.Caption = u("Tô màu ô khác")
        Me.Controls("lblHelp").Caption = u("Bấm Chọn bảng A/B, bấm tab sheet rồi quét khối. So ô góc trên trái A với ô góc trên trái B; không cần cột mã.")
        status.Caption = u("Ô giống nhau: bỏ màu nền. Vàng: ô khác. Cam: phần dư. Hoàn tác để khôi phục màu cũ.")
    End If
End Sub

Private Sub modeBox_Change()
    If loading Then Exit Sub
    On Error GoTo NeedHeaders
    UpdateMode
    If modeBox.ListIndex = 1 Then
        If Not sourceA Is Nothing Then LoadTable "A", sourceA
        If Not sourceB Is Nothing Then LoadTable "B", sourceB
        If modeBox.ListIndex = 1 Then AutoPairs
    End If
    Exit Sub
NeedHeaders:
    loading = False
    status.Caption = u("Chế độ theo mã cần bảng có tiêu đề và dòng dữ liệu. Hãy chọn lại bảng A/B.")
End Sub

Private Sub ChooseTable(ByVal side As String)
    Dim chosen As Range, reason As String
    Me.Hide
    On Error Resume Next
    If modeBox.ListIndex = 0 Then
        Set chosen = Application.InputBox(u("Chọn sheet rồi quét khối dữ liệu bảng ") & side & u(". Có thể chọn sheet hoặc file khác."), "Compare Tables", Type:=8)
    Else
        Set chosen = Application.InputBox(u("Chọn cả dòng tiêu đề và các dòng dữ liệu của bảng ") & side, "Compare Tables", Type:=8)
    End If
    Err.Clear
    On Error GoTo Failed
    If Not chosen Is Nothing Then
        LoadTable side, chosen
        If modeBox.ListIndex = 1 Then AutoPairs
    End If
    Me.Show vbModeless
    Exit Sub
Failed:
    reason = Err.Description: loading = False
    On Error GoTo 0
    Me.Show vbModeless
    ShowMsg reason, u("Chọn bảng"), 48
End Sub

Private Sub LoadTable(ByVal side As String, ByVal chosen As Range)
    Dim j As Long, name As String
    If modeBox.ListIndex = 0 Then
        ValidateHighlightRange chosen
        If side = "A" Then
            Set sourceA = chosen
            addressA.Text = chosen.Address(External:=True): addressA.ControlTipText = addressA.Text
        Else
            Set sourceB = chosen
            addressB.Text = chosen.Address(External:=True): addressB.ControlTipText = addressB.Text
        End If
        Exit Sub
    End If
    ValidateCompareRange chosen
    For j = 1 To chosen.Columns.Count
        If IsError(chosen.Cells(1, j).Value2) Then Err.Raise 5, , u("Tiêu đề có ô lỗi.")
    Next j
    loading = True
    If side = "A" Then
        Set sourceA = chosen
        addressA.Text = chosen.Address(External:=True): addressA.ControlTipText = addressA.Text
        keyBoxA.Clear: columnBoxA.Clear
    Else
        Set sourceB = chosen
        addressB.Text = chosen.Address(External:=True): addressB.ControlTipText = addressB.Text
        keyBoxB.Clear: columnBoxB.Clear
    End If
    For j = 1 To chosen.Columns.Count
        name = Trim$(CStr(chosen.Cells(1, j).Value2))
        If name = "" Then name = u("Cột ") & j
        name = j & ". " & name
        If side = "A" Then
            keyBoxA.AddItem name: columnBoxA.AddItem name
        Else
            keyBoxB.AddItem name: columnBoxB.AddItem name
        End If
    Next j
    If side = "A" Then keyBoxA.ListIndex = 0 Else keyBoxB.ListIndex = 0
    loading = False
End Sub

Public Sub ConfigureTables(ByVal a As Range, ByVal b As Range, Optional ByVal aKey As Long = 1, Optional ByVal bKey As Long = 1)
    modeBox.ListIndex = 1
    LoadTable "A", a
    LoadTable "B", b
    loading = True
    keyBoxA.ListIndex = aKey - 1: keyBoxB.ListIndex = bKey - 1
    loading = False
    AutoPairs
End Sub

Public Sub ConfigureCellRanges(ByVal a As Range, ByVal b As Range)
    modeBox.ListIndex = 0
    LoadTable "A", a
    LoadTable "B", b
End Sub

Private Sub AutoPairs()
    Dim spec As String, entry As Variant, pair As Variant
    If sourceA Is Nothing Or sourceB Is Nothing Then Exit Sub
    If keyBoxA.ListIndex < 0 Or keyBoxB.ListIndex < 0 Then Exit Sub
    spec = CompareAutoPairs(sourceA, sourceB, keyBoxA.ListIndex + 1, keyBoxB.ListIndex + 1)
    pairsList.Clear
    manualMappings = False
    If spec <> "" Then
        For Each entry In Split(spec, ",")
            pair = Split(CStr(entry), ":")
            StorePair CLng(pair(0)), CLng(pair(1))
        Next entry
    End If
End Sub

Private Sub StorePair(ByVal a As Long, ByVal b As Long)
    Dim i As Long
    If a = keyBoxA.ListIndex + 1 Or b = keyBoxB.ListIndex + 1 Then Err.Raise 5, , u("Chọn cột thông tin, không chọn cột mã.")
    For i = 0 To pairsList.ListCount - 1
        If CLng(pairsList.List(i, 0)) = a Or CLng(pairsList.List(i, 1)) = b Then Err.Raise 5, , u("Cột này đã có trong một cặp. Bỏ cặp cũ trước khi ghép lại.")
    Next i
    pairsList.AddItem CStr(a)
    i = pairsList.ListCount - 1
    pairsList.List(i, 1) = CStr(b)
    pairsList.List(i, 2) = columnBoxA.List(a - 1)
    pairsList.List(i, 3) = columnBoxB.List(b - 1)
End Sub

Private Sub RemoveKeyPairs()
    Dim i As Long
    If loading Then Exit Sub
    If pairsList Is Nothing Then Exit Sub
    If Not manualMappings Then
        AutoPairs
        Exit Sub
    End If
    For i = pairsList.ListCount - 1 To 0 Step -1
        If CLng(pairsList.List(i, 0)) = keyBoxA.ListIndex + 1 Or CLng(pairsList.List(i, 1)) = keyBoxB.ListIndex + 1 Then pairsList.RemoveItem i
    Next i
End Sub

Private Sub pickA_Click()
    ChooseTable "A"
End Sub
Private Sub pickB_Click()
    ChooseTable "B"
End Sub
Private Sub keyBoxA_Change()
    RemoveKeyPairs
End Sub
Private Sub keyBoxB_Change()
    RemoveKeyPairs
End Sub
Private Sub addPair_Click()
    On Error GoTo Failed
    If columnBoxA.ListIndex < 0 Or columnBoxB.ListIndex < 0 Then Err.Raise 5, , u("Chọn một cột của mỗi bảng trước khi thêm cặp.")
    StorePair columnBoxA.ListIndex + 1, columnBoxB.ListIndex + 1
    manualMappings = True
    Exit Sub
Failed:
    ShowMsg Err.Description, u("Ghép cột"), 48
End Sub
Private Sub removePair_Click()
    If pairsList.ListIndex >= 0 Then
        pairsList.RemoveItem pairsList.ListIndex
        manualMappings = True
    End If
End Sub
Private Sub closeForm_Click()
    Unload Me
End Sub

Private Sub runCompare_Click()
    Dim report As Workbook, spec As String, i As Long, tolerance As Double, reason As String, painted As Long
    On Error GoTo Failed
    If sourceA Is Nothing Or sourceB Is Nothing Then Err.Raise 5, , u("Chọn đủ bảng A và bảng B.")
    If Not IsNumeric(toleranceBox.Text) Then Err.Raise 5, , u("Sai số phải là một số không âm.")
    tolerance = CDbl(toleranceBox.Text)
    If tolerance < 0 Then Err.Raise 5, , u("Sai số không được âm.")
    If modeBox.ListIndex = 0 Then
        runCompare.Enabled = False
        painted = HighlightCompareRanges(sourceA, sourceB, CBool(ignoreCaseBox.Value), CBool(trimBox.Value), tolerance)
        status.Caption = u("Đã đánh dấu ") & painted & u(" ô trên hai bảng. Vàng: khác; cam: phần dư. Hoàn tác để gỡ màu.")
        If painted = 0 Then status.Caption = u("Hai vùng giống nhau theo tùy chọn đã chọn. Đã bỏ màu nền trong vùng; có thể Hoàn tác.")
        runCompare.Enabled = True
        Exit Sub
    End If
    If keyBoxA.ListIndex < 0 Or keyBoxB.ListIndex < 0 Then Err.Raise 5, , u("Chọn cột mã cho cả hai bảng.")
    For i = 0 To pairsList.ListCount - 1
        If spec <> "" Then spec = spec & ","
        spec = spec & pairsList.List(i, 0) & ":" & pairsList.List(i, 1)
    Next i
    If spec = "" Then spec = "-"
    runCompare.Enabled = False
    Me.Hide
    Set report = CompareTables(sourceA, sourceB, keyBoxA.ListIndex + 1, keyBoxB.ListIndex + 1, spec, CBool(ignoreCaseBox.Value), CBool(trimBox.Value), tolerance)
    report.Activate
    Unload Me
    Exit Sub
Failed:
    reason = Err.Description
    On Error GoTo 0
    runCompare.Enabled = True
    Me.Show vbModeless
    ShowMsg reason, u("Đối chiếu hai bảng"), 48
End Sub
