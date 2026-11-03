Option Explicit
Private WithEvents saveButton As MSForms.CommandButton
Private WithEvents formulaButton As MSForms.CommandButton
Private WithEvents closeButton As MSForms.CommandButton
Private sourceBox As MSForms.ComboBox, targetBox As MSForms.ComboBox
Private WithEvents providerBox As MSForms.ComboBox
Private loading As Boolean
Private emailBox As MSForms.TextBox
Private status As MSForms.Label

Private Sub Label(ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single)
    Dim item As MSForms.Label
    Set item = Me.Controls.Add("Forms.Label.1", name, True)
    item.Caption = u(caption): item.Left = x: item.Top = y: item.Width = width: item.Height = height
    item.WordWrap = True
End Sub

Private Function Button(ByVal name As String, ByVal caption As String, ByVal x As Single, ByVal y As Single, ByVal width As Single) As MSForms.CommandButton
    Dim item As MSForms.CommandButton
    Set item = Me.Controls.Add("Forms.CommandButton.1", name, True)
    item.Caption = u(caption): item.Left = x: item.Top = y: item.Width = width: item.Height = 28
    Set Button = item
End Function

Private Function LanguageCombo(ByVal name As String, ByVal x As Single) As MSForms.ComboBox
    Dim item As MSForms.ComboBox, names As Variant, codes As Variant, i As Long
    Set item = Me.Controls.Add("Forms.ComboBox.1", name, True)
    item.Left = x: item.Top = 65: item.Width = 220: item.Height = 24
    item.Style = fmStyleDropDownList: item.ColumnCount = 2: item.BoundColumn = 2
    item.ColumnWidths = "160 pt;50 pt": item.ListRows = 12
    names = Array("Tiếng Việt", "Tiếng Anh", "Tiếng Nhật", "Tiếng Hàn", "Trung giản thể", "Trung phồn thể", "Tiếng Pháp", "Tiếng Đức", "Tiếng Thái", "Tiếng Tây Ban Nha", "Tiếng Nga", "Tiếng Indonesia")
    codes = Array("vi", "en", "ja", "ko", "zh-CN", "zh-TW", "fr", "de", "th", "es", "ru", "id")
    For i = 0 To UBound(codes)
        item.AddItem u(names(i)): item.List(i, 1) = codes(i)
    Next i
    Set LanguageCombo = item
End Function

Private Sub UserForm_Initialize()
    loading = True
    Me.Caption = "Cai dat dich - Tien ich Excel"
    Me.Width = 520: Me.Height = 586: Me.StartUpPosition = 1
    Me.Font.Name = "Segoe UI": Me.Font.Size = 10
    Label "lblTitle", "CÀI ĐẶT DỊCH NGÔN NGỮ", 16, 12, 472, 22
    Me.Controls("lblTitle").Font.Bold = True
    Label "lblSource", "1. Văn bản đang là tiếng", 16, 43, 220, 20
    Label "lblTarget", "Bạn muốn dịch sang", 268, 43, 220, 20
    Set sourceBox = LanguageCombo("cboSource", 16)
    Set targetBox = LanguageCombo("cboTarget", 268)
    sourceBox.Value = TranslationSourceLanguage()
    targetBox.Value = TranslationTargetLanguage()
    If sourceBox.ListIndex < 0 Then sourceBox.ListIndex = 0
    If targetBox.ListIndex < 0 Then targetBox.ListIndex = 1
    Label "lblProvider", "2. Chọn cách dịch", 16, 104, 472, 20
    Set providerBox = Me.Controls.Add("Forms.ComboBox.1", "cboProvider", True)
    providerBox.Left = 16: providerBox.Top = 128: providerBox.Width = 472: providerBox.Height = 24
    providerBox.Style = fmStyleDropDownList: providerBox.ColumnCount = 2: providerBox.BoundColumn = 2
    providerBox.ColumnWidths = "440 pt;0 pt"
    providerBox.AddItem u("Tự động: MyMemory hết lượt thì dùng Google web"): providerBox.List(0, 1) = "auto"
    providerBox.AddItem u("Chỉ dùng MyMemory"): providerBox.List(1, 1) = "mymemory"
    providerBox.AddItem u("Chỉ dùng Google web (không cần email)"): providerBox.List(2, 1) = "googleweb"
    providerBox.Value = TranslationProviderMode()
    Label "lblProviderHelp", "", 16, 164, 472, 44
    Label "lblEmail", "3. Email của bạn — có thể để trống", 16, 216, 472, 20
    Set emailBox = Me.Controls.Add("Forms.TextBox.1", "txtEmail", True)
    emailBox.Left = 16: emailBox.Top = 240: emailBox.Width = 472: emailBox.Height = 24
    emailBox.MaxLength = 254: emailBox.Text = TranslationEmail()
    Label "lblEmailHelp", "", 16, 276, 472, 48
    Label "lblPrivacy", "Cần Internet. Văn bản sẽ gửi tới dịch vụ đang dùng; email chỉ gửi MyMemory. Google web có thể bị giới hạn hoặc tạm ngừng.", 16, 332, 472, 36
    Label "lblHelp", "Bấm Lưu cài đặt → Đóng. Sau đó chọn các ô trong một cột và bấm Dịch ngôn ngữ. Để trống cột ngay bên phải để nhận bản dịch.", 16, 376, 472, 32
    Label "lblFormula", "Hoặc nhập công thức vào ô Excel (A2 là ô chứa văn bản):" & vbCrLf & "=Dich(A2) — dùng ngôn ngữ đã lưu" & vbCrLf & "=Dich(A2,""vi"",""en"") — Việt sang Anh" & vbCrLf & "=Dich(A2,""en"",""vi"") — Anh sang Việt", 16, 414, 472, 60
    Me.Controls("lblFormula").Font.Size = 9
    Label "lblStatus", "Lựa chọn được nhớ cho lần sau. Lưu cài đặt chưa bắt đầu dịch.", 16, 480, 472, 30
    Set status = Me.Controls("lblStatus")
    Set formulaButton = Button("btnFormula", "Cách dùng công thức", 16, 518, 160)
    Set saveButton = Button("btnSaveSettings", "Lưu cài đặt", 242, 518, 130)
    saveButton.Default = True
    Set closeButton = Button("btnClose", "Đóng", 384, 518, 104)
    closeButton.Cancel = True
    loading = False
    UpdateHelp

End Sub

Private Sub saveButton_Click()
    On Error GoTo Failed
    ValidateProviderSettings providerBox.Value
    SaveTranslationSettings sourceBox.Value, targetBox.Value, emailBox.Text
    SaveProviderSettings providerBox.Value
    emailBox.Text = TranslationEmail()
    status.Caption = u("Đã lưu. Bấm Đóng, chọn ô văn bản rồi bấm Dịch ngôn ngữ để dịch ngay.")
    Exit Sub
Failed:
    status.Caption = Err.Description
End Sub

Private Sub formulaButton_Click()
    Ribbon_TranslateHelp Nothing
End Sub

Private Sub closeButton_Click()
    Unload Me
End Sub

Private Sub providerBox_Change()
    If Not loading Then UpdateHelp
End Sub

Private Sub UpdateHelp()
    Select Case providerBox.Value
        Case "auto"
            Me.Controls("lblProviderHelp").Caption = u("Dùng MyMemory trước. Khi MyMemory báo hết hạn mức, tự chuyển Google web. Bạn không cần đăng ký hay nhập mã truy cập.")
        Case "mymemory"
            Me.Controls("lblProviderHelp").Caption = u("Chỉ dịch bằng MyMemory. Nếu hết hạn mức, add-in sẽ báo và dừng; bạn có thể đổi sang Google web.")
        Case "googleweb"
            Me.Controls("lblProviderHelp").Caption = u("Dịch bằng Google web, miễn phí và không cần email. Đây là kết nối web không chính thức, có thể bị giới hạn.")
    End Select
    emailBox.Enabled = (providerBox.Value <> "googleweb")
    If emailBox.Enabled Then
        Me.Controls("lblEmailHelp").Caption = u("MyMemory: để trống được 5.000 ký tự/ngày; email hợp lệ được 50.000 ký tự/ngày theo dịch vụ. Nhập email đang dùng của bạn. Muốn bỏ email: xóa rồi Lưu cài đặt.")
    Else
        Me.Controls("lblEmailHelp").Caption = u("Google web không dùng email. Email MyMemory đã lưu vẫn được giữ để dùng khi bạn chọn MyMemory hoặc Tự động; không gửi email này cho Google.")
    End If
End Sub
