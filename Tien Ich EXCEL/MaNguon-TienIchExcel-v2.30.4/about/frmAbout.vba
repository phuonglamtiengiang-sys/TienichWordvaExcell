Option Explicit
Private WithEvents closeButton As MSForms.CommandButton
Private WithEvents zoomButton As MSForms.CommandButton
Private qrFrame As MSForms.Frame
Private qrPicture As MSForms.Image
Private titleLabel As MSForms.Label
Private versionLabel As MSForms.Label
Private authorLabel As MSForms.Label
Private emailLabel As MSForms.Label
Private coffeeLabel As MSForms.Label
Private thanksLabel As MSForms.Label
Private referenceLabel As MSForms.Label
Private brandPicture As MSForms.Image
Private momoLabel As MSForms.Label
Private qrFallbackLabel As MSForms.Label
Private enlarged As Boolean
Private qrAvailable As Boolean
Private Const ABOUT_SCALE As Single = 0.75

Private Function S(ByVal value As Single) As Single
    S = value * ABOUT_SCALE
End Function

Private Function AddLabel(ByVal controlName As String, ByVal text As String, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, Optional ByVal fontSize As Single = 11) As MSForms.Label
    Dim item As MSForms.Label
    Set item = Me.Controls.Add("Forms.Label.1", controlName, True)
    item.Caption = u(text): item.Left = S(x): item.Top = S(y)
    item.Width = S(width): item.Height = S(height): item.Font.Size = S(fontSize)
    item.WordWrap = True: item.BackStyle = fmBackStyleTransparent
    item.ForeColor = RGB(42, 47, 55)
    Set AddLabel = item
End Function

Private Sub UserForm_Initialize()
    On Error GoTo InitFailed
    Me.Caption = "Gioi thieu - Tien ich Excel"
    Me.Width = S(720): Me.Height = S(600): Me.StartUpPosition = 1
    Me.BackColor = RGB(250, 250, 252)
    Me.Font.Name = "Segoe UI": Me.Font.Size = S(10)

    Set titleLabel = AddLabel("lblTitle", "TI\u1ec6N \u00cdCH EXCEL VI\u1ec6T NAM", 22, 20, 650, 28, 18)
    titleLabel.Visible = False
    Set brandPicture = Me.Controls.Add("Forms.Image.1", "imgBrandLogo", True)
    brandPicture.PictureSizeMode = fmPictureSizeModeZoom
    brandPicture.BorderStyle = fmBorderStyleNone
    brandPicture.Left = S(18): brandPicture.Top = S(7): brandPicture.Width = S(684): brandPicture.Height = S(90)
    On Error Resume Next
    Set brandPicture.Picture = LoadPicture(ThisWorkbook.Path & Application.PathSeparator & "Assets" & Application.PathSeparator & "TienIchExcel-AboutHeader.bmp")
    If Err.Number <> 0 Then brandPicture.Visible = False
    Err.Clear
    On Error GoTo InitFailed
    Set versionLabel = AddLabel("lblVersion", "Phi\u00ean b\u1ea3n 2.30.4 \u2014 Microsoft Excel Windows", 22, 105, 310, 28)
    Set authorLabel = AddLabel("lblAuthor", "T\u00e1c gi\u1ea3: H\u1ed3 Ph\u01b0\u01a1ng L\u00e2m", 22, 139, 310, 25, 12)
    Set emailLabel = AddLabel("lblEmail", "Email: hophuonglam@outlook.com", 22, 169, 310, 25, 10)
    Set coffeeLabel = AddLabel("lblCoffee", "N\u1ebfu b\u1ea1n d\u00f9ng Add-in n\u00e0y, xin g\u1eedi t\u1eb7ng t\u00f4i 1 ly c\u00e0 ph\u00ea \u0111\u1ec3 t\u00f4i c\u00f3 \u0111\u1ed9ng l\u1ef1c ph\u00e1t tri\u1ec3n ti\u1ebfp.", 22, 234, 290, 116, 14)
    coffeeLabel.ForeColor = RGB(159, 31, 109)
    Set thanksLabel = AddLabel("lblThanks", "C\u1ea3m \u01a1n b\u1ea1n \u0111\u00e3 s\u1eed d\u1ee5ng v\u00e0 \u1ee7ng h\u1ed9!", 22, 365, 290, 35, 11)
    Set referenceLabel = AddLabel("lblReferences", "Ngu\u1ed3n tham kh\u1ea3o: VietQR, BWIP-JS, QRServer, MyMemory, Google Translate, Vietcombank, NSO, MaSoThue.com v\u00e0 Thongtincongty.vn. Xem THAM-KHAO.txt trong g\u00f3i m\u00e3 ngu\u1ed3n.", 22, 410, 290, 82, 8)
    Set momoLabel = AddLabel("lblMomo", "MoMo \u2022 HO PHUONG LAM", 340, 105, 360, 24, 11)
    momoLabel.TextAlign = fmTextAlignCenter
    momoLabel.Visible = False

    Set qrFrame = Me.Controls.Add("Forms.Frame.1", "fraSupportQR", True)
    qrFrame.Left = S(340): qrFrame.Top = S(105): qrFrame.Width = S(360): qrFrame.Height = S(395)
    qrFrame.Caption = "": qrFrame.BorderStyle = fmBorderStyleNone
    qrFrame.SpecialEffect = fmSpecialEffectFlat: qrFrame.BackColor = RGB(255, 237, 246)
    Set qrPicture = qrFrame.Controls.Add("Forms.Image.1", "imgSupportQR", True)
    qrPicture.PictureSizeMode = fmPictureSizeModeZoom
    qrPicture.BorderStyle = fmBorderStyleNone
    qrPicture.Width = S(360): qrPicture.Height = S(395)
    qrPicture.Left = 0: qrPicture.Top = 0
    qrPicture.ControlTipText = u("M\u00e3 QR MoMo c\u1ee7a H\u1ed3 Ph\u01b0\u01a1ng L\u00e2m. B\u1ea5m Ph\u00f3ng to \u0111\u1ec3 xem r\u00f5 h\u01a1n.")
    On Error Resume Next
    Set qrPicture.Picture = LoadPicture(ThisWorkbook.Path & Application.PathSeparator & "Assets" & Application.PathSeparator & "support-momo.bmp")
    qrAvailable = (Err.Number = 0)
    Err.Clear
    On Error GoTo InitFailed

    Set zoomButton = Me.Controls.Add("Forms.CommandButton.1", "btnZoom", True)
    zoomButton.Caption = u("Ph\u00f3ng to m\u00e3 QR")
    zoomButton.Left = S(340): zoomButton.Top = S(525): zoomButton.Width = S(175): zoomButton.Height = S(30)
    Set closeButton = Me.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    closeButton.Caption = u("\u0110\u00f3ng")
    closeButton.Left = S(530): closeButton.Top = S(525): closeButton.Width = S(170): closeButton.Height = S(30)
    closeButton.Cancel = True

    If Not qrAvailable Then
        Set qrFallbackLabel = AddLabel("lblQRUnavailable", "M\u00e3 QR ch\u01b0a t\u1ea3i \u0111\u01b0\u1ee3c. B\u1ea1n v\u1eabn c\u00f3 th\u1ec3 xem th\u00f4ng tin add-in.", 350, 240, 330, 60, 10)
        qrPicture.Visible = False
        zoomButton.Caption = u("Kh\u00f4ng c\u00f3 m\u00e3 QR")
        zoomButton.Enabled = False
    End If
    Exit Sub

InitFailed:
    MsgBox u("Kh\u00f4ng th\u1ec3 t\u1ea3i \u0111\u1ea7y \u0111\u1ee7 n\u1ed9i dung Gi\u1edbi thi\u1ec7u: ") & Err.Description, vbInformation, u("Ti\u1ec7n \u00edch Excel")
End Sub

Private Sub UserForm_Activate()
    Me.Caption = "Gioi thieu - Tien ich Excel"
End Sub

Private Sub zoomButton_Click()
    On Error GoTo SafeExit
    If Not qrAvailable Then Exit Sub
    enlarged = Not enlarged
    If enlarged Then
        Me.Width = S(720): Me.Height = S(850)
        authorLabel.Visible = False: emailLabel.Visible = False
        coffeeLabel.Visible = False: thanksLabel.Visible = False
        brandPicture.Visible = False: referenceLabel.Visible = False: versionLabel.Visible = False
        momoLabel.Left = S(80): momoLabel.Top = S(40): momoLabel.Width = S(560)
        qrFrame.Left = S(80): qrFrame.Top = S(70): qrFrame.Width = S(560): qrFrame.Height = S(612)
        qrPicture.Width = S(560): qrPicture.Height = S(612)
        qrPicture.Left = 0: qrPicture.Top = 0
        zoomButton.Caption = u("Thu nh\u1ecf")
        zoomButton.Left = S(80): zoomButton.Top = S(710): zoomButton.Width = S(250)
        closeButton.Left = S(390): closeButton.Top = S(710): closeButton.Width = S(250)
    Else
        Me.Width = S(720): Me.Height = S(600)
        authorLabel.Visible = True: emailLabel.Visible = True
        coffeeLabel.Visible = True: thanksLabel.Visible = True
        brandPicture.Visible = True: referenceLabel.Visible = True: versionLabel.Visible = True
        momoLabel.Left = S(340): momoLabel.Top = S(105): momoLabel.Width = S(360)
        qrFrame.Left = S(340): qrFrame.Top = S(105): qrFrame.Width = S(360): qrFrame.Height = S(395)
        qrPicture.Width = S(360): qrPicture.Height = S(395)
        qrPicture.Left = 0: qrPicture.Top = 0
        zoomButton.Caption = u("Ph\u00f3ng to m\u00e3 QR")
        zoomButton.Left = S(340): zoomButton.Top = S(525): zoomButton.Width = S(175)
        closeButton.Left = S(530): closeButton.Top = S(525): closeButton.Width = S(170)
    End If
SafeExit:
End Sub

Private Sub closeButton_Click()
    Unload Me
End Sub
