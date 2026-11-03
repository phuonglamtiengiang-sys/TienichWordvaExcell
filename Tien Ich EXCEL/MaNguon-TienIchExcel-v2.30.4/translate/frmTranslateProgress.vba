Option Explicit
Private WithEvents actionButton As MSForms.CommandButton
Private status As MSForms.Label
Private running As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "Dich ngon ngu - Tien ich Excel"
    Me.Width = 480: Me.Height = 190: Me.StartUpPosition = 1
    Me.Font.Name = "Segoe UI": Me.Font.Size = 10
    Set status = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    status.Left = 16: status.Top = 16: status.Width = 436: status.Height = 88
    status.WordWrap = True: status.Caption = u("Đang dịch theo cài đặt mặc định...")
    Set actionButton = Me.Controls.Add("Forms.CommandButton.1", "btnAction", True)
    actionButton.Left = 332: actionButton.Top = 114: actionButton.Width = 120: actionButton.Height = 28
    actionButton.Caption = u("Dừng")
    running = True
End Sub

Public Sub SetProgress(ByVal message As String)
    status.Caption = message
    Me.Repaint
End Sub

Public Sub BeginRun()
    running = True
    actionButton.Caption = u("Dừng")
    actionButton.Cancel = False
    SetProgress u("Đang dịch theo cài đặt mặc định...")
End Sub

Public Sub Finish(ByVal result As String)
    running = False
    actionButton.Caption = u("Đóng")
    actionButton.Cancel = True
    If Left$(result, 3) = "OK:" Then
        SetProgress u("Đã dịch ") & Mid$(result, 4) & u(" ô vào cột bên phải. Giữ nguyên nguồn. Có thể bấm Hoàn tác trong tab Tiện ích.")
    Else
        SetProgress result
    End If
End Sub

Private Sub actionButton_Click()
    If running Then
        CancelDirectTranslation
        SetProgress u("Đang dừng; chờ yêu cầu mạng hiện tại kết thúc. Chưa ghi kết quả.")
    Else
        Unload Me
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If running Then
        Cancel = True
        actionButton_Click
    End If
End Sub
