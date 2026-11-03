Attribute VB_Name = "mod_ThongTin"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" ( _
        ByVal hwnd As Long, ByVal lpText As Long, _
        ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

' Ham giai ma Unicode dang \uXXXX dung chung cho toan bo Add-in
Public Function u(ByVal str As String) As String
    Dim i As Long, res As String, hexCode As String
    res = ""
    i = 1
    Do While i <= Len(str)
        If Mid(str, i, 2) = "\u" And i + 5 <= Len(str) Then
            hexCode = Mid(str, i + 2, 4)
            res = res & ChrW(val("&H" & hexCode))
            i = i + 6
        Else
            res = res & Mid(str, i, 1)
            i = i + 1
        End If
    Loop
    u = res
End Function

Public Sub ShowMsg(ByVal prompt As String, Optional ByVal title As String = "Th\u00f4ng b\u00e1o", Optional ByVal buttons As Long = 64)
    MessageBoxW Application.hwnd, StrPtr(u(prompt)), StrPtr(u(title)), buttons
End Sub

' Callback nut Gioi thieu
Public Sub Ribbon_About(control As IRibbonControl)
    ShowMsg "Add-in Ti\u1ec7n \u00edch Excel Vi\u1ec7t Nam" & vbCrLf & _
            "T\u00e1c gi\u1ea3: H\u1ed3 Ph\u01b0\u01a1ng L\u00e2m" & vbCrLf & _
            "Email: hophuonglam@outlook.com", "Gi\u1edbi thi\u1ec7u", 64
End Sub

