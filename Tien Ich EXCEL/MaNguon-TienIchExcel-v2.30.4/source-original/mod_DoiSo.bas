Attribute VB_Name = "mod_DoiSo"
Option Explicit

' Ham giai ma Unicode mo rong cho phep doc moi ky tu dac biet
Private Function UDoc(ByVal str As String) As String
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
    UDoc = res
End Function

' --- DOC SO TIENG VIET ---
Private Function Read3DigitsVN(ByVal num As Long, ByVal isFirstGroup As Boolean) As String
    Dim chuSo As Variant
    chuSo = Array(UDoc("kh\u00f4ng"), UDoc("m\u1ed9t"), "hai", "ba", UDoc("b\u1ed1n"), _
                  UDoc("n\u0103m"), UDoc("s\u00e1u"), UDoc("b\u1ea3y"), UDoc("t\u00e1m"), UDoc("ch\u00edn"))
    
    Dim c As Long, d As Long, u As Long
    c = num \ 100
    d = (num Mod 100) \ 10
    u = num Mod 10
    
    Dim res As String
    res = ""
    
    If c > 0 Or Not isFirstGroup Then
        res = res & chuSo(c) & UDoc(" tr\u0103m ")
    End If
    
    If d > 1 Then
        res = res & chuSo(d) & UDoc(" m\u01b0\u01a1i ")
        If u = 1 Then
            res = res & UDoc("m\u1ed1t ")
        ElseIf u = 5 Then
            res = res & UDoc("l\u0103m ")
        ElseIf u > 0 Then
            res = res & chuSo(u) & " "
        End If
    ElseIf d = 1 Then
        res = res & UDoc("m\u01b0\u1eddi ")
        If u = 5 Then
            res = res & UDoc("l\u0103m ")
        ElseIf u > 0 Then
            res = res & chuSo(u) & " "
        End If
    ElseIf d = 0 And u > 0 Then
        If c > 0 Or Not isFirstGroup Then res = res & UDoc("l\u1ebb ")
        res = res & chuSo(u) & " "
    End If
    
    Read3DigitsVN = res
End Function

' Cong thuc Excel: =DocSo(A1, [DonVi])
Public Function DocSo(ByVal targetNum As Variant, Optional ByVal donVi As String = " \u0111\u1ed3ng.") As String
    Dim num As Double, isNegative As Boolean
    If IsEmpty(targetNum) Or Not IsNumeric(targetNum) Then DocSo = "": Exit Function
    
    num = CDbl(targetNum)
    If num = 0 Then
        DocSo = UDoc("Kh\u00f4ng" & donVi)
        Exit Function
    End If
    
    If num < 0 Then
        isNegative = True
        num = Abs(num)
    End If
    
    Dim groups(1 To 5) As Long, i As Long
    Dim tempNum As Double
    tempNum = Fix(num)
    
    For i = 1 To 5
        groups(i) = tempNum - (Fix(tempNum / 1000) * 1000)
        tempNum = Fix(tempNum / 1000)
        If tempNum = 0 Then Exit For
    Next i
    
    Dim groupUnits As Variant
    groupUnits = Array("", UDoc("ngh\u00ecn"), UDoc("tri\u1ec7u"), UDoc("t\u1ef7"), UDoc("ngh\u00ecn t\u1ef7"))
    
    Dim strRes As String, isFirst As Boolean
    strRes = ""
    isFirst = True
    
    For i = 5 To 1 Step -1
        If groups(i) > 0 Then
            strRes = strRes & Read3DigitsVN(groups(i), isFirst) & groupUnits(i - 1) & " "
            isFirst = False
        End If
    Next i
    
    strRes = Application.Trim(strRes)
    If isNegative Then strRes = UDoc("\u00c2m ") & strRes
    
    ' Viet hoa chu cai dau
    strRes = UCase(Left(strRes, 1)) & Mid(strRes, 2) & UDoc(donVi)
    DocSo = strRes
End Function

' --- DOC SO TIENG ANH ---
Private Function Read3DigitsEN(ByVal num As Long) As String
    Dim ones As Variant, tens As Variant
    ones = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", _
                 "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    
    Dim c As Long, d As Long, u As Long, res As String
    c = num \ 100
    d = num Mod 100
    res = ""
    
    If c > 0 Then
        res = ones(c) & " Hundred "
    End If
    
    If d >= 20 Then
        res = res & tens(d \ 10) & " "
        If (d Mod 10) > 0 Then res = res & ones(d Mod 10) & " "
    ElseIf d > 0 Then
        res = res & ones(d) & " "
    End If
    
    Read3DigitsEN = Trim(res)
End Function

' Cong thuc Excel: =DocSoEN(A1, [Unit])
Public Function DocSoEN(ByVal targetNum As Variant, Optional ByVal unitName As String = " Dollars.") As String
    Dim num As Double, isNegative As Boolean
    If IsEmpty(targetNum) Or Not IsNumeric(targetNum) Then DocSoEN = "": Exit Function
    
    num = CDbl(targetNum)
    If num = 0 Then
        DocSoEN = "Zero" & unitName
        Exit Function
    End If
    
    If num < 0 Then
        isNegative = True
        num = Abs(num)
    End If
    
    Dim groups(1 To 5) As Long, i As Long
    Dim tempNum As Double
    tempNum = Fix(num)
    
    For i = 1 To 5
        groups(i) = tempNum - (Fix(tempNum / 1000) * 1000)
        tempNum = Fix(tempNum / 1000)
        If tempNum = 0 Then Exit For
    Next i
    
    Dim groupUnitsEN As Variant
    groupUnitsEN = Array("", "Thousand", "Million", "Billion", "Trillion")
    
    Dim strRes As String, grpText As String
    strRes = ""
    
    For i = 5 To 1 Step -1
        If groups(i) > 0 Then
            grpText = Read3DigitsEN(groups(i))
            If grpText <> "" Then
                strRes = strRes & grpText & " " & groupUnitsEN(i - 1) & " "
            End If
        End If
    Next i
    
    strRes = Application.Trim(strRes)
    If isNegative Then strRes = "Negative " & strRes
    DocSoEN = strRes & unitName
End Function

' --- CALLBACKS RIBBON ---
Public Sub Ribbon_NumToWordsVN(control As IRibbonControl)
    Dim cell As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    For Each cell In Selection
        If IsNumeric(cell.Value) And cell.Value <> "" Then
            cell.Offset(0, 1).Value = DocSo(cell.Value)
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Ribbon_NumToWordsEN(control As IRibbonControl)
    Dim cell As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    For Each cell In Selection
        If IsNumeric(cell.Value) And cell.Value <> "" Then
            cell.Offset(0, 1).Value = DocSoEN(cell.Value)
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

