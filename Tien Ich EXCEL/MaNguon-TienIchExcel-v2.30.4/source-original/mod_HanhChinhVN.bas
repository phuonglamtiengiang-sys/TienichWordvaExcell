Attribute VB_Name = "mod_HanhChinhVN"
Option Explicit

Private Function ChuanHoaCCCD(ByVal cccd As Variant) As String
    Dim s As String
    s = Trim(CStr(cccd))
    s = Replace(s, " ", "")
    s = Replace(s, ChrW(160), "")
    If IsNumeric(s) And InStr(s, ".") = 0 And InStr(s, "-") = 0 Then
        If Len(s) >= 9 And Len(s) < 12 Then s = Right("000000000000" & s, 12)
    End If
    ChuanHoaCCCD = s
End Function

Private Function GetTenTinhCCCD(ByVal maTinh As String, Optional ByVal LayTenCu As Boolean = False) As String
    Select Case maTinh
        Case "001": GetTenTinhCCCD = u("H\u00e0 N\u1ed9i")
        Case "002": GetTenTinhCCCD = u("H\u00e0 Giang")
        Case "004": GetTenTinhCCCD = u("Cao B\u1eb1ng")
        Case "006": GetTenTinhCCCD = u("B\u1eafc K\u1ea1n")
        Case "008": GetTenTinhCCCD = u("Tuy\u00ean Quang")
        Case "010": GetTenTinhCCCD = u("L\u00e0o Cai")
        Case "011": GetTenTinhCCCD = u("\u0110i\u1ec7n Bi\u00ean")
        Case "012": GetTenTinhCCCD = u("Lai Ch\u00e2u")
        Case "014": GetTenTinhCCCD = u("S\u01a1n La")
        Case "015": GetTenTinhCCCD = u("Y\u00ean B\u00e1i")
        Case "017": GetTenTinhCCCD = u("H\u00f2a B\u00ecnh")
        Case "019": GetTenTinhCCCD = u("Th\u00e1i Nguy\u00ean")
        Case "020": GetTenTinhCCCD = u("L\u1ea1ng S\u01a1n")
        Case "022": GetTenTinhCCCD = u("Qu\u1ea3ng Ninh")
        Case "024": GetTenTinhCCCD = u("B\u1eafc Giang")
        Case "025": GetTenTinhCCCD = u("Ph\u00fa Th\u1ecd")
        Case "026": GetTenTinhCCCD = u("V\u0129nh Ph\u00fac")
        Case "027": GetTenTinhCCCD = u("B\u1eafc Ninh")
        Case "030": GetTenTinhCCCD = u("H\u1ea3i D\u01b0\u01a1ng")
        Case "031": GetTenTinhCCCD = u("H\u1ea3i Ph\u00f2ng")
        Case "033": If LayTenCu Then GetTenTinhCCCD = u("H\u00e0 T\u00e2y (c\u0169)") Else GetTenTinhCCCD = u("H\u00e0 N\u1ed9i")
        Case "034": GetTenTinhCCCD = u("Th\u00e1i B\u00ecnh")
        Case "035": GetTenTinhCCCD = u("H\u00e0 Nam")
        Case "036": GetTenTinhCCCD = u("Nam \u0110\u1ecbnh")
        Case "037": GetTenTinhCCCD = u("Ninh B\u00ecnh")
        Case "038": GetTenTinhCCCD = u("Thanh H\u00f3a")
        Case "040": GetTenTinhCCCD = u("Ngh\u1ec7 An")
        Case "042": GetTenTinhCCCD = u("H\u00e0 T\u0129nh")
        Case "044": GetTenTinhCCCD = u("Qu\u1ea3ng B\u00ecnh")
        Case "045": GetTenTinhCCCD = u("Qu\u1ea3ng Tr\u1ecb")
        Case "046": If LayTenCu Then GetTenTinhCCCD = u("Th\u1eeba Thi\u00ean Hu\u1ebf") Else GetTenTinhCCCD = u("Hu\u1ebf")
        Case "048": GetTenTinhCCCD = u("\u0110\u00e0 N\u1eb5ng")
        Case "049": GetTenTinhCCCD = u("Qu\u1ea3ng Nam")
        Case "051": GetTenTinhCCCD = u("Qu\u1ea3ng Ng\u00e3i")
        Case "052": GetTenTinhCCCD = u("B\u00ecnh \u0110\u1ecbnh")
        Case "054": GetTenTinhCCCD = u("Ph\u00fa Y\u00ean")
        Case "056": GetTenTinhCCCD = u("Kh\u00e1nh H\u00f2a")
        Case "058": GetTenTinhCCCD = u("Ninh Thu\u1eadn")
        Case "060": GetTenTinhCCCD = u("B\u00ecnh Thu\u1eadn")
        Case "062": GetTenTinhCCCD = u("Kon Tum")
        Case "064": GetTenTinhCCCD = u("Gia Lai")
        Case "066": GetTenTinhCCCD = u("\u0110\u1eafk L\u1eafk")
        Case "067": GetTenTinhCCCD = u("\u0110\u1eafk N\u00f4ng")
        Case "068": GetTenTinhCCCD = u("L\u00e2m \u0110\u1ed3ng")
        Case "070": GetTenTinhCCCD = u("B\u00ecnh Ph\u01b0\u1edbc")
        Case "072": GetTenTinhCCCD = u("T\u00e2y Ninh")
        Case "074": GetTenTinhCCCD = u("B\u00ecnh D\u01b0\u01a1ng")
        Case "075": GetTenTinhCCCD = u("\u0110\u1ed3ng Nai")
        Case "077": If LayTenCu Then GetTenTinhCCCD = u("B\u00e0 R\u1ecba - V\u0169ng T\u00e0u") Else GetTenTinhCCCD = u("B\u00e0 R\u1ecba \u2013 V\u0169ng T\u00e0u")
        Case "079": If LayTenCu Then GetTenTinhCCCD = u("Gia \u0110\u1ecbnh / S\u00e0i G\u00f2n") Else GetTenTinhCCCD = u("TP. H\u1ed3 Ch\u00ed Minh")
        Case "080": GetTenTinhCCCD = u("Long An")
        Case "082": GetTenTinhCCCD = u("Ti\u1ec1n Giang")
        Case "083": GetTenTinhCCCD = u("B\u1ebfn Tre")
        Case "084": GetTenTinhCCCD = u("Tr\u00e0 Vinh")
        Case "086": GetTenTinhCCCD = u("V\u0129nh Long")
        Case "087": GetTenTinhCCCD = u("\u0110\u1ed3ng Th\u00e1p")
        Case "089": GetTenTinhCCCD = u("An Giang")
        Case "091": GetTenTinhCCCD = u("Ki\u00ean Giang")
        Case "092": If LayTenCu Then GetTenTinhCCCD = u("C\u1ea7n Th\u01a1 (t\u1ec9nh c\u0169)") Else GetTenTinhCCCD = u("C\u1ea7n Th\u01a1")
        Case "093": GetTenTinhCCCD = u("H\u1eadu Giang")
        Case "094": GetTenTinhCCCD = u("S\u00f3c Tr\u0103ng")
        Case "095": GetTenTinhCCCD = u("B\u1ea1c Li\u00eau")
        Case "096": GetTenTinhCCCD = u("C\u00e0 Mau")
        Case Else: GetTenTinhCCCD = ""
    End Select
End Function

Public Function KiemTraCCCD(ByVal cccd As Variant) As Boolean
    Dim s As String, maTinh As String, maGT As Integer
    s = ChuanHoaCCCD(cccd)
    If Len(s) <> 12 Or Not IsNumeric(s) Then KiemTraCCCD = False: Exit Function
    maTinh = Left(s, 3)
    If GetTenTinhCCCD(maTinh) = "" Then KiemTraCCCD = False: Exit Function
    maGT = CInt(Mid(s, 4, 1))
    If maGT < 0 Or maGT > 9 Then KiemTraCCCD = False: Exit Function
    KiemTraCCCD = True
End Function

Public Function CCCD_GioiTinh(ByVal cccd As Variant) As String
    Dim s As String, maGT As Integer
    s = ChuanHoaCCCD(cccd)
    If Not KiemTraCCCD(s) Then CCCD_GioiTinh = "": Exit Function
    maGT = CInt(Mid(s, 4, 1))
    Select Case maGT
        Case 0, 2, 4, 6, 8: CCCD_GioiTinh = "Nam"
        Case 1, 3, 5, 7, 9: CCCD_GioiTinh = u("N\u1eef")
    End Select
End Function

Public Function CCCD_NamSinh(ByVal cccd As Variant) As Variant
    Dim s As String, maGT As Integer, theKy As Integer
    s = ChuanHoaCCCD(cccd)
    If Not KiemTraCCCD(s) Then CCCD_NamSinh = "": Exit Function
    maGT = CInt(Mid(s, 4, 1))
    Select Case maGT
        Case 0, 1: theKy = 1900
        Case 2, 3: theKy = 2000
        Case 4, 5: theKy = 2100
        Case 6, 7: theKy = 2200
        Case 8, 9: theKy = 2300
    End Select
    CCCD_NamSinh = theKy + CInt(Mid(s, 5, 2))
End Function

Public Function CCCD_TinhThanh(ByVal cccd As Variant, Optional ByVal LayTenCu As Boolean = False) As String
    Dim s As String, maTinh As String
    s = ChuanHoaCCCD(cccd)
    If Not KiemTraCCCD(s) Then CCCD_TinhThanh = "": Exit Function
    maTinh = Left(s, 3)
    CCCD_TinhThanh = GetTenTinhCCCD(maTinh, LayTenCu)
End Function

Private Function ExtractJSONValue(ByVal json As String, ByVal key As String) As String
    Dim posKey As Long, posStart As Long, posEnd As Long, searchKey As String, rawVal As String
    searchKey = """" & key & """"
    posKey = InStr(1, json, searchKey, vbTextCompare)
    If posKey = 0 Then Exit Function
    posStart = InStr(posKey + Len(searchKey), json, ":")
    If posStart = 0 Then Exit Function
    posStart = posStart + 1
    While Mid(json, posStart, 1) = " " Or Mid(json, posStart, 1) = vbTab
        posStart = posStart + 1
    Wend
    If Mid(json, posStart, 1) = """" Then
        posStart = posStart + 1
        posEnd = InStr(posStart, json, """")
        If posEnd > 0 Then rawVal = Mid(json, posStart, posEnd - posStart)
    ElseIf Mid(json, posStart, 4) = "null" Then
        rawVal = ""
    Else
        posEnd = InStr(posStart, json, ",")
        If posEnd = 0 Then posEnd = InStr(posStart, json, "}")
        If posEnd > 0 Then rawVal = Trim(Mid(json, posStart, posEnd - posStart))
    End If
    rawVal = Replace(rawVal, "\""", """")
    rawVal = Replace(rawVal, "\/", "/")
    rawVal = Replace(rawVal, ChrW(160), " ")
    ExtractJSONValue = Application.Trim(u(rawVal))
End Function

Public Sub Ribbon_CheckCCCD(control As IRibbonControl)
    Dim cell As Range, cccd As String, count As Long
    If TypeName(Selection) <> "Range" Then Exit Sub
    count = 0: Application.ScreenUpdating = False
    For Each cell In Selection
        cccd = ChuanHoaCCCD(cell.Value)
        If cccd <> "" Then
            If KiemTraCCCD(cccd) Then
                cell.NumberFormat = "@"
                cell.Value = cccd
                cell.Offset(0, 1).Value = CCCD_GioiTinh(cccd)
                cell.Offset(0, 2).Value = CCCD_NamSinh(cccd)
                cell.Offset(0, 3).Value = CCCD_TinhThanh(cccd)
                count = count + 1
            Else
                cell.Offset(0, 1).Value = u("CCCD kh\u00f4ng h\u1ee3p l\u1ec7")
                cell.Offset(0, 2).ClearContents
                cell.Offset(0, 3).ClearContents
            End If
        End If
    Next cell
    Application.ScreenUpdating = True
    ShowMsg "\u0110\u00e3 x\u1eed l\u00fd v\u00e0 ki\u1ec3m tra xong " & count & " m\u00e3 CCCD!"
End Sub

Public Sub Ribbon_LookupMST(control As IRibbonControl)
    Dim cell As Range, mst As String, cleanMST As String, parentMST As String
    Dim compName As String, compAddr As String, count As Long
    Dim http As Object, url As String, resJSON As String
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    count = 0
    
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    http.Option(4) = 13056: http.Option(9) = 2048
    
    Application.ScreenUpdating = False
    For Each cell In Selection
        mst = Trim(CStr(cell.Value))
        cleanMST = Replace(Replace(mst, " ", ""), ChrW(160), "")
        If cleanMST <> "" Then
            parentMST = cleanMST
            If InStr(cleanMST, "-") > 0 Then parentMST = Split(cleanMST, "-")(0)
            url = "https://api.vietqr.io/v2/business/" & parentMST
            http.Open "GET", url, False
            http.setRequestHeader "Content-Type", "application/json"
            http.send
            
            If Err.Number = 0 And http.Status = 200 Then
                resJSON = http.responseText
                If InStr(resJSON, """code"":""00""") > 0 Then
                    compName = ExtractJSONValue(resJSON, "name")
                    compAddr = ExtractJSONValue(resJSON, "address")
                    If InStr(cleanMST, "-") > 0 Then
                        compName = compName & " - Chi nh" & ChrW(225) & "nh " & Split(cleanMST, "-")(1)
                    End If
                Else
                    compName = u("Kh\u00f4ng t\u00ecm th\u1ea5y MST"): compAddr = ""
                End If
            Else
                compName = u("L\u1ed7i k\u1ebft n\u1ed1i"): compAddr = ""
            End If
            cell.Offset(0, 1).Value = compName
            cell.Offset(0, 2).Value = compAddr
            count = count + 1
        End If
    Next cell
    Application.ScreenUpdating = True
    On Error GoTo 0
    ShowMsg "\u0110\u00e3 tra c\u1ee9u xong cho " & count & " M\u00e3 s\u1ed1 thu\u1ebf!"
End Sub
