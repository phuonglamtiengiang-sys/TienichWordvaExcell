Option Explicit
Public Function QRImageExtension(ByVal bytes As Variant, ByVal kind As String) As String
    Dim first As Long, doc As Object, node As Object, attr As Object, allowed As String, value As String
    first = LBound(bytes)
    If UBound(bytes) - first < 23 Or UBound(bytes) - first > 10000000 Then Err.Raise 5, , "Du lieu anh khong hop le."
    If bytes(first) = 137 And bytes(first + 1) = 80 And bytes(first + 2) = 78 And bytes(first + 3) = 71 Then
        QRImageExtension = ".png": Exit Function
    End If
    If kind <> "vietqr" Then Err.Raise 5, , "Dich vu khong tra ve anh PNG hop le."
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False: doc.resolveExternals = False: doc.validateOnParse = False
    doc.setProperty "ProhibitDTD", True
    If Not doc.LoadXML(DecodeUTF8(bytes)) Then Err.Raise 5, , "Dich vu khong tra ve anh VietQR hop le."
    If doc.DocumentElement.baseName <> "svg" Or doc.DocumentElement.namespaceURI <> "http://www.w3.org/2000/svg" Then Err.Raise 5, , "Phan hoi khong phai anh SVG."
    allowed = "|svg|title|desc|rect|g|path|image|text|tspan|defs|linearGradient|stop|clipPath|circle|ellipse|line|polyline|polygon|"
    For Each node In doc.SelectNodes("//*")
        If node.namespaceURI <> "http://www.w3.org/2000/svg" Or InStr(1, allowed, "|" & node.baseName & "|", vbBinaryCompare) = 0 Then Err.Raise 5, , "SVG co thanh phan khong ho tro."
        For Each attr In node.Attributes
            value = LCase$(CStr(attr.Text))
            If LCase$(Left$(attr.baseName, 2)) = "on" Or InStr(value, "url(") > 0 Then Err.Raise 5, , "SVG co noi dung khong ho tro."
            If attr.baseName = "href" Then
                If Left$(value, 22) <> "data:image/png;base64," And Left$(value, 23) <> "data:image/jpeg;base64," Then Err.Raise 5, , "Khong tai tai nguyen ngoai tu SVG."
            End If
        Next attr
    Next node
    QRImageExtension = ".svg"
End Function
