Attribute VB_Name = "mod_MaVachQR"
Option Explicit

' Ham chen anh ma vach / QR vao o
Private Sub ChenAnhVaoCell(ByVal targetCell As Range, ByVal imgURL As String, ByVal imgHeight As Double)
    On Error Resume Next
    Dim pic As Picture
    Set pic = targetCell.Parent.Pictures.Insert(imgURL)
    If Not pic Is Nothing Then
        pic.Top = targetCell.Top + 2
        pic.Left = targetCell.Left + 2
        pic.Height = imgHeight
    End If
    On Error GoTo 0
End Sub

' 1. Tao Barcode Code 128
Public Sub Ribbon_Code128(control As IRibbonControl)
    Dim cell As Range, val As String, url As String
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    For Each cell In Selection
        val = Trim(CStr(cell.Value))
        If val <> "" Then
            url = "https://bwipjs-api.metafloor.com/?bcid=code128&text=" & val & "&scale=2&includetext"
            Call ChenAnhVaoCell(cell.Offset(0, 1), url, 40)
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

' 2. Tao Barcode Code 39
Public Sub Ribbon_Code39(control As IRibbonControl)
    Dim cell As Range, val As String, url As String
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    For Each cell In Selection
        val = Trim(CStr(cell.Value))
        If val <> "" Then
            url = "https://bwipjs-api.metafloor.com/?bcid=code39&text=" & UCase(val) & "&scale=2&includetext"
            Call ChenAnhVaoCell(cell.Offset(0, 1), url, 40)
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

' 3. Tao Barcode EAN-13
Public Sub Ribbon_EAN13(control As IRibbonControl)
    Dim cell As Range, val As String, url As String
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    For Each cell In Selection
        val = Trim(CStr(cell.Value))
        If Len(val) = 12 Or Len(val) = 13 Then
            url = "https://bwipjs-api.metafloor.com/?bcid=ean13&text=" & val & "&scale=2&includetext"
            Call ChenAnhVaoCell(cell.Offset(0, 1), url, 40)
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

' 4. Tao QR Code
Public Sub Ribbon_QRCode(control As IRibbonControl)
    Dim cell As Range, val As String, url As String
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    For Each cell In Selection
        val = Trim(CStr(cell.Value))
        If val <> "" Then
            url = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" & val
            Call ChenAnhVaoCell(cell.Offset(0, 1), url, 80)
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

' 5. XOA TOAN BO ANH MA (BARCODE / QR CODE) TRONG VUNG CHON
Public Sub Ribbon_Clear_Barcodes(control As IRibbonControl)
    Dim selRange As Range, shp As Shape, count As Long
    Dim cellLeft As Double, cellTop As Double, cellRight As Double, cellBottom As Double
    Dim shpLeft As Double, shpTop As Double
    
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set selRange = Selection
    count = 0
    
    cellLeft = selRange.Left
    cellTop = selRange.Top
    cellRight = selRange.Left + selRange.Width
    cellBottom = selRange.Top + selRange.Height
    
    For Each shp In selRange.Parent.Shapes
        shpLeft = shp.Left
        shpTop = shp.Top
        ' Kiem tra vi tri anh nam ben trong khoi duoc chon
        If shpLeft >= cellLeft And shpLeft <= cellRight And shpTop >= cellTop And shpTop <= cellBottom Then
            shp.Delete
            count = count + 1
        End If
    Next shp
    
    ShowMsg "\u0110\u00e3 x\u00f3a " & count & " \u1ea3nh m\u00e3 trong v\u00f9ng ch\u1ecdn!", "X\u00f3a m\u00e3"
End Sub
