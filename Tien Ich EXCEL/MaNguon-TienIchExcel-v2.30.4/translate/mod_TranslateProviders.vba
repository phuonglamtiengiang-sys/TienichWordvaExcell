Option Explicit
Private waits As Object
Private waitReasons As Object
Private lastProvider As String

Public Function TranslationProviderMode() As String
    Dim mode As String
    mode = GetSetting("TienIchExcel", "Translate", "ProviderMode", "auto")
    Select Case mode
        Case "auto", "mymemory", "googleweb": TranslationProviderMode = mode
        Case Else: TranslationProviderMode = "mymemory"
    End Select
End Function

Public Sub ValidateProviderSettings(ByVal mode As String)
    If mode <> "auto" And mode <> "mymemory" And mode <> "googleweb" Then Err.Raise 5, , u("Hãy chọn Tự động, MyMemory hoặc Google web.")
End Sub

Public Sub SaveProviderSettings(ByVal mode As String)
    If TranslationIsBusy() Then Err.Raise 5, , u("Đang dịch; hãy chờ hoặc bấm Dừng trước khi lưu.")
    ValidateProviderSettings mode
    SaveSetting "TienIchExcel", "Translate", "ProviderMode", mode
    Set waits = Nothing: Set waitReasons = Nothing
End Sub

Public Function ProviderConfigurationKey() As String
    ProviderConfigurationKey = TranslationProviderMode()
End Function

Public Function TranslationLastProvider() As String
    TranslationLastProvider = lastProvider
End Function

Private Function ProviderJSONString(ByVal json As String, ByRef position As Long) As String
    Dim ch As String, code As Long
    position = position + 1
    Do While position <= Len(json)
        ch = Mid$(json, position, 1): position = position + 1
        If ch = """" Then Exit Function
        If ch = "\" Then
            ch = Mid$(json, position, 1): position = position + 1
            Select Case ch
                Case "u"
                    code = CLng("&H" & Mid$(json, position, 4))
                    If code > 32767 Then code = code - 65536
                    ch = ChrW(code): position = position + 4
                Case "n": ch = vbLf
                Case "r": ch = vbCr
                Case "t": ch = vbTab
                Case "b": ch = Chr$(8)
                Case "f": ch = Chr$(12)
                Case """", "\", "/"
                Case Else: Err.Raise 5
            End Select
        End If
        ProviderJSONString = ProviderJSONString & ch
    Loop
    Err.Raise 5, , "Unterminated JSON string"
End Function

Public Function GoogleWebTranslation(ByVal json As String) As String
    Dim i As Long, depth As Long, first As Boolean, ch As String, token As String
    json = Trim$(json)
    If Left$(json, 1) <> "[" Then Err.Raise 5
    i = 1
    Do While i <= Len(json)
        ch = Mid$(json, i, 1)
        Select Case ch
            Case "["
                depth = depth + 1
                If depth = 3 Then first = True
            Case "]"
                depth = depth - 1
                If depth = 1 Then Exit Function
            Case """"
                token = ProviderJSONString(json, i)
                If depth = 3 And first Then GoogleWebTranslation = GoogleWebTranslation & token: first = False
                i = i - 1
            Case "n", "0" To "9"
                If depth = 3 And first Then first = False
        End Select
        i = i + 1
    Loop
    Err.Raise 5, , "Incomplete translation response"
End Function

Public Function ProviderResponse(ByVal provider As String, ByVal status As Long, ByVal json As String, ByVal sourceText As String, ByRef translated As String, ByRef reason As String, ByRef quota As Boolean) As Boolean
    On Error GoTo Failed
    translated = "": reason = "": quota = (status = 429)
    If provider = "mymemory" Then
        If ExtractJSONValue(json, "quotaFinished") = "true" Or ExtractJSONValue(json, "responseStatus") = "429" Then quota = True
    End If
    If quota Then reason = u("Hết hạn mức hoặc bị giới hạn lượt gọi."): Exit Function
    If status <> 200 Then reason = "HTTP " & status: Exit Function
    If provider = "googleweb" Then
        translated = GoogleWebTranslation(json)
        If Len(translated) = 0 Then reason = u("Google web không trả bản dịch."): Exit Function
        ProviderResponse = True
        Exit Function
    End If
    If provider = "mymemory" Then
        ProviderResponse = ParseTranslation(json, translated, reason, sourceText)
        Exit Function
    End If
    reason = u("Dịch vụ chưa được hỗ trợ.")
    Exit Function
Failed:
    reason = u("Không đọc được phản hồi của máy chủ.")
End Function

Private Function RequestProvider(ByVal provider As String, ByVal text As String, ByVal source As String, ByVal target As String) As Variant
    Dim http As Object, url As String, body As String, translated As String, reason As String, quota As Boolean
    Dim key As String, status As Long, json As String, expires As Date
    On Error GoTo Failed
    key = provider
    If provider = "mymemory" Then key = key & "|" & TranslationEmail()
    If waits Is Nothing Then Set waits = CreateObject("Scripting.Dictionary"): Set waitReasons = CreateObject("Scripting.Dictionary")
    If waits.Exists(key) Then
        If Now < waits(key) Then
            RequestProvider = Array(False, waitReasons(key)(0), waits(key), waitReasons(key)(1)): Exit Function
        End If
    End If
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 3000, 3000, 5000, 8000
    Select Case provider
        Case "mymemory"
            url = TranslationRequestURL(text, source, target)
            http.Open "GET", url, False
            http.Send
        Case "googleweb"
            url = "https://translate.googleapis.com/translate_a/single?client=gtx&dt=t&sl=" & EncodeURL(source) & "&tl=" & EncodeURL(target) & "&q=" & EncodeURL(text)
            http.Open "GET", url, False
            http.Send
        Case Else
            Err.Raise 5, , u("Dịch vụ không hợp lệ.")
    End Select
    status = http.Status: json = DecodeUTF8(http.ResponseBody)
    If ProviderResponse(provider, status, json, text, translated, reason, quota) Then
        lastProvider = provider
        RequestProvider = Array(True, translated, Now + 365, False)
        Exit Function
    End If
    GoTo RememberFailure
Failed:
    reason = u("Không kết nối được máy chủ dịch.")
RememberFailure:
    expires = DateAdd("s", 60, Now)
    If quota Then expires = DateAdd("n", 15, Now)
    waits(key) = expires
    waitReasons(key) = Array(provider & ": " & Left$(reason, 250), quota)
    RequestProvider = Array(False, provider & ": " & Left$(reason, 250), expires, quota)
End Function

Public Function ProviderTranslate(ByVal text As String, ByVal source As String, ByVal target As String) As Variant
    Dim mode As String, entry As Variant, errors As String, provider As Variant
    mode = TranslationProviderMode()
    If mode <> "auto" Then
        ProviderTranslate = RequestProvider(mode, text, source, target)
        Exit Function
    End If
    entry = RequestProvider("mymemory", text, source, target)
    If entry(0) Or Not entry(3) Then ProviderTranslate = entry: Exit Function
    errors = entry(1)
    entry = RequestProvider("googleweb", text, source, target)
    If entry(0) Then ProviderTranslate = entry: Exit Function
    ProviderTranslate = Array(False, errors & vbCrLf & entry(1) & vbCrLf & u("Chưa dịch được. Thử lại sau hoặc đổi cách dịch trong Cài đặt dịch."), entry(2), True)
End Function
