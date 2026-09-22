Attribute VB_Name = "AppTestRouter"
Option Explicit

'Private Function generateResponseExtraHeaders() As CVtHttpaHeaders
'    Dim extraHeaders As New CVtHttpaHeaders
'    ' --- [Policy 1: 디버깅 및 커스텀 정보] ---
'    ' 2. MULTI-KEEP: 여러 개의 쿠키 설정 (브라우저는 이를 각각의 쿠키로 인식함)
'    ' 엔진 내부적으로는 'Set-Cookie'라는 키에 값의 배열(Array)이 생성되어야 함
'    Call extraHeaders.AddHeader("Set-Cookie", "UserRole=Admin; Path=/; HttpOnly; SameSite=Strict")
'    Call extraHeaders.AddHeader("Set-Cookie", "NodeID=Edge-01; Max-Age=86400")
'    Call extraHeaders.AddHeader("Set-Cookie", "Theme=Dark; Path=/")
'
'    ' 3. 캐시 제어 및 프록시 지시어
'    ' 쉼표(,)로 결합하지 않고 개별 헤더 라인으로 송출하는 것이 안전할 때가 있음
'    Call extraHeaders.AddHeader("Cache-Control", "no-store")
'    Call extraHeaders.AddHeader("Cache-Control", "max-age=0")
'
'    ' 4. 보안 정책 (CORS 가변 헤더)
'    Call extraHeaders.AddHeader("Access-Control-Expose-Headers", "X-VBA-Token")
'    Call extraHeaders.AddHeader("Access-Control-Expose-Headers", "X-VBA-Version")
'
'
'     Set generateResponseExtraHeaders = extraHeaders
'
'
'End Function

Private Const MODULE_NAME As String = "AppTestRouter"

Public Sub GetStatus( _
    ByVal s As CVtHttpaSession, _
    ByVal CustomData As CWasCustomData _
)

    Const PROC_NAME As String = "GetStatus"
    
    Dim StatusCodeParam As String
    Dim StatusCode As Long

    Dim errorOut As String
    Dim response As CVtHttpaResponse

    Dim rspMsg As String

    ' --------------------------------------------------------
    ' 1. Get statusCode from Query String
    '
    ' Example:
    ' /vba/response/status?statusCode=404
    ' --------------------------------------------------------
    StatusCodeParam = _
        s.request.QueryParams.GetParam("statusCode")

    ' --------------------------------------------------------
    ' 2. Determine Status Code
    ' --------------------------------------------------------
    If Len(StatusCodeParam) = 0 Or _
       Not IsNumeric(StatusCodeParam) Then

        StatusCode = 200

    Else

        StatusCode = CLng(StatusCodeParam)

    End If

    ' --------------------------------------------------------
    ' 3. Send Response
    ' --------------------------------------------------------
    Set response = VtHttpaFactory.VtHttpaCreateResponse()

    response.StatusCode = StatusCode
    Call response.Headers.AddHeader("X-VTNODE", "MVP")
        

    If response.SendStatus( _
            s, _
            StatusCode, _
            errorOut) Then

'        Debug.Print _
'            "Send Status Success: " & StatusCode

    Else

'        Debug.Print _
'            "Send Status Failed: " & errorOut

    End If

    ' --------------------------------------------------------
    ' 4. Log
    ' --------------------------------------------------------
    rspMsg = _
        "Response.SendStatus(s, " & _
        StatusCode & _
        ") -> " & _
        errorOut

   ' Debug.Print rspMsg

    VtLogInfo MODULE_NAME, PROC_NAME, rspMsg
    CustomData.ResponseStatus = rspMsg

    'UISheetLog.WriteResponseToSheet s, rspMsg

    Set response = Nothing

End Sub
 

Public Sub GetText( _
    ByVal s As CVtHttpaSession, _
    ByVal CustomData As CWasCustomData _
)

    Const PROC_NAME As String = "GetText"
    Dim userName As String
    Dim responseMsg As String

    Dim response As CVtHttpaResponse
    Dim errorOut As String
    Dim rspMsg As String

    ' --------------------------------------------------------
    ' 1. Get name parameter
    '
    ' Example:
    ' /vba/text?name=낭만C인
    ' --------------------------------------------------------
    userName = _
        s.request.QueryParams.GetParam("name")

    If Len(userName) = 0 Then
        userName = "Guest"
    End If

    ' --------------------------------------------------------
    ' 2. Build Response Text
    ' --------------------------------------------------------
    responseMsg = _
        "<html><body>" & _
        "<h1>Hello, " & userName & "!</h1>" & _
        "<p>This is a response from " & _
        "<b>VTNode.io</b> engine running on Excel VBA.</p>" & _
        "</body></html>"

    ' --------------------------------------------------------
    ' 3. Create VBA Response
    ' --------------------------------------------------------
    Set response = VtHttpaFactory.VtHttpaCreateResponse()
    

    response.StatusCode = 200

    Call response.Headers.SetHeader("X-VTNODE", "MVP")
    Call response.Headers.SetHeader("Content-Type", "text/json; charset=utf-8")

    ' --------------------------------------------------------
    ' 4. Send Response
    ' --------------------------------------------------------
    If response.SendText( _
            s, _
            responseMsg, _
            errorOut) Then

        rspMsg = _
            "Response.SendText(name=" & _
            userName & _
            ") -> Success"

'        Debug.Print _
'            "Send Text Success"

    Else

        rspMsg = _
            "Response.SendText(name=" & _
            userName & _
            ") -> Failed: " & _
            errorOut

'        Debug.Print _
'            "Send Text Failed: " & _
'            errorOut

    End If

CleanUp:

    If Len(rspMsg) = 0 Then

        rspMsg = _
            "Response.SendText(name=" & _
            userName & _
            ") -> Failed: " & _
            errorOut

    End If

'    Debug.Print rspMsg

    VtLogInfo MODULE_NAME, PROC_NAME, rspMsg
    
    
    CustomData.ResponseStatus = rspMsg

    Set response = Nothing

    'UISheetLog.WriteResponseToSheet s, responseMsg

End Sub



Public Sub GetFile( _
    ByVal s As CVtHttpaSession, _
    ByVal CustomData As CWasCustomData _
)

    Const PROC_NAME As String = "GetFile"
    Dim fileName As String
    Dim FullPath As String

    Dim response As CVtHttpaResponse
    Dim errorOut As String
    Dim rspMsg As String

    ' --------------------------------------------------------
    ' 1. File Name
    ' --------------------------------------------------------
    fileName = "intro-16k-16bit.wav"

    ' --------------------------------------------------------
    ' 2. Build Full Path
    ' --------------------------------------------------------
    FullPath = _
        VtPathCombine( _
            ThisWorkbook.path, _
            fileName)

    ' --------------------------------------------------------
    ' 3. Check File
    ' --------------------------------------------------------
    If dir$(FullPath) = vbNullString Then

        Debug.Print _
            "SendFile Error: File not found -> " & _
            FullPath

        Set response = VtHttpaFactory.VtHttpaCreateResponse()
        

        response.StatusCode = 404

        Call response.SendStatus( _
            s, _
            404, _
            errorOut)

        Set response = Nothing

        rspMsg = _
            "Response.SendFile(" & _
            FullPath & _
            ") -> File Not Found"

        'Debug.Print rspMsg
        CustomData.ResponseStatus = rspMsg

        Exit Sub

    End If

    ' --------------------------------------------------------
    ' 4. Create VBA Response
    ' --------------------------------------------------------
    Set response = VtHttpaFactory.VtHttpaCreateResponse()
    
    response.StatusCode = 200

    Call response.Headers.SetHeader( _
        "X-VTNODE", _
        "MVP")

    Call response.Headers.SetHeader( _
        "Content-Type", _
        "audio/wav")

    ' --------------------------------------------------------
    ' 5. Send Response
    ' --------------------------------------------------------
    If response.SendFile( _
            s, _
            FullPath, _
            errorOut) Then

        rspMsg = _
            "Response.SendFile(" & _
            FullPath & _
            ") -> Success"

'        Debug.Print _
'            "Send File Success: " & _
'            FullPath

    Else

        rspMsg = _
            "Response.SendFile(" & _
            FullPath & _
            ") -> Failed: " & _
            errorOut

'        Debug.Print _
'            "Send File Failed: " & _
'            errorOut

    End If

    ' --------------------------------------------------------
    ' 6. Log
    ' --------------------------------------------------------
    'Debug.Print rspMsg

    VtLogInfo MODULE_NAME, PROC_NAME, rspMsg
    CustomData.ResponseStatus = rspMsg

    Set response = Nothing

    'UISheetLog.WriteResponseToSheet s, rspMsg

End Sub

Public Sub GetMultipart(ByVal s As CVtHttpaSession, ByVal CustomData As CWasCustomData)
    
    Const PROC_NAME As String = "GetMultipart"
    On Error GoTo ErrorHandler


    Dim errorOut As String
    Dim response As CVtHttpaResponse
    
    ' --------------------------------------------------------
    ' 4. Create VBA Response
    ' --------------------------------------------------------
    Set response = VtHttpaFactory.VtHttpaCreateResponse()


    Call response.Headers.SetHeader("X-VTNODE", "MVP")
    
    
    
    Dim part As CVtHttpaContentPart
    
    Set part = response.AddText("{""result"": ""success""}", "application/json", "form1")
    Call part.Headers.AddHeader("X-HEADER1", "My Header1")
    
    
    Dim fileName As String: fileName = "intro-16k-16bit.wav"
    Dim FullPath As String: FullPath = VtPathCombine(ThisWorkbook.path, fileName)
    Set part = response.AddFile(FullPath, "audio/wav", "form2")
    Call part.Headers.AddHeader("X-HEADER2", "My Header2")
    
    
    Dim Res As Boolean
    Dim rspMsg As String
    
    
    Res = response.SendMultipart(s, "form-data", errorOut)
    
    If Res Then
    
       rspMsg = "SendMultipart OK"
    Else
        rspMsg = "SendMultipart Failed." & errorOut
       
       
       
    End If
    
    
            
CleanUp:
    'Debug.Print rspMsg
    
    VtLogInfo MODULE_NAME, PROC_NAME, rspMsg
    CustomData.ResponseStatus = rspMsg
    
    Set response = Nothing
    
    If Not Res Then
      Call s.request.QueryParams.AddParam("statusCode", "501")
      Call GetStatus(s, CustomData)
    End If
    
    
    Exit Sub

ErrorHandler:
           
    rspMsg = "response_multipart Error: " & Err.description
    
    'Call UITabLog.WriteLog("response_multipart Error: " & Err.Description)
    
    GoTo CleanUp
    
End Sub



Public Sub PostText( _
    ByVal s As CVtHttpaSession, _
    ByVal CustomData As CWasCustomData)
    
    Const PROC_NAME As String = "PostText"
   
    Dim Body As CVtHttpaContentPart
    
    Set Body = s.request.Body
    

    VtLogInfo MODULE_NAME, PROC_NAME, "Body.Length=" & Body.ByteLength
    VtLogInfo MODULE_NAME, PROC_NAME, "Body.Text=" & Body.GetText()
    
        
    Call RespondOK(s, CustomData)
    
    
'    Call s.Request.QueryParams.AddParam("statusCode", "200")
'    Call GetStatus(s, CustomData)
    


End Sub
Public Sub PostFile( _
    ByVal s As CVtHttpaSession, _
    ByVal CustomData As CWasCustomData)
    
     Const PROC_NAME As String = "PostFile"
     Dim Body As CVtHttpaContentPart
     Set Body = s.request.Body
    
    
     VtLogInfo MODULE_NAME, PROC_NAME, "Body.Length=" & Body.ByteLength
     'VtLogInfo MODULE_NAME, PROC_NAME, "Body.Text=" & Body.GetText()
    
        
    Call RespondOK(s, CustomData)
    
'     Call s.Request.QueryParams.AddParam("statusCode", "404")
'     Call GetStatus(s, CustomData)

End Sub


Public Sub PostMultipart( _
    ByVal s As CVtHttpaSession, _
    ByVal CustomData As CWasCustomData)

    Const PROC_NAME As String = "PostMultipart"

    Dim Contents As CVtHttpaContents
    Dim ContentPart As CVtHttpaContentPart
    Dim i As Long

    Set Contents = s.request.Contents

    For i = 1 To Contents.Count

        Set ContentPart = Contents.GetPartAt(i)

        ' 여기서 각 multipart part 처리
        VtLogInfo MODULE_NAME, PROC_NAME, "Part #" & i
        VtLogInfo MODULE_NAME, PROC_NAME, "Content-Type: " & ContentPart.ContentType

    Next i

    Call RespondOK(s, CustomData)

End Sub


