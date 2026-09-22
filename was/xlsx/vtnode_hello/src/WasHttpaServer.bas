Attribute VB_Name = "WasHttpaServer"
Option Explicit

Private Const MODULE_NAME As String = "WasHttpaServer"

'Private mFormVTNode As frmVTNode

Public Sub Init()

  ' Set mFormVTNode = frmVTNode
   
End Sub


Public Sub Uninit()

End Sub

Public Sub HandleAccepted(ByVal httpaSession As CVtHttpaSession)
   
    Const PROC_NAME As String = "HandleAccepted"
     
     
   Set httpaSession.CustomData = New CWasCustomData
   
   'Call VtLogInfo(MODULE_NAME, PROC_NAME, "Accepted->" & CStr(httpaSession.CID))

  
End Sub

Public Sub HandleDisconnected(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleDisconnected"
   
   'Call VtLogInfo(MODULE_NAME, PROC_NAME, "Disconnected->" & CStr(httpaSession.CID))


End Sub


''' <summary>
''' HTTP 트랜잭션이 시작되는 시점 (요청 수신 완료 및 파싱 직후)
''' </summary>
Public Sub HandleTransactionBegin(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleTransactionBegin"
    Dim Obj As Object
    Set Obj = httpaSession.CustomData
    
    If Not Obj Is Nothing Then
        Dim CustomData As CWasCustomData
        Set CustomData = Obj ' 이 순간부터 자동 완성(IntelliSense)이 살아납니다.
        
        CustomData.IsRequest = False
        
        ' 1. 정밀 시작 타임스탬프 기록 (자정 이후 경과한 초 단위를 소수점으로 반환)
        CustomData.BeginTime = Timer
        
       ' WasUILogger.OnBeginLog CustomData.RequestTimestamp
        
        
    
        'Call VtLogInfo(MODULE_NAME, PROC_NAME, "Begin Transaction->" & CStr(httpaSession.CID))
        
        
        
        ' 2. [추가] 시트 로깅 모듈을 호출하여 A열(시간), B열(URI) 기록 및 행 번호(RowIndex) 바인딩
        'Call UISheetLog.CreateRow(httpaSession, CustomData)
        
        'Debug.Print "Row=" & CustomData.SheetRowIndex
        
        
    End If
End Sub

''' <summary>
''' HTTP 트랜잭션이 완전히 종료되는 시점 (응답 전송 완료 및 세션 정리 직전)
''' </summary>
Public Sub HandleTransactionEnd(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleTransactionEnd"

    Dim Obj As Object
    Set Obj = httpaSession.CustomData
    
    If Not Obj Is Nothing Then
        Dim CustomData As CWasCustomData
        Set CustomData = Obj
        
        
        ' 1. 정밀 종료 타임스탬프 기록
        CustomData.EndTime = Timer
        
        If CustomData.IsRequest Then
'           WasUILogger.OnEndLog CustomData.RequestTimestamp, CustomData.ElapsedTime
        End If
        
        
        'Call VtLogInfo(MODULE_NAME, PROC_NAME, "End Transaction->" & CStr(httpaSession.CID))

        ' 3. [추가] 시트 로깅 모듈을 호출하여 해당 행의 C열(3번째 cell)에 최종 결과 기록
        'Call UISheetLog.SetResonseInfo(httpaSession, CustomData)
    End If
End Sub

Public Sub HandleRequest(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleRequest"
    Dim summary As String
    Dim lowPath As String
    Dim Obj As Object
    Dim CustomData As CWasCustomData
        
    Set CustomData = httpaSession.CustomData
    
    'WasAudio.Play AUDIO_REQUEST_RECEIVED
                    
                     
    summary = httpaSession.request.ToSummary()
                
    
    'Call WasHttpaStatus.WriteLog(summary)
    
    Call VtLogInfo(MODULE_NAME, PROC_NAME, "---------------------------------------------------------------------")
    Call VtLogInfo(MODULE_NAME, PROC_NAME, ">>> HTTP REQUEST:" & vbCrLf & summary)
    
                
    summary = httpaSession.request.Contents.ToSummary()
    Call VtLogInfo(MODULE_NAME, PROC_NAME, ">>> REQUEST CONTENTS:" & vbCrLf & summary)
                
                  
                  
    CustomData.IsRequest = True
'    WasUILogger.OnBeginLog CustomData.RequestTimestamp
'    WasUILogger.OnRequestLog CustomData.RequestTimestamp, httpaSession.socket, httpaSession.request
'
                       
    ' 대소문자 구분을 없애기 위해 LCase를 활용하는 것이 안전합니다.
    
    'Call UISheetLog.SetRequestInfo(httpaSession, CustomData)


    lowPath = LCase(httpaSession.request.path)

    'Debug.Print "request.path=" & lowPath
    
            
    Dim uriPath As String
    Dim fileName As String
            
    If ParseAsFilePath(lowPath, uriPath, fileName) Then
               
                If uriPath = "/" And Len(Trim$(fileName)) = 0 Then
                    fileName = "default.html"
                End If
 
                Call RespondFile(httpaSession, uriPath, fileName, CustomData)
    Else
            
               Call RespondNotFound(httpaSession, CustomData)
    End If
    
    ' route
'    Select Case lowPath
'
'
'
'        'for Test Only
'        Case "/test/get/status"
'            Call AppTestRouter.GetStatus(httpaSession, CustomData)
'
'        Case "/test/get/text"
'            Call AppTestRouter.GetText(httpaSession, CustomData)
'
'        Case "/test/get/file"
'            Call AppTestRouter.GetFile(httpaSession, CustomData)
'        Case "/test/get/multipart"
'
'            Call AppTestRouter.GetMultipart(httpaSession, CustomData)
'        Case "/test/post/text"
'            Call AppTestRouter.PostText(httpaSession, CustomData)
'
'        Case "/test/post/file"
'            Call AppTestRouter.PostFile(httpaSession, CustomData)
'        Case "/test/post/multipart"
'
'            Call AppTestRouter.PostMultipart(httpaSession, CustomData)
'
'        Case Else
'
'            Dim uriPath As String
'            Dim fileName As String
'
'            If ParseAsFilePath(lowPath, uriPath, fileName) Then
'
'                If uriPath = "/" And Len(Trim$(fileName)) = 0 Then
'                    fileName = "default.html"
'                End If
'
'                Call RespondFile(httpaSession, uriPath, fileName, CustomData)
'            Else
'
'               Call RespondNotFound(httpaSession, CustomData)
'            End If
'
'
'    End Select
'

End Sub

Public Function ParseAsFilePath( _
    ByVal path As String, _
    ByRef outPath As String, _
    ByRef outFilename As String) As Boolean

    Dim SlashPos As Long
    Dim DotPos As Long

    outPath = vbNullString
    outFilename = vbNullString
    ParseAsFilePath = False

    If Len(path) = 0 Then Exit Function

    ' --------------------------------------------------------
    ' Split Path / Filename
    ' --------------------------------------------------------
    SlashPos = InStrRev(path, "/")

    If SlashPos > 0 Then

        If SlashPos = 1 Then
            outPath = "/"
        Else
            outPath = Left$(path, SlashPos - 1)
        End If

        outFilename = Mid$(path, SlashPos + 1)

    Else

        outFilename = path

    End If

    ' --------------------------------------------------------
    ' Check Filename
    ' --------------------------------------------------------
    If Len(outFilename) = 0 Then
        
        ParseAsFilePath = True
        
        Exit Function
     End If
     

    ' --------------------------------------------------------
    ' Check Extension
    ' --------------------------------------------------------
    DotPos = InStrRev(outFilename, ".")

    If DotPos <= 1 Then Exit Function
    If DotPos >= Len(outFilename) Then Exit Function

    ParseAsFilePath = True

End Function

Public Sub RespondStatus(ByVal s As CVtHttpaSession, _
                         ByVal StatusCode As Long, _
                         ByVal CustomData As CWasCustomData)
    
    Const PROC_NAME As String = "RespondStatus"
                         
    Dim response As CVtHttpaResponse
    Dim SendRes As Boolean
    Dim SendError As String
    Dim rspMsg As String



    ' --------------------------------------------------------
    ' 1. Create Response
    ' --------------------------------------------------------
    Set response = VtHttpaFactory.VtHttpaCreateResponse()
        

    ' --------------------------------------------------------
    ' 2. Send 404 Not Found
    ' --------------------------------------------------------
    SendRes = response.SendStatus( _
                    s, _
                    StatusCode, _
                    SendError)


    ' --------------------------------------------------------
    ' 3. Log
    ' --------------------------------------------------------
    If SendRes Then

        rspMsg = "SendStatus(" & CStr(StatusCode) & "-> OK"

    Else

        rspMsg = "SendStatus(" & CStr(StatusCode) & "-> Failed." & SendError

    End If


    'Debug.Print rspMsg

    CustomData.ResponseStatus = rspMsg
    Call VtLogInfo(MODULE_NAME, PROC_NAME, rspMsg)

  '  WasUILogger.OnResponseLog CustomData.RequestTimestamp, response, rspMsg
    
    
    
    
    Set response = Nothing


                         
                         
End Sub
                         
                        
    
    

Public Sub RespondNotFound( _
    ByVal s As CVtHttpaSession, _
    ByVal CustomData As CWasCustomData)

    Const PROC_NAME As String = "RespondNotFound"
    Dim response As VTNodeRuntime.CVtHttpaResponse
    Dim SendRes As Boolean
    Dim SendError As String
    Dim rspMsg As String

    Const STATUS_NOT_FOUND As Long = 404


    ' --------------------------------------------------------
    ' 1. Create Response
    ' --------------------------------------------------------
    Set response = VTNodeRuntime.VtHttpaFactory.VtHttpaCreateResponse()
    
    

    ' --------------------------------------------------------
    ' 2. Send 404 Not Found
    ' --------------------------------------------------------
    SendRes = response.SendStatus( _
                    s, _
                    STATUS_NOT_FOUND, _
                    SendError)


    ' --------------------------------------------------------
    ' 3. Log
    ' --------------------------------------------------------
    If SendRes Then

        rspMsg = _
            "RespondNotFound: SendStatus(404) -> OK"

    Else

        rspMsg = _
            "RespondNotFound: SendStatus(404) -> Failed: " & _
            SendError

    End If


    'Debug.Print rspMsg


    CustomData.ResponseStatus = rspMsg
    Call VtLogInfo(MODULE_NAME, PROC_NAME, rspMsg)
    
   ' WasUILogger.OnResponseLog CustomData.RequestTimestamp, response, rspMsg

    Set response = Nothing

End Sub

Public Sub RespondBadRequest( _
    ByVal s As CVtHttpaSession, _
    ByVal CustomData As CWasCustomData, _
    Optional ByVal reason As String = vbNullString)

    Const PROC_NAME As String = "RespondBadRequest"
    Const STATUS_BAD_REQUEST As Long = 400
    Dim response As CVtHttpaResponse
    Dim SendRes As Boolean
    Dim SendError As String
    Dim rspMsg As String



    ' --------------------------------------------------------
    ' 1. Create Response
    ' --------------------------------------------------------
     Set response = VTNodeRuntime.VtHttpaFactory.VtHttpaCreateResponse()


    ' --------------------------------------------------------
    ' 2. Send 400 Bad Request
    ' --------------------------------------------------------
    SendRes = response.SendStatus( _
                    s, _
                    STATUS_BAD_REQUEST, _
                    SendError)


    ' --------------------------------------------------------
    ' 3. Log
    ' --------------------------------------------------------
    If SendRes Then

        rspMsg = _
            "RespondBadRequest: SendStatus(400) -> OK"

    Else

        rspMsg = _
            "RespondBadRequest: SendStatus(400) -> Failed: " & _
            SendError

    End If


    ' --------------------------------------------------------
    ' 4. Wasend Reason
    ' --------------------------------------------------------
    If Len(reason) > 0 Then
        rspMsg = rspMsg & " / Reason: " & reason
    End If


   ' Debug.Print rspMsg

    'Call UITabLog.WriteLog(rspMsg)

    CustomData.ResponseStatus = rspMsg
    
    Call VtLogInfo(MODULE_NAME, PROC_NAME, rspMsg)


   ' WasUILogger.OnResponseLog CustomData.RequestTimestamp, response, rspMsg

    Set response = Nothing

End Sub

Public Sub RespondOK( _
    ByVal s As CVtHttpaSession, _
    ByVal CustomData As CWasCustomData)

    Const PROC_NAME As String = "RespondOK"
    Dim response As CVtHttpaResponse
    Dim SendRes As Boolean
    Dim SendError As String
    Dim rspMsg As String

    Const STATUS_OK As Long = 200


    ' --------------------------------------------------------
    ' 1. Create Response
    ' --------------------------------------------------------
   Set response = VTNodeRuntime.VtHttpaFactory.VtHttpaCreateResponse()

    Call response.Headers.SetHeader("Connection", "close")
    

    ' --------------------------------------------------------
    ' 2. Send 200 OK
    ' --------------------------------------------------------
    SendRes = response.SendStatus( _
                    s, _
                    STATUS_OK, _
                    SendError)


    ' --------------------------------------------------------
    ' 3. Log
    ' --------------------------------------------------------
    If SendRes Then

        rspMsg = _
            "RespondOK: SendStatus(200) -> OK"

    Else

        rspMsg = _
            "RespondOK: SendStatus(200) -> Failed: " & _
            SendError

    End If


    'Debug.Print rspMsg


    CustomData.ResponseStatus = rspMsg
    Call VtLogInfo(MODULE_NAME, PROC_NAME, rspMsg)


   ' WasUILogger.OnResponseLog CustomData.RequestTimestamp, response, rspMsg
    Set response = Nothing

End Sub



Public Sub RespondJson( _
    ByVal s As CVtHttpaSession, _
    ByVal CustomData As CWasCustomData, _
    ByVal json As Object)

    Const PROC_NAME As String = "RespondJson"
    Dim response As CVtHttpaResponse
    Dim JsonText As String
    Dim SendRes As Boolean
    Dim SendError As String
    Dim rspMsg As String

    ' --------------------------------------------------------
    ' 1. JSON Object -> String
    ' --------------------------------------------------------
    JsonText = JsonConverter.ConvertToJson(json)

    ' --------------------------------------------------------
    ' 2. Create Response
    ' --------------------------------------------------------
  Set response = VTNodeRuntime.VtHttpaFactory.VtHttpaCreateResponse()

    response.StatusCode = 200

    Call response.Headers.SetHeader( _
        "Content-Type", _
        "application/json; charset=utf-8")

    ' --------------------------------------------------------
    ' 3. Send JSON
    ' --------------------------------------------------------
    SendError = vbNullString

    SendRes = response.SendText( _
        s, _
        JsonText, _
        SendError)

    ' --------------------------------------------------------
    ' 4. Log
    ' --------------------------------------------------------
    If SendRes Then

        rspMsg = "RespondJson: -> OK"

    Else

        rspMsg = "RespondJson: -> Failed: " & SendError

    End If


    'Call UITabLog.WriteLog(rspMsg)

    CustomData.ResponseStatus = rspMsg
    Call VtLogInfo(MODULE_NAME, PROC_NAME, rspMsg & "," & JsonText)
    
   ' WasUILogger.OnResponseLog CustomData.RequestTimestamp, response, rspMsg

    Set response = Nothing

End Sub

Public Function URIToLocalPath( _
    ByVal uriPath As String, _
    ByVal fileName As String) As String

    Dim LocalPath As String

    LocalPath = WasWWWRootDirectory()

    ' URI path -> local directory
    If Len(uriPath) > 0 And uriPath <> "/" Then

        If Left$(uriPath, 1) = "/" Then
            uriPath = Mid$(uriPath, 2)
        End If

        LocalPath = VtPathUtils.VtPathCombine( _
                        LocalPath, _
                        Replace(uriPath, "/", "\"))

    End If

    ' Filename
    LocalPath = VtPathUtils.VtPathCombine( _
                    LocalPath, _
                    fileName)

    URIToLocalPath = LocalPath

End Function
Public Sub RespondFile( _
    ByVal s As CVtHttpaSession, _
    ByVal uriPath As String, _
    ByVal fileName As String, _
    ByVal CustomData As CWasCustomData)
    
    Const PROC_NAME As String = "RespondFile"
    Dim LocalPath As String
    
    LocalPath = URIToLocalPath(uriPath, fileName)
    
    
    
    ' --------------------------------------------------------
    ' 3. Check File
    ' --------------------------------------------------------
    If dir$(LocalPath) = vbNullString Then

      
        Debug.Print _
            "RespondFile Error: File not found -> " & _
            LocalPath
        
        Call RespondStatus(s, 404, CustomData)
        

        Exit Sub

    End If
    
    
    Dim rspMsg As String
    Dim response As CVtHttpaResponse
    Dim SendError As String
    'Dim SendRes As Boolean
    
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
        GetContentTypeFromFilePath(LocalPath))

    ' --------------------------------------------------------
    ' 5. Send Response
    ' --------------------------------------------------------
    If response.SendFile( _
            s, _
            LocalPath, _
            SendError) Then

        rspMsg = _
            "Response.SendFile(" & _
            LocalPath & _
            ") -> Success"

'        Debug.Print _
'            "Send File Success: " & _
'            FullPath

    Else

        rspMsg = _
            "Response.SendFile(" & _
            LocalPath & _
            ") -> Failed: " & _
            SendError

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

   ' WasUILogger.OnResponseLog CustomData.RequestTimestamp, response, rspMsg
    Set response = Nothing

     
    
End Sub



Public Function GetContentTypeFromFilePath( _
    ByVal filePath As String) As String

    Dim ext As String

    ext = LCase$(VtPathUtils.VtPathGetExtension(filePath))

    If Left$(ext, 1) = "." Then
        ext = Mid$(ext, 2)
    End If

    Select Case ext

        Case "html", "htm"
            GetContentTypeFromFilePath = "text/html"

        Case "css"
            GetContentTypeFromFilePath = "text/css"

        Case "js"
            GetContentTypeFromFilePath = "text/javascript"

        Case "json"
            GetContentTypeFromFilePath = "application/json"

        Case "txt"
            GetContentTypeFromFilePath = "text/plain"

        Case "png"
            GetContentTypeFromFilePath = "image/png"

        Case "jpg", "jpeg"
            GetContentTypeFromFilePath = "image/jpeg"

        Case "gif"
            GetContentTypeFromFilePath = "image/gif"

        Case "svg"
            GetContentTypeFromFilePath = "image/svg+xml"

        Case "ico"
            GetContentTypeFromFilePath = "image/x-icon"

        Case Else
            GetContentTypeFromFilePath = "application/octet-stream"

    End Select

End Function
