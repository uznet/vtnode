VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private mSessions As Object  ' ? Dim → Private
Private mStartedAt As Date

Private mRequestCount As Long
Private mWebsocketClients   As Long



Private Const DEFAULT_MAX_ACTIVITY_COUNT As Long = 10
Private mLatestActivity() As String
Private mActivityCount As Long
Private mMaxActivityCount As Long
Private Sub Class_Initialize()
    
    
'    m_capacity = DEFAULT_CAPACITY
'    ReDim m_arr(0 To m_capacity - 1)
    
    
    Set mSessions = CreateObject("Scripting.Dictionary")
      
    SetMaxActivityCount DEFAULT_MAX_ACTIVITY_COUNT
    'Call VTNode.VtNodeLoadEngineDLL
    
    
'    If mSessions Is Nothing Then
'       Debug.Print " CreateObject(Scripting.Dictionary) failed"
'
'    Else
'       Debug.Print " CreateObject(Scripting.Dictionary) OK"
'
'    End If
       

End Sub

Private Sub Class_Terminate()
    If Not mSessions Is Nothing Then
        mSessions.RemoveAll
        Set mSessions = Nothing
    End If
End Sub


' start service
Public Function Begin(ByVal vtnodeHomeDir As String, ByVal vbaHomeDir As String, ByVal iniFile As String) As VTHttpaEngineStartResult

  
  Dim startResult As VTHttpaEngineStartResult
  
'  Dim initResult As VTNode.VTNodeInitReult
  
  
  If Me.IsRunning Then
    startResult.success = True
    startResult.emsg = "Already Running..."
    Begin = startResult
     
    Exit Function
         
  End If
    
     
  
  
    
   startResult.success = False
   startResult.emsg = ""
   
   
'    '1. init
'    initResult = VTNode.VtNodeInit()
'    If Not initResult.success Then
'
'       startResult.emsg = initResult.emsg
'
'       Begin = startResult
'
'       Exit Function
'
'    End If
'
'    ' 2. verify
'    Dim verifyResult As VTNode.VTNodeVerifyReult
'
'    verifyResult = VTNode.VtNodeVerify(vtnodeHomeDir)
'
'    If Not verifyResult.success Then
'
'       startResult.emsg = verifyResult.emsg
'       Begin = startResult
'       Exit Function
'    End If
    
  '3.start
  Begin = VtHttpaEngine.VtHttpaStart(vbaHomeDir, iniFile)
  
  If Begin.success Then
        mStartedAt = Now
        mRequestCount = 0
        mWebsocketClients = 0
  Else
    mStartedAt = 0
  
    mRequestCount = 0
    mWebsocketClients = 0
    
  End If
  
  
  
    
End Function


Public Sub Terminate()
   If IsRunning Then
      Call VtHttpaEngine.VtHttpaStop
      
'   Call VTNode.VtNodeUnInit
   End If
    
   mStartedAt = 0
  
 
    
End Sub



Public Property Get IsRunning() As Boolean

  'IsRunning = VtHttpaEngineIsRun()
  
  IsRunning = VtHttpaEngine.VtHttpaIsRun()
  
  
End Property


Private Sub Parse(ByVal s As CVtHttpaSession, ByVal eid As LongLong)
    Select Case s.EventType

        Case VTHTTPA_EVENT_ERROR, VTHTTPA_EVENT_NET_ERROR, VTHTTPA_EVENT_HT_ERROR
            s.ErrorMsg = VtHttpaEventGetErrorMessage(eid)  ' ? Let (자동)
            
        Case VTHTTPA_EVENT_HT_TIMEOUT, VTHTTPA_EVENT_WS_TIMEOUT
             s.TimeoutReason = VtHttpaEventGetTimeoutReason(eid)  ' ? Let (자동)
             
        
        Case VTHTTPA_EVENT_NET_ACCEPTED
        
            s.Role = VTHTTPA_SESSION_ROLE_SERVER
            Set s.Socket = New CVtHttpaSocket
            Call s.Socket.initFromEvent(eid)

        Case VTHTTPA_EVENT_NET_CONNECTING    ' 클라이언트 접속 시작
            s.Role = VTHTTPA_SESSION_ROLE_CLIENT

            Set s.Socket = New CVtHttpaSocket
            Call s.Socket.initFromEvent(eid)

        Case VTHTTPA_EVENT_NET_CONNECTED     ' 클라이언트 접속 성공
            Call s.Socket.initFromEvent(eid)

        Case VTHTTPA_EVENT_NET_CONNECT_FAILED ' 클라이언트 접속 실패
            s.ErrorMsg = VtHttpaEventGetErrorMessage(eid)  ' ? Let (자동)
            
        
             
        Case VTHTTPA_EVENT_NET_DISCONNECTED
            'RaiseEvent OnNetDisconnected(s)
            mSessions.Remove s.key
            

'        Case VTHTTPA_EVENT_NET_ERROR
'            s.ErrorMsg = VtHttpaEventGetErrorMessage(eid)  ' ? Let (자동)
'            RaiseEvent OnError(s)
'        Case VTHTTPA_EVENT_HT_ERROR
'            s.ErrorMsg = VtHttpaEventGetErrorMessage(eid)  ' ? Let (자동)
'            RaiseEvent OnError(s)

'        Case VTHTTPA_EVENT_HT_TIMEOUT
'            s.TimeoutReason = VtHttpaEventGetTimeoutReason(eid)  ' ? Let (자동)
'            RaiseEvent OnTimeout(s)

        Case VTHTTPA_EVENT_HT_TRANSACTION_BEGIN ' 트랜잭션 시작 (헤더 파싱 완료)
            'RaiseEvent OnHttpBeginTransaction(s)
            
            
            Call s.NewTransaction ' reinit Request,Response,Websco
            
                                   
            
        Case VTHTTPA_EVENT_HT_TRANSACTION_END:   ' 트랜잭션 종료 (요청/응답 완료)
            'RaiseEvent OnHttpEndTransaction(s)
        
        Case VTHTTPA_EVENT_HT_REQUEST_COMPLETED
            
            

            Set s.Request = New CVtHttpaRequest
                    
            If s.Request.initFromEvent(eid) Then
               mRequestCount = mRequestCount + 1
            
            ' Record only after request initialization succeeds.
               AddActivity_Request s.Request
            End If
                
            
    '

            'RaiseEvent OnHttpRequestReceived(s)
        Case VTHTTPA_EVENT_HT_RESPONSE_COMPLETED ' 응답 수신 완료 (Body 포함)
             Call s.Response.initFromEvent(eid)  ' ? 괄호 제거 또는 Call 추가
            'RaiseEvent OnHttpResponseReceived(s)
           
        
        Case VTHTTPA_EVENT_WS_ESTABLISHED
                   
            mWebsocketClients = mWebsocketClients + 1
            
            'Dim ws As CVtHttpaWebsocket
            Set s.Websocket = New CVtHttpaWebsocket
            Call s.Websocket.Initialize(s.CID)
            s.Websocket.state = VTHTTPA_WEBSOCKET_STATE_CONNECTED
            s.Websocket.Protocol = VtHttpaWsEventGetSelectedProtocol(eid)
                                                       
                        
            'RaiseEvent OnWsEstablished(s)

        Case VTHTTPA_EVENT_WS_TEXT_RECEIVED
            
            s.Websocket.LastText = VtHttpaWsEventGetText(eid)
            
            'RaiseEvent OnWsTextReceived(s)

        Case VTHTTPA_EVENT_WS_BINARY_RECEIVED
            s.Websocket.LastBinary = VtHttpaWsEventGetBinary(eid)
            'RaiseEvent OnWsBinaryReceived(s)

        Case VTHTTPA_EVENT_WS_CLOSED
            s.Websocket.state = VTHTTPA_WEBSOCKET_STATE_CLOSED
            
            'RaiseEvent OnWsClosed(s)

'        Case VTHTTPA_EVENT_WS_TIMEOUT
'            s.TimeoutReason = VtHttpaEventGetTimeoutReason(eid)  ' ? Let (자동)
'            RaiseEvent OnTimeout(s)

'        Case VTHTTPA_EVENT_WS_ERROR
'            s.ErrorMsg = VtHttpaEventGetErrorMessage(eid)  ' ? Let (자동)
'            RaiseEvent OnError(s)
    End Select
End Sub
Private Function GetSession(ByVal CID As LongLong) As CVtHttpaSession
    Dim s As CVtHttpaSession
    Dim key As String: key = CStr(CID)
    
    ' Dictionary(mSessions)를 사용한다고 가정
    If Not mSessions.Exists(key) Then
        ' 1. 새로운 세션 생성 및 초기화
        Set s = New CVtHttpaSession  ' 객체 생성 시 Set 필수
        s.Init CID, Me                  ' 내부 iid(또는 cid) 설정
        ' 2. 관리 대장에 추가
        mSessions.Add key, s
    Else
        ' 3. 기존에 있다면 꺼내오기
        Set s = mSessions(key)      ' 객체 할당 시 Set 필수
        
    End If
    
    ' 4. 함수 결과 반환 (객체이므로 Set 필수)
    Set GetSession = s
End Function
Public Function PollServerEvent() As CVtHttpaSession

         
    Dim eid As LongLong
    
    eid = VtHttpaServer.VtHttpaServerEventDequeue()
    
    If eid <> 0 Then
    
       Dim s As CVtHttpaSession
       Dim CID As LongLong
       
       CID = VtHttpaEventGetInstanceId(eid)
       
       Set s = GetSession(CID)
       
       s.EventId = eid
       s.EventType = VtHttpaEventGetType(eid)
     
      
       Call Parse(s, eid)
       Call VtHttpaEventFree(eid)
       
       Set PollServerEvent = s
    Else
    
       Set PollServerEvent = Nothing
    End If
    
            
End Function


Public Function AddSession(ByVal CID As LongLong) As CVtHttpaSession
    ' 내부적으로 getSession을 호출하여 객체를 생성하거나 기존 것을 반환
    Set AddSession = GetSession(CID)
End Function

Public Function Connect(ByVal serverAddr As String, _
                        ByVal serverPort As Long, _
                        ByVal secure As Boolean, _
                        ByVal timeoutSec As Long) As CVtHttpaSession

    Dim CID As LongLong
    Dim s As CVtHttpaSession
    
    ' 1. 로우레벨 엔진 접속 (C API 호출)
    ' 함수 호출 시에는 변수명만 전달합니다.
    CID = VtHttpaConnect(serverAddr, serverPort, secure, timeoutSec)
    
    ' 2. 접속 결과 확인 (cid가 0이면 실패)
    If CID <> 0 Then
        ' 3. 관리자(mApp)에 세션을 등록하고, 래핑된 세션 객체를 받아옴
        ' 앞서 만든 mApp.AddSession(cid)이 CVtHttpaSession을 반환하도록 설계되었습니다.
        Set s = AddSession(CID)
        s.Role = VTHTTPA_SESSION_ROLE_CLIENT
        
        Debug.Print "[VTNode] Connect Success. CID: " & CID
    Else
        Debug.Print "[VTNode] Connect Failed to " & serverAddr
        Set s = Nothing
    End If
    
    ' 4. 최종 결과 반환 (객체이므로 Set 사용)
    Set Connect = s

End Function

Public Sub Disconnect(ByVal s As CVtHttpaSession)
    On Error GoTo ErrHandler

    ' 1. 방어적 코드: 세션 객체가 유효한지 확인
    If s Is Nothing Then Exit Sub
    
    Dim CID As LongLong: CID = s.IID
    Dim key As String: key = s.key

    ' 2. 로우레벨 엔진에 연결 종료 명령 (C API 호출)
    If CID <> 0 Then
        ' 엔진 내부에서 소켓을 닫고 리소스를 정리하도록 호출
        Call VtHttpaDisconnect(CID)
        Debug.Print "[VTNode] Engine Disconnect Called for CID: " & CID
    End If

    ' 3. 세션 관리 장부(mSessions)에서 제거
    ' Dictionary를 사용 중이라면 Exists 확인 후 제거
'    If mSessions.Exists(key) Then
'        mSessions.Remove key
'        Debug.Print "[VTNode] Session removed from Manager: " & key
'    End If
    
    ' 4. 세션 객체 내부 상태 업데이트 및 해제
    's.IsConnected = False
    Set s = Nothing

    Exit Sub

ErrHandler:
    Debug.Print "!!! Error in Disconnect (CID: " & CID & "): " & Err.Description
    Resume Next
End Sub

' 2. (추가 제안) 관리 중인 모든 세션 정리
Public Sub DisconnectAll()
    Dim s As CVtHttpaSession
    For Each s In m_Sessions
        Call VtHttpaDisconnect(s.IID)
    Next s
    ' 컬렉션 비우기
    Set m_Sessions = New Collection
    Debug.Print "[VTNode] All managed sessions disconnected."
End Sub


' collect runtime states




Public Function GetHttpInfo(ByRef networkInfo As VtRuntimeNetworkInfo) As VtRuntimeHttpInfo

    Dim info As VtRuntimeHttpInfo
    Dim getResult As Boolean

    info.Enabled = False
    info.Port = 0
    info.HttpUrl = vbNullString
    info.WebSocketUrl = vbNullString

    getResult = VtHttpaGetHttpInfo(info.Port)

    If Not getResult Then GoTo Done
    If info.Port <= 0 Then GoTo Done

    info.Enabled = True

    If Len(Trim$(networkInfo.LocalIp)) = 0 Then
        GoTo Done
    End If

    info.HttpUrl = _
        "http://" & networkInfo.LocalIp _
        & ":" & CStr(info.Port)

    info.WebSocketUrl = _
        "ws://" & networkInfo.LocalIp _
        & ":" & CStr(info.Port)

Done:
    GetHttpInfo = info
End Function

Public Function GetHttpsInfo( _
    ByRef networkInfo As VtRuntimeNetworkInfo) _
    As VtRuntimeHttpsInfo

    Dim info As VtRuntimeHttpsInfo
    Dim certFile As String
    Dim keyFile As String
    Dim getResult As Boolean

    info.Enabled = False
    info.Port = 0
    info.HttpsUrl = vbNullString
    info.SecureWebSocketUrl = vbNullString

    getResult = VtHttpaGetHttpsInfo( _
                    info.Port, _
                    info.Enabled, _
                    certFile, _
                    keyFile)

    info.CertificateFile = certFile
    info.PrivateKeyFile = keyFile
    
    
    If Not getResult Then GoTo Done
    If Not info.Enabled Then GoTo Done

    If info.Port <= 0 Then
        info.Enabled = False
        GoTo Done
    End If

    If Len(Trim$(networkInfo.LocalIp)) = 0 Then
        GoTo Done
    End If

    info.HttpsUrl = _
        "https://" & networkInfo.LocalIp _
        & ":" & CStr(info.Port)

    info.SecureWebSocketUrl = _
        "wss://" & networkInfo.LocalIp _
        & ":" & CStr(info.Port)

Done:
    GetHttpsInfo = info
End Function

Public Function GetTimingInfo() As VtRuntimeTimingInfo

    Dim info As VtRuntimeTimingInfo

    info.StartedAt = mStartedAt
    info.UptimeSeconds = 0

    If info.StartedAt <> 0 Then
        info.UptimeSeconds = DateDiff( _
                                 "s", _
                                 info.StartedAt, _
                                 Now)

        If info.UptimeSeconds < 0 Then
            info.UptimeSeconds = 0
        End If
    End If

    GetTimingInfo = info
End Function

Public Function GetStatsInfo() As VtRuntimeStatsInfo

  Dim info As VtRuntimeStatsInfo
  
  info.TotalRequests = mRequestCount
  info.WebSocketClients = mWebsocketClients
  GetStatsInfo = info
  
  
End Function

Public Sub SetMaxActivityCount(ByVal value As Long)

    Dim newActivity() As String
    Dim keepCount As Long
    Dim index As Long

    If value < 1 Then
        Err.Raise vbObjectError + 1100, _
                  TypeName(Me) & ".SetMaxActivityCount", _
                  "Max activity count must be greater than zero."
    End If

    '
    ' Preserve the newest entries when resizing.
    '
    keepCount = mActivityCount

    If keepCount > value Then
        keepCount = value
    End If

    ReDim newActivity(0 To value - 1)

    For index = 0 To keepCount - 1
        newActivity(index) = mLatestActivity(index)
    Next index

    mLatestActivity = newActivity
    mActivityCount = keepCount
    mMaxActivityCount = value

End Sub
Public Sub ClearActivities()

    If mMaxActivityCount < 1 Then
        mMaxActivityCount = DEFAULT_MAX_ACTIVITY_COUNT
    End If

    ReDim mLatestActivity(0 To mMaxActivityCount - 1)
    mActivityCount = 0

End Sub
Public Sub AddActivity(ByVal activityText As String)

    Dim index As Long

    activityText = Trim$(activityText)

    If Len(activityText) = 0 Then Exit Sub

    If mMaxActivityCount < 1 Then
        SetMaxActivityCount DEFAULT_MAX_ACTIVITY_COUNT
    End If

    '
    ' Shift older entries toward the end.
    '
    For index = mMaxActivityCount - 1 To 1 Step -1
        mLatestActivity(index) = mLatestActivity(index - 1)
    Next index

    '
    ' The newest activity is always stored at index 0.
    '
    mLatestActivity(0) = activityText

    If mActivityCount < mMaxActivityCount Then
        mActivityCount = mActivityCount + 1
    End If

End Sub
Public Function GetLatestActivity() As String

    Dim index As Long
    Dim result As String

    For index = 0 To mActivityCount - 1

        If Len(result) > 0 Then
            result = result & vbCrLf
        End If

        result = result & mLatestActivity(index)

    Next index

    GetLatestActivity = result

End Function


Public Sub AddActivity_Request(ByVal Request As CVtHttpaRequest)

    Debug.Print "AddActivity_Request"
    
    Dim activityText As String

    If Request Is Nothing Then Exit Sub

    activityText = Format$(Now, "hh:nn:ss") _
                 & "  " & VTHttpaMethodToString(Request.Method) _
                 & " " & Request.Path

    AddActivity activityText

End Sub
Public Sub AddActivity_Response( _
    ByVal Response As CVtHttpaResponse, _
    ByVal responseType As VtResponseActivityType)

    Debug.Print "AddActivity_Response"
    
    
    Dim activityText As String
    Dim typeText As String

    If Response Is Nothing Then Exit Sub

    Select Case responseType

        Case VT_RESPONSE_STATUS
            typeText = "STATUS"

        Case VT_RESPONSE_TEXT
            typeText = "TEXT"

        Case VT_RESPONSE_FILE
            typeText = "FILE"

        Case VT_RESPONSE_MULTIPART
            typeText = "MULTIPART"

        Case Else
            typeText = "UNKNOWN"

    End Select

    activityText = Format$(Now, "hh:nn:ss") _
                 & "  " & CStr(Response.StatusCode) _
                 & "  [" & typeText & "]"

    AddActivity activityText

End Sub

Public Function GetLastError() As String
  
  GetLastError = "Last Error"
  
   
End Function
