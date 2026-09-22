Attribute VB_Name = "WasHttpaService"
'Attribute VB_Name = "WasHttpaService"

Option Explicit





' WasHttpaService.bas 같은 표준 모듈에 선언
Public Enum WasHttpaServiceStatus
    WAS_HTTPA_SERVICE_UNINITIALIZED = 0
    WAS_HTTPA_SERVICE_STARTING = 10
    WAS_HTTPA_SERVICE_RUNNING = 20
    WAS_HTTPA_SERVICE_STOPPING = 30
    WAS_HTTPA_SERVICE_STOPPED = 40
    WAS_HTTPA_SERVICE_FAULTED = 90
End Enum

Public Enum WasHttpaLoopStatus
    WAS_HTTPA_LOOP_IDLE = 0
    WAS_HTTPA_LOOP_SCHEDULED = 10
    WAS_HTTPA_LOOP_STARTING = 20
    WAS_HTTPA_LOOP_RUNNING = 30
    WAS_HTTPA_LOOP_STOP_REQUESTED = 40
    WAS_HTTPA_LOOP_STOPPING = 50
    WAS_HTTPA_LOOP_STOPPED = 60
    WAS_HTTPA_LOOP_FAULTED = 90
End Enum


Private Const MODULE_NAME As String = "WasHttpaService"


Private mHttpaService As CVtHttpaService
Private mHttpaStatus As WasHttpaServiceStatus
Private mLoopStatus As WasHttpaLoopStatus
Private mStopLoop As Boolean
Private mLastError As String


Private Sub PrintEventHeader(ByVal s As CVtHttpaSession)
    Const PROC_NAME As String = "PrintEventHeader"
    Dim ename As String
    ename = VTHttpaEventTypeToString(s.EventType)

    Dim laddr As String, raddr As String
    Dim lport As Long, rport As Long
    
    ' 고정 폭 출력을 위한 래핑 변수
    Dim remoteTarget As String
    Dim localTarget  As String

    If Not s.socket Is Nothing Then
        laddr = s.socket.LocalAddr
        lport = s.socket.LocalPort
        raddr = s.socket.RemoteAddr
        rport = s.socket.RemotePort
    Else
        laddr = "0.0.0.0"
        lport = 0
        raddr = "0.0.0.0"
        rport = 0
    End If

    ' [핵심] IP(15자)와 Port(5자)의 서식을 고정 폭으로 포맷팅 (총 21자 규격)
    ' 예: "192.168.0.4     :55207" 형태로 우측 공백 패딩 처리
    remoteTarget = Left(raddr & Space(15), 15) & ":" & Left(rport & Space(5), 5)
    localTarget = Left(laddr & Space(15), 15) & ":" & Left(lport & Space(5), 5)

    ' 이벤트 명 폭 고정 (24자)
    ename = Left(ename & Space(24), 24)

    ' CID 폭 고정 (8자 패딩하여 정렬 유지)
    Dim cidStr As String
    cidStr = Left("CID:" & Hex(s.CID) & Space(8), 8)

    ' 세션 역할(Role)에 따른 데이터 흐름 방향 정렬
    If s.Role = VTHTTPA_SESSION_ROLE_SERVER Then
        Call VtLogInfo(MODULE_NAME, PROC_NAME, _
                    cidStr & " | " & _
                    ename & " | " & _
                    remoteTarget & " -> " & localTarget)
    Else
        Call VtLogInfo(MODULE_NAME, PROC_NAME, _
                    cidStr & " | " & _
                    ename & " | " & _
                    localTarget & " -> " & remoteTarget)
    End If
End Sub

Private Sub OnHttpaEvent(ByVal httpaSession As CVtHttpaSession)
  
    Const PROC_NAME As String = "OnHttpaEvent"
    
    Call PrintEventHeader(httpaSession)
    
    Select Case httpaSession.EventType

        Case VTHTTPA_EVENT_ERROR
            Call WasHttpaError.HandleError(httpaSession)
            
        Case VTHTTPA_EVENT_NET_ERROR
            Call WasHttpaError.HandleNetError(httpaSession)
            
        Case VTHTTPA_EVENT_HT_ERROR ' HTTP Transaction Error
            Call WasHttpaError.HandleHTError(httpaSession)
        
        Case VTHTTPA_EVENT_HT_TIMEOUT, VTHTTPA_EVENT_WS_TIMEOUT
            Call WasHttpaError.HandleTimeout(httpaSession)
        
        Case VTHTTPA_EVENT_NET_ACCEPTED
            Call WasHttpaServer.HandleAccepted(httpaSession)
            
        Case VTHTTPA_EVENT_NET_CONNECTING    ' 클라이언트 접속 시작
            Call WasHttpaClient.HandleConnecting(httpaSession)
            
        Case VTHTTPA_EVENT_NET_CONNECT_FAILED ' 클라이언트 접속 실패
            Call WasHttpaClient.HandleConnectFailed(httpaSession)
            
        Case VTHTTPA_EVENT_NET_CONNECTED     ' 클라이언트 접속 성공
            Call WasHttpaClient.HandleConnected(httpaSession)
             
        Case VTHTTPA_EVENT_NET_DISCONNECTED
        
            If httpaSession.Role = VTHTTPA_SESSION_ROLE_CLIENT Then
                 Call WasHttpaClient.HandleDisconnected(httpaSession)
            Else
                 Call WasHttpaServer.HandleDisconnected(httpaSession)
            End If

        Case VTHTTPA_EVENT_HT_TRANSACTION_BEGIN ' 트랜잭션 시작 (헤더 파싱 완료)
           ' 클라이언트 모드이고, 사용자가 입력한 경로가 있을 때만 실행
            If httpaSession.Role = VTHTTPA_SESSION_ROLE_CLIENT Then
               Call WasHttpaClient.HandleTransactionBegin(httpaSession)
            Else
               Call WasHttpaServer.HandleTransactionBegin(httpaSession)
            
            End If
           
        Case VTHTTPA_EVENT_HT_TRANSACTION_END:   ' 트랜잭션 종료 (요청/응답 완료)
            If httpaSession.Role = VTHTTPA_SESSION_ROLE_CLIENT Then
               Call WasHttpaClient.HandleTransactionEnd(httpaSession)
            Else
               Call WasHttpaServer.HandleTransactionEnd(httpaSession)
            
            End If
        
        Case VTHTTPA_EVENT_HT_REQUEST_COMPLETED
            Call WasHttpaServer.HandleRequest(httpaSession)
            Call UpdateRuntimeStates
            
            
        Case VTHTTPA_EVENT_HT_RESPONSE_COMPLETED ' 응답 수신 완료 (Body 포함)
            Call WasHttpaClient.HandleResponse(httpaSession)
        Case VTHTTPA_EVENT_WS_ESTABLISHED
            Call WasHttpaWebSocket.HandleEstablised(httpaSession)
            
        Case VTHTTPA_EVENT_WS_TEXT_RECEIVED
            Call WasHttpaWebSocket.HandleReceivedText(httpaSession)
                     
        Case VTHTTPA_EVENT_WS_BINARY_RECEIVED
            Call WasHttpaWebSocket.HandleReceivedBinary(httpaSession)

        Case VTHTTPA_EVENT_WS_CLOSED
            Call WasHttpaWebSocket.HandleClosed(httpaSession)
            
        Case Else
            Call VtLogInfo(MODULE_NAME, PROC_NAME, "Unhandled Event Type: " & httpaSession.EventType)
    End Select
    

End Sub

Private Sub InitializeStatus()
    mHttpaStatus = WAS_HTTPA_SERVICE_UNINITIALIZED
    mLoopStatus = WAS_HTTPA_LOOP_IDLE
    mStopLoop = False
    mLastError = vbNullString
End Sub

Private Sub Init()

    If mHttpaService Is Nothing Then
        Set mHttpaService = _
            VTNodeRuntime.VtHttpaFactory.VtHttpaCreateService()
    End If
    
End Sub


Private Sub Uninit()
    
   If Not mHttpaService Is Nothing Then
       Set mHttpaService = Nothing
   End If
   
   
   
End Sub

Public Function WasHttpaServiceStatusToString( _
    ByVal status As WasHttpaServiceStatus) As String

    Select Case status
        Case WAS_HTTPA_SERVICE_UNINITIALIZED
            WasHttpaServiceStatusToString = "UNINITIALIZED"

        Case WAS_HTTPA_SERVICE_STARTING
            WasHttpaServiceStatusToString = "STARTING"

        Case WAS_HTTPA_SERVICE_RUNNING
            WasHttpaServiceStatusToString = "RUNNING"

        Case WAS_HTTPA_SERVICE_STOPPING
            WasHttpaServiceStatusToString = "STOPPING"

        Case WAS_HTTPA_SERVICE_STOPPED
            WasHttpaServiceStatusToString = "STOPPED"

        Case WAS_HTTPA_SERVICE_FAULTED
            WasHttpaServiceStatusToString = "FAULTED"

        Case Else
            WasHttpaServiceStatusToString = _
                "UNKNOWN(" & CStr(status) & ")"
    End Select
End Function
Public Function WasHttpaLoopStatusToString( _
    ByVal status As WasHttpaLoopStatus) As String

    Select Case status
        Case WAS_HTTPA_LOOP_IDLE
            WasHttpaLoopStatusToString = "IDLE"

        Case WAS_HTTPA_LOOP_SCHEDULED
            WasHttpaLoopStatusToString = "SCHEDULED"

        Case WAS_HTTPA_LOOP_STARTING
            WasHttpaLoopStatusToString = "STARTING"

        Case WAS_HTTPA_LOOP_RUNNING
            WasHttpaLoopStatusToString = "RUNNING"

        Case WAS_HTTPA_LOOP_STOP_REQUESTED
            WasHttpaLoopStatusToString = "STOP_REQUESTED"

        Case WAS_HTTPA_LOOP_STOPPING
            WasHttpaLoopStatusToString = "STOPPING"

        Case WAS_HTTPA_LOOP_STOPPED
            WasHttpaLoopStatusToString = "STOPPED"

        Case WAS_HTTPA_LOOP_FAULTED
            WasHttpaLoopStatusToString = "FAULTED"

        Case Else
            WasHttpaLoopStatusToString = _
                "UNKNOWN(" & CStr(status) & ")"
    End Select
End Function

Private Sub SetHttpaStatus( _
    ByVal newStatus As WasHttpaServiceStatus)

    If mHttpaStatus = newStatus Then Exit Sub

    VtLogInfo MODULE_NAME, _
              "SetHttpaStatus", _
              "Status changed: " _
              & WasHttpaServiceStatusToString(mHttpaStatus) _
              & " -> " _
              & WasHttpaServiceStatusToString(newStatus)

    mHttpaStatus = newStatus
End Sub
Private Sub SetLoopStatus( _
    ByVal newStatus As WasHttpaLoopStatus)

    If mLoopStatus = newStatus Then Exit Sub

    VtLogInfo MODULE_NAME, _
              "SetLoopStatus", _
              "Status changed: " _
              & WasHttpaLoopStatusToString(mLoopStatus) _
              & " -> " _
              & WasHttpaLoopStatusToString(newStatus)

    mLoopStatus = newStatus
End Sub

Public Sub TraceState(ByVal location As String)
    Debug.Print location, ThisWorkbook.FullName, _
                mHttpaStatus, mLoopStatus, _
                (mHttpaService Is Nothing)
End Sub
Public Property Get HttpaStatus() As WasHttpaServiceStatus
    HttpaStatus = mHttpaStatus
End Property

Public Property Get LoopStatus() As WasHttpaLoopStatus
    LoopStatus = mLoopStatus
End Property

Public Property Get LastError() As String
    LastError = mLastError
End Property

Public Property Get IsRunning() As Boolean

 If Not mHttpaService Is Nothing Then
   Debug.Print "IsRunning:not null"
   IsRunning = mHttpaService.IsRunning
 Else
   Debug.Print "IsRunning:null"
   IsRunning = False
 
 End If
 
'    IsRunning = _
'        (mHttpaStatus = WAS_HTTPA_SERVICE_RUNNING)
End Property

Public Property Get IsLoopRunning() As Boolean
    IsLoopRunning = _
        (mLoopStatus = WAS_HTTPA_LOOP_RUNNING)
End Property

Public Function RequiresForceTerminate() As Boolean
    Select Case mHttpaStatus
        Case WAS_HTTPA_SERVICE_FAULTED
            RequiresForceTerminate = True
            Exit Function

        Case WAS_HTTPA_SERVICE_RUNNING
            Select Case mLoopStatus
                Case WAS_HTTPA_LOOP_FAULTED, _
                     WAS_HTTPA_LOOP_STOPPED, _
                     WAS_HTTPA_LOOP_IDLE

                    RequiresForceTerminate = True
                    Exit Function
            End Select

        Case WAS_HTTPA_SERVICE_STOPPING
            If mLoopStatus = WAS_HTTPA_LOOP_FAULTED Then
                RequiresForceTerminate = True
                Exit Function
            End If
    End Select

    RequiresForceTerminate = False
End Function
Public Function GetStatusDescription() As String
    GetStatusDescription = _
        "Service Status: " _
        & WasHttpaServiceStatusToString(mHttpaStatus) _
        & vbCrLf _
        & "Loop Status: " _
        & WasHttpaLoopStatusToString(mLoopStatus) _
        & vbCrLf _
        & "Last Error: " _
        & IIf(Len(mLastError) = 0, "-", mLastError)
End Function

Public Function RequestStart( _
    ByVal vtnodeHomeDir As String, _
    ByVal vbaHomeDir As String, _
    ByVal configFile As String, _
    ByRef eout) As Boolean

    Const PROC_NAME As String = "RequestStart"

    Dim startResult As VTHttpaEngineStartResult
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    RequestStart = False
    eout = vbNullString
    mLastError = vbNullString

    Select Case mHttpaStatus
        Case WAS_HTTPA_SERVICE_RUNNING

            If mLoopStatus = WAS_HTTPA_LOOP_RUNNING _
               Or mLoopStatus = WAS_HTTPA_LOOP_SCHEDULED _
               Or mLoopStatus = WAS_HTTPA_LOOP_STARTING Then

                eout = "HTTPA service is already running."
                RequestStart = True
                Exit Function
            End If

            eout = "HTTPA service is running, but event loop is not available."
            Exit Function

        Case WAS_HTTPA_SERVICE_STARTING

            eout = "HTTPA service start is already in progress."
            RequestStart = True
            Exit Function

        Case WAS_HTTPA_SERVICE_STOPPING

            eout = "HTTPA service is stopping."
            Exit Function

        Case WAS_HTTPA_SERVICE_UNINITIALIZED, _
             WAS_HTTPA_SERVICE_STOPPED, _
             WAS_HTTPA_SERVICE_FAULTED

            ' 시작 가능

        Case Else

            eout = "RequestStart is not allowed. serviceStatus=" _
                 & WasHttpaServiceStatusToString(mHttpaStatus)

            Exit Function
    End Select

    If Len(Trim$(vtnodeHomeDir)) = 0 Then
        eout = "VTNode home directory is empty."
        Exit Function
    End If

    If Len(Trim$(vbaHomeDir)) = 0 Then
        eout = "VBA home directory is empty."
        Exit Function
    End If

    If Len(Trim$(configFile)) = 0 Then
        eout = "Configuration file path is empty."
        Exit Function
    End If

    If Len(dir$(configFile, vbNormal)) = 0 Then
        eout = "Configuration file not found. path=" & configFile
        Exit Function
    End If

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Start requested. vtnodeHomeDir=" & vtnodeHomeDir _
              & ", vbaHomeDir=" & vbaHomeDir _
              & ", configFile=" & configFile

    mHttpaStatus = WAS_HTTPA_SERVICE_STARTING
    mLoopStatus = WAS_HTTPA_LOOP_IDLE
    mStopLoop = False

    Init

    If mHttpaService Is Nothing Then
        mHttpaStatus = WAS_HTTPA_SERVICE_FAULTED
        mLastError = "CVtHttpaService instance creation failed."
        eout = mLastError

        VtLogError MODULE_NAME, PROC_NAME, mLastError
        Exit Function
    End If

    startResult = mHttpaService.Begin( _
                        vtnodeHomeDir, _
                        vbaHomeDir, _
                        configFile)

    If Not startResult.success Then
        mHttpaStatus = WAS_HTTPA_SERVICE_FAULTED
        mLoopStatus = WAS_HTTPA_LOOP_IDLE
        mLastError = startResult.emsg
        eout = startResult.emsg

        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "Start failed. reason=" & startResult.emsg

        Exit Function
    End If

    mHttpaStatus = WAS_HTTPA_SERVICE_RUNNING
    mLoopStatus = WAS_HTTPA_LOOP_SCHEDULED

    Application.OnTime _
        Now + TimeValue("00:00:01"), _
        "WasHttpaService.TimerForBegin"

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Start request accepted; event loop scheduled."

    RequestStart = True
    Exit Function

EH:
    errNumber = Err.Number
    errDescription = Err.description

    mHttpaStatus = WAS_HTTPA_SERVICE_FAULTED
    mLoopStatus = WAS_HTTPA_LOOP_FAULTED
    mLastError = errDescription
    eout = errDescription

    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "Unexpected error. number=" & CStr(errNumber) _
               & ", description=" & errDescription

    RequestStart = False
End Function

Public Function RequestStop(ByRef eout As String) As Boolean
    Const PROC_NAME As String = "RequestStop"

    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    RequestStop = False
    eout = vbNullString
    mLastError = vbNullString

    Select Case mHttpaStatus
        Case WAS_HTTPA_SERVICE_UNINITIALIZED, _
             WAS_HTTPA_SERVICE_STOPPED

            eout = "HTTPA service is already stopped."
            RequestStop = True
            Exit Function

        Case WAS_HTTPA_SERVICE_STARTING

            eout = "HTTPA service is starting. Stop cannot be requested yet."
            Exit Function

        Case WAS_HTTPA_SERVICE_STOPPING

            eout = "HTTPA service stop is already in progress."
            RequestStop = True
            Exit Function

        Case WAS_HTTPA_SERVICE_FAULTED

            If mHttpaService Is Nothing Then
                eout = "HTTPA service is faulted and no service instance exists."
                RequestStop = True
                Exit Function
            End If

        Case WAS_HTTPA_SERVICE_RUNNING

            ' 정상적인 중지 요청 처리

        Case Else

            eout = "RequestStop is not allowed. serviceStatus=" _
                 & WasHttpaServiceStatusToString(mHttpaStatus)

            Exit Function
    End Select

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Stop requested. loopStatus=" _
              & WasHttpaLoopStatusToString(mLoopStatus)

    Select Case mLoopStatus
        Case WAS_HTTPA_LOOP_SCHEDULED, _
             WAS_HTTPA_LOOP_STARTING, _
             WAS_HTTPA_LOOP_RUNNING

            mLoopStatus = WAS_HTTPA_LOOP_STOP_REQUESTED
            mStopLoop = True

            eout = "Stop request accepted."
            RequestStop = True

        Case WAS_HTTPA_LOOP_STOP_REQUESTED, _
             WAS_HTTPA_LOOP_STOPPING

            eout = "Event loop stop is already in progress."
            RequestStop = True

        Case WAS_HTTPA_LOOP_IDLE, _
             WAS_HTTPA_LOOP_STOPPED, _
             WAS_HTTPA_LOOP_FAULTED

            ' 실행 중인 Loop가 없으므로 Service 종료를 바로 예약한다.
            mHttpaStatus = WAS_HTTPA_SERVICE_STOPPING

            Application.OnTime _
                Now + TimeValue("00:00:01"), _
                "WasHttpaService.TimerForTerminate"

            eout = "Service termination scheduled."
            RequestStop = True

        Case Else

            eout = "Unknown event loop status=" & CStr(mLoopStatus)
            RequestStop = False
    End Select

    If RequestStop Then
        VtLogInfo MODULE_NAME, _
                  PROC_NAME, _
                  eout
    Else
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   eout
    End If

    Exit Function

EH:
    errNumber = Err.Number
    errDescription = Err.description

    mLastError = errDescription
    eout = errDescription

    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "Unexpected error. number=" & CStr(errNumber) _
               & ", description=" & errDescription

    RequestStop = False
End Function
    
Public Function ForceTerminate(ByRef errorOut As String) As Boolean
    Const PROC_NAME As String = "ForceTerminate"

    On Error GoTo EH

    ForceTerminate = False
    errorOut = vbNullString

    ' An active loop must unwind before its native service is released.
    Select Case mLoopStatus
        Case WAS_HTTPA_LOOP_STARTING, WAS_HTTPA_LOOP_RUNNING, _
             WAS_HTTPA_LOOP_STOP_REQUESTED, WAS_HTTPA_LOOP_STOPPING
            mStopLoop = True
            errorOut = "Event loop is still active. Retry after it has stopped."
            mLastError = errorOut
            Exit Function
    End Select

    mLastError = vbNullString
    mStopLoop = True
    mHttpaStatus = WAS_HTTPA_SERVICE_STOPPING

    ' Reuse the synchronous native termination and reference cleanup.
    TimerForTerminate

    If mHttpaStatus <> WAS_HTTPA_SERVICE_STOPPED Then
        errorOut = mLastError
        If Len(errorOut) = 0 Then
            errorOut = "HTTPA service termination failed."
            mLastError = errorOut
        End If
        Exit Function
    End If

    mLoopStatus = WAS_HTTPA_LOOP_STOPPED
    mLastError = vbNullString
    ForceTerminate = True
    Exit Function

EH:
    errorOut = "ForceTerminate failed. number=" & CStr(Err.Number) _
             & ", description=" & Err.description
    mLastError = errorOut
    mHttpaStatus = WAS_HTTPA_SERVICE_FAULTED
    On Error Resume Next
    VtLogError MODULE_NAME, PROC_NAME, errorOut
    On Error GoTo 0
End Function
    
Private Sub EnterHttpaLoop()
    Const PROC_NAME As String = "EnterHttpaLoop"

    Dim httpaSession As CVtHttpaSession
    Dim errNumber As Long
    Dim errDescription As String
    Dim loopFailed As Boolean

    On Error GoTo FATAL_ERROR

    '
    ' TimerForBegin에서 STARTING 상태로 변경한 후
    ' 이 함수가 호출되어야 한다.
    '
    If mHttpaStatus <> WAS_HTTPA_SERVICE_RUNNING Then
        Err.Raise _
            vbObjectError + 7201, _
            MODULE_NAME & "." & PROC_NAME, _
            "HTTPA service is not running."
    End If

    If mHttpaService Is Nothing Then
        Err.Raise _
            vbObjectError + 7202, _
            MODULE_NAME & "." & PROC_NAME, _
            "HTTPA service instance is not available."
    End If

    '
    ' RequestStop에서 이미 True로 설정했을 수 있으므로
    ' 여기에서 mStopLoop를 False로 초기화하면 안 된다.
    '
    If mStopLoop _
       Or mLoopStatus = WAS_HTTPA_LOOP_STOP_REQUESTED Then

        VtLogInfo MODULE_NAME, _
                  PROC_NAME, _
                  "Event loop start cancelled by stop request."

        mLoopStatus = WAS_HTTPA_LOOP_STOPPING
        GoTo NORMAL_UNINIT
    End If

    mLoopStatus = WAS_HTTPA_LOOP_STARTING

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Event loop initialization started."

    WasHttpaClient.Init
    WasHttpaError.Init
    WasHttpaNet.Init
    WasHttpaServer.Init
    WasHttpaWebSocket.Init

    WasAudio.Play AUDIO_WAS_STARTED

    mLoopStatus = WAS_HTTPA_LOOP_RUNNING

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Event loop started."

    WasHttpaService.UpdateRuntimeStates
    
    WasHttpaService.TraceState "Loop started"
    Do While Not mStopLoop
        Set httpaSession = _
            mHttpaService.PollServerEvent()

        If Not httpaSession Is Nothing Then
            OnHttpaEvent httpaSession
            Set httpaSession = Nothing
        Else
            VtWin32.Sleep 1
        End If

        '
        ' 다른 Workbook 및 RequestStop 호출이 실행될 수 있도록
        ' Excel 메시지 큐를 처리한다.
        '
        DoEvents
    Loop

    mLoopStatus = WAS_HTTPA_LOOP_STOPPING

WasHttpaService.TraceState "Loop stopped"
    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Event loop stopping."

NORMAL_UNINIT:

    '
    ' 하나의 Uninit에서 오류가 발생해도
    ' 나머지 정리 작업을 계속 실행한다.
    '
    On Error Resume Next

    Set httpaSession = Nothing

    WasHttpaWebSocket.Uninit
    WasHttpaServer.Uninit
    WasHttpaNet.Uninit
    WasHttpaError.Uninit
    WasHttpaClient.Uninit

    WasAudio.Play AUDIO_WAS_STOPPED

    On Error GoTo 0

    If Not loopFailed Then
        mLoopStatus = WAS_HTTPA_LOOP_STOPPED

        VtLogInfo MODULE_NAME, _
                  PROC_NAME, _
                  "Event loop stopped."
    End If

    '
    ' Event Loop가 완전히 반환된 다음
    ' Native HTTPA Service 종료를 예약한다.
    '
    On Error GoTo SCHEDULE_ERROR

    Application.OnTime _
        Now + TimeValue("00:00:01"), _
        "WasHttpaService.TimerForTerminate"

    Exit Sub

FATAL_ERROR:
    errNumber = Err.Number
    errDescription = Err.description

    loopFailed = True
    mLoopStatus = WAS_HTTPA_LOOP_FAULTED
    mLastError = errDescription

    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "Runtime error. number=" & CStr(errNumber) _
               & ", description=" & errDescription

    Resume NORMAL_UNINIT

SCHEDULE_ERROR:
    errNumber = Err.Number
    errDescription = Err.description

    mLastError = errDescription
    mHttpaStatus = WAS_HTTPA_SERVICE_FAULTED

    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "Failed to schedule TimerForTerminate. number=" _
               & CStr(errNumber) _
               & ", description=" & errDescription

    '
    ' 예약에 실패하면 Native Service를 현재 호출에서 직접 종료한다.
    '
    TimerForTerminate
End Sub


Private Sub PumpEvents()
    Dim session As CVtHttpaSession

    mScheduled = False
    If Not mRunning Then Exit Sub

    Do
        Set session = mHttpaService.PollServerEvent()
        If session Is Nothing Then Exit Do

        OnHttpaEvent session
        Set session = Nothing
    Loop

    If Not mRunning Then Exit Sub

    mNextRun = Now + TimeSerial(0, 0, 1)
    Application.OnTime mNextRun, _
                       "WasHttpaService.PumpEvents"
    mScheduled = True
End Sub



'Public Sub EnterHttpaLoop()
'
'    Dim httpaSession As CVtHttpaSession
'
'    mStopEventLoop = False
'
'
'    Call WasHttpaClient.Init(Nothing)
'    Call WasHttpaError.Init(Nothing)
'    Call WasHttpaNet.Init(Nothing)
'    Call WasHttpaServer.Init(Nothing)
'    Call WasHttpaWebSocket.Init
'
'
'
'    Do While Not mStopEventLoop
'
'        Set httpaSession = mHttpaService.PollServerEvent()
'
'        If Not httpaSession Is Nothing Then
'            Call OnHttpaEvent(httpaSession)
'
'        Else
'            ' 2. [★최적화] 이벤트 큐가 비어있을 때:
'            '    스핀 락(Spin Lock)으로 인한 CPU 100% 폭주 및 엑셀 프리징을 방지하기 위해
'            '    1ms 동안 스레드 컨텍스트 타임을 OS 윈도우 커널에 완전히 양보합니다.
'            VtWin32.Sleep 1
'
'        End If
'
'        DoEvents
'    Loop
'
'
'    Call WasHttpaWebSocket.Uninit
'    Call WasHttpaServer.Uninit
'    Call WasHttpaNet.Uninit
'    Call WasHttpaError.Uninit
'    Call WasHttpaClient.Uninit
'
'
'
'    Application.OnTime Now + TimeValue("00:00:01"), "WasHttpaService.TimerForTerminate"
'
'    Debug.Print "end timer"
'End Sub
' 표준 모듈(Module1 등)에 이 코드를 추가해 주세요.
Public Sub TimerForBegin()
    Const PROC_NAME As String = "TimerForBegin"

    On Error GoTo EH

    VtLogInfo MODULE_NAME, PROC_NAME, "TimerForBegin"

    If mLoopStatus = WAS_HTTPA_LOOP_STOP_REQUESTED _
       Or mStopLoop Then

        VtLogInfo MODULE_NAME, _
                  PROC_NAME, _
                  "Event loop start cancelled by stop request."

        mLoopStatus = WAS_HTTPA_LOOP_STOPPED
        mHttpaStatus = WAS_HTTPA_SERVICE_STOPPING

        Application.OnTime _
            Now + TimeValue("00:00:01"), _
            "WasHttpaService.TimerForTerminate"

        Exit Sub
    End If

    If mHttpaStatus <> WAS_HTTPA_SERVICE_RUNNING Then
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "Event loop cannot start. serviceStatus=" _
                   & WasHttpaServiceStatusToString(mHttpaStatus)

        Exit Sub
    End If

    If mLoopStatus <> WAS_HTTPA_LOOP_SCHEDULED Then
        Exit Sub
    End If

    mLoopStatus = WAS_HTTPA_LOOP_STARTING

    EnterHttpaLoop
    Exit Sub

EH:
    mLoopStatus = WAS_HTTPA_LOOP_FAULTED
    mLastError = Err.description

    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "TimerForBegin failed. number=" _
               & CStr(Err.Number) _
               & ", description=" & Err.description
End Sub

Public Sub TimerForTerminate()
    Const PROC_NAME As String = "TimerForTerminate"

    Dim errNumber As Long
    Dim errDescription As String
    Dim audioError As String

    On Error GoTo EH

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "TimerForTerminate started."

    If mHttpaStatus = WAS_HTTPA_SERVICE_UNINITIALIZED _
       Or mHttpaStatus = WAS_HTTPA_SERVICE_STOPPED Then

        VtLogInfo MODULE_NAME, _
                  PROC_NAME, _
                  "HTTPA service is already stopped."

        Exit Sub
    End If

    mHttpaStatus = WAS_HTTPA_SERVICE_STOPPING

    '
    ' 음성 재생 오류는 HTTPA Service 종료를 방해하지 않도록
    ' 별도로 처리한다.
    '
    On Error Resume Next

    WasAudio.WaitForComplete 5000

    If Err.Number <> 0 Then
        audioError = _
            "Audio wait failed. number=" & CStr(Err.Number) _
            & ", description=" & Err.description

        Err.Clear
    End If

    On Error GoTo EH

    If Len(audioError) > 0 Then
        VtLogError MODULE_NAME, PROC_NAME, audioError
    End If

    '
    ' 실제 HTTPA Engine 종료
    '
    If Not mHttpaService Is Nothing Then
        mHttpaService.Terminate
    End If

    '
    ' Service 객체 참조 해제
    '
    Uninit

    mHttpaStatus = WAS_HTTPA_SERVICE_STOPPED

    If mLoopStatus <> WAS_HTTPA_LOOP_FAULTED Then
        mLoopStatus = WAS_HTTPA_LOOP_STOPPED
    End If

    mStopLoop = False

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "HTTPA service stopped."

    WasHttpaService.UpdateRuntimeStates

    Exit Sub

EH:
    errNumber = Err.Number
    errDescription = Err.description

    mLastError = errDescription
    mHttpaStatus = WAS_HTTPA_SERVICE_FAULTED

    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "Terminate failed. number=" & CStr(errNumber) _
               & ", description=" & errDescription

    '
    ' Terminate에서 오류가 발생해도 VBA 객체 참조는 정리한다.
    '
    On Error Resume Next
    Uninit
    On Error GoTo 0
End Sub

'================================ App Runtime States ========================

Private Sub CollectAppInfo(ByRef info As VtRuntimeAppInfo)
    info.ApplicationName = ThisWorkbook.Name
    info.WorkbookPath = ThisWorkbook.FullName
End Sub

Private Sub CollectNetworkInfo(ByRef info As VtRuntimeNetworkInfo)


     Dim eout As String
     Dim r   As Boolean
          
     r = VtNetwork.VtNetGetLocalIp(info.LocalIp, eout)
     r = VtNetwork.VtNetGetExternalIp(info.ExternalIp, eout)
     
     
'    info.LocalIp = VtGetLocalIpAddress()
'    info.ExternalIp = GetExternalIpAddress()
End Sub
Private Function CollectRuntimeStatus() As VtRuntimeStatus

    '
    ' No service instance
    '
    If mHttpaService Is Nothing Then
        CollectRuntimeStatus = VT_RUNTIME_STOPPED
        Exit Function
    End If

    '
    ' Service instance exists and is running
    '
    If mHttpaService.IsRunning Then
        CollectRuntimeStatus = VT_RUNTIME_RUNNING
    Else
        CollectRuntimeStatus = VT_RUNTIME_STOPPED
    End If

End Function
'Public Type VtNodeRuntimeState
'    Status As VtRuntimeStatus
'    App As VtRuntimeAppInfo
'    network As VtRuntimeNetworkInfo
'    Http As VtRuntimeHttpInfo
'    Https As VtRuntimeHttpsInfo
'    Timing As VtRuntimeTimingInfo
'    Stats As VtRuntimeStatsInfo
'    LatestActivity As String
'    LastError As String
'End Type
Private Function CollectRuntimeState() As VtNodeRuntimeState

    Dim state As VtNodeRuntimeState

    state.status = CollectRuntimeStatus()

    Call CollectAppInfo(state.App)
    Call CollectNetworkInfo(state.network)

    If Not mHttpaService Is Nothing Then

        state.Http = _
            mHttpaService.GetHttpInfo(state.network)

        state.Https = _
            mHttpaService.GetHttpsInfo(state.network)

        state.Timing = _
            mHttpaService.GetTimingInfo()

        state.Stats = _
            mHttpaService.GetStatsInfo()

        state.LatestActivity = _
            mHttpaService.GetLatestActivity()

        state.LastError = _
            mHttpaService.GetLastError()

    End If

    CollectRuntimeState = state

End Function

Public Sub ShowRuntimeStates()

    Dim state As VtNodeRuntimeState

    On Error GoTo EH

    '
    ' Collect the current runtime state.
    '
    state = CollectRuntimeState()
    
    '
    ' Update the GET STARTED worksheet.
    '
    ShowRuntimeStatesOnSheet state


    '
    ' Update the console before showing it.
    '
    VtRuntimeUIControl.UpdateRuntimeState state

    '
    ' Show the Runtime Console modelessly.
    '
    VtRuntimeUIControl.ShowVtRuntimeUI

    Exit Sub

EH:
    Err.Raise Err.Number, _
              "ShowRuntimeStates", _
              Err.description

End Sub


Public Sub UpdateRuntimeStates()

    Dim state As VtNodeRuntimeState

    On Error GoTo EH

    '
    ' Collect once.
    '
    state = CollectRuntimeState()

    '
    ' Always update the worksheet.
    '
    ShowRuntimeStatesOnSheet state

    '
    ' Update the Runtime Console only when it is visible.
    '
    If VtRuntimeUIControl.IsVtRuntimeUIVisible Then
        VtRuntimeUIControl.UpdateRuntimeState state
    End If

    Exit Sub

EH:
    Debug.Print "[UpdateRuntimeStates] " _
              & Err.Number & " : " _
              & Err.description

End Sub


Private Sub ShowRuntimeStatesOnSheet( _
    ByRef state As VtNodeRuntimeState)

    Const SHEET_NAME As String = "VTNode WAS"

    Dim ws As Worksheet
    Dim statusText As String
    Dim serviceUrl As String
    Dim connectionCount As Long

    On Error GoTo EH

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    ' Update only the three runtime fields exposed by the VTNode WAS MVP shell.
    statusText = VtRuntimeStatusToString(state.status)

    If state.Http.enabled And Len(Trim$(state.Http.HttpUrl)) > 0 Then
        serviceUrl = Trim$(state.Http.HttpUrl)
    ElseIf state.Https.enabled And Len(Trim$(state.Https.HttpsUrl)) > 0 Then
        serviceUrl = Trim$(state.Https.HttpsUrl)
    Else
        serviceUrl = "-"
    End If

    WasUIController.UI_SetAccessUrl serviceUrl
    WasUIController.UI_SetNetwork state.Http, state.Https

    With ws.Range("F11:H12")
        .Cells(1, 1).Value2 = statusText

        Select Case state.status
            Case VT_RUNTIME_RUNNING
                .Font.Color = RGB(21, 128, 61)
            Case VT_RUNTIME_STARTING, VT_RUNTIME_STOPPING
                .Font.Color = RGB(180, 83, 9)
            Case VT_RUNTIME_FAULTED
                .Font.Color = RGB(220, 38, 38)
            Case Else
                .Font.Color = RGB(194, 65, 12)
        End Select
    End With

    Exit Sub

    statusText = VtRuntimeStatusToString(state.status)
    connectionCount = state.Stats.WebSocketClients

    If state.Http.enabled Then
        serviceUrl = Trim$(state.Http.HttpUrl)
    Else
        serviceUrl = vbNullString
    End If

    With ws

        '
        ' SERVER STATUS
        '
        .Range("B19").Value2 = _
            ChrW$(&H25CF) & "  " & statusText

        '
        ' SERVICE CHANNELS
        '
        If state.Http.enabled Then
            .Range("I19").Value2 = "ACTIVE"
        Else
            .Range("I19").Value2 = "INACTIVE"
        End If

        If state.Https.enabled Then
            .Range("I21").Value2 = "ENABLED"
        Else
            .Range("I21").Value2 = "DISABLED"
        End If

        If state.Http.enabled _
           And Len(Trim$(state.Http.WebSocketUrl)) > 0 Then

            .Range("I23").Value2 = "ENABLED"
        Else
            .Range("I23").Value2 = "DISABLED"
        End If

        '
        ' PORT / CONNECTIONS
        '
        If state.Http.enabled Then
            .Range("B22").Value2 = _
                "Port  " & CStr(state.Http.Port) _
                & "    " & ChrW$(&H2022) _
                & "    Connections  " _
                & CStr(connectionCount)
        Else
            .Range("B22").Value2 = _
                "Port  -" _
                & "    " & ChrW$(&H2022) _
                & "    Connections  " _
                & CStr(connectionCount)
        End If

        '
        ' HTTP URL
        '
        If Len(serviceUrl) > 0 Then
            .Range("D27").Value2 = serviceUrl
        Else
            .Range("D27").Value2 = "-"
        End If

    End With

    '
    ' Synchronize the QR image with the new D27 value.
    '
    SetQRCode _
        IsRunning:=state.Http.enabled, _
        serviceUrl:=CStr(ws.Range("D27").Value2)

    Exit Sub

EH:
    Err.Raise Err.Number, _
              "ShowRuntimeStatesOnSheet", _
              Err.description

End Sub

'

