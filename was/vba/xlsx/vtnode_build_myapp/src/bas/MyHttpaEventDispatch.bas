Attribute VB_Name = "MyHttpaEventDispatch"
Option Explicit

Private Const MODULE_NAME As String = "MyHttpaEventDispatch"




Private Sub PrintEventHeader(ByVal s As CVtHttpaSession)
    Const PROC_NAME As String = "PrintEventHeader"
    Dim ename As String
    ename = VTHttpaEventTypeToString(s.EventType)

    Dim laddr As String, raddr As String
    Dim lport As Long, rport As Long

    ' 고정 폭 출력을 위한 래핑 변수
    Dim remoteTarget As String
    Dim localTarget As String

    If Not s.Socket Is Nothing Then
        laddr = s.Socket.LocalAddr
        lport = s.Socket.LocalPort
        raddr = s.Socket.RemoteAddr
        rport = s.Socket.RemotePort
    Else
        laddr = "0.0.0.0"
        lport = 0
        raddr = "0.0.0.0"
        rport = 0
    End If

    ' [핵심] IP(15자)와 Port(5자)의 서식을 고정 폭으로 포맷팅 (총 21자 규격)
    ' 예: "192.168.0.4     :55207" 형태로 우측 공백 패딩 처리
    remoteTarget = left(raddr & Space(15), 15) & ":" & left(rport & Space(5), 5)
    localTarget = left(laddr & Space(15), 15) & ":" & left(lport & Space(5), 5)

    ' 이벤트 명 폭 고정 (24자)
    ename = left(ename & Space(24), 24)

    ' CID 폭 고정 (8자 패딩하여 정렬 유지)
    Dim cidStr As String
    cidStr = left("CID:" & Hex(s.CID) & Space(8), 8)

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

Public Sub DispatchHttpaEvent(ByVal httpaSession As CVtHttpaSession)

    Const PROC_NAME As String = "DispatchHttpaEvent"

    Call PrintEventHeader(httpaSession)

    Select Case httpaSession.EventType

        Case VTHTTPA_EVENT_ERROR
            Call MyHttpaErrorEvent.HandleError(httpaSession)

        Case VTHTTPA_EVENT_NET_ERROR
            Call MyHttpaErrorEvent.HandleNetError(httpaSession)

        Case VTHTTPA_EVENT_HT_ERROR ' HTTP Transaction Error
            Call MyHttpaErrorEvent.HandleHTError(httpaSession)

        Case VTHTTPA_EVENT_HT_TIMEOUT, VTHTTPA_EVENT_WS_TIMEOUT
            Call MyHttpaErrorEvent.HandleTimeout(httpaSession)

        Case VTHTTPA_EVENT_NET_ACCEPTED
            Call MyHttpaServerEvent.HandleAccepted(httpaSession)

        Case VTHTTPA_EVENT_NET_CONNECTING    ' 클라이언트 접속 시작
            Call MyHttpaClientEvent.HandleConnecting(httpaSession)

        Case VTHTTPA_EVENT_NET_CONNECT_FAILED ' 클라이언트 접속 실패
            Call MyHttpaClientEvent.HandleConnectFailed(httpaSession)

        Case VTHTTPA_EVENT_NET_CONNECTED     ' 클라이언트 접속 성공
            Call MyHttpaClientEvent.HandleConnected(httpaSession)

        Case VTHTTPA_EVENT_NET_DISCONNECTED

            If httpaSession.Role = VTHTTPA_SESSION_ROLE_CLIENT Then
                Call MyHttpaClientEvent.HandleDisconnected(httpaSession)
            Else
                Call MyHttpaServerEvent.HandleDisconnected(httpaSession)
            End If

        Case VTHTTPA_EVENT_HT_TRANSACTION_BEGIN ' 트랜잭션 시작 (헤더 파싱 완료)
            ' 클라이언트 모드이고, 사용자가 입력한 경로가 있을 때만 실행
            If httpaSession.Role = VTHTTPA_SESSION_ROLE_CLIENT Then
                Call MyHttpaClientEvent.HandleTransactionBegin(httpaSession)
            Else
                Call MyHttpaServerEvent.HandleTransactionBegin(httpaSession)

            End If

        Case VTHTTPA_EVENT_HT_TRANSACTION_END   ' 트랜잭션 종료 (요청/응답 완료)
            If httpaSession.Role = VTHTTPA_SESSION_ROLE_CLIENT Then
                Call MyHttpaClientEvent.HandleTransactionEnd(httpaSession)
            Else
                Call MyHttpaServerEvent.HandleTransactionEnd(httpaSession)

            End If

        Case VTHTTPA_EVENT_HT_REQUEST_COMPLETED
            Call MyHttpaServerEvent.HandleRequest(httpaSession)
            Call UpdateRuntimeStates


        Case VTHTTPA_EVENT_HT_RESPONSE_COMPLETED ' 응답 수신 완료 (Body 포함)
            Call MyHttpaClientEvent.HandleResponse(httpaSession)
        Case VTHTTPA_EVENT_WS_ESTABLISHED
            Call MyHttpaWebSocketEvent.HandleEstablised(httpaSession)

        Case VTHTTPA_EVENT_WS_TEXT_RECEIVED
            Call MyHttpaWebSocketEvent.HandleReceivedText(httpaSession)

        Case VTHTTPA_EVENT_WS_BINARY_RECEIVED
            Call MyHttpaWebSocketEvent.HandleReceivedBinary(httpaSession)

        Case VTHTTPA_EVENT_WS_CLOSED
            Call MyHttpaWebSocketEvent.HandleClosed(httpaSession)

        Case Else
            Call VtLogInfo(MODULE_NAME, PROC_NAME, "Unhandled Event Type: " & httpaSession.EventType)
    End Select


End Sub
