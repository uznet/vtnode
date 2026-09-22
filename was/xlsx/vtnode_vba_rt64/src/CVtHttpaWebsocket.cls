VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaWebsocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' --- [Private Members] ---
Private mCid       As LongLong
Private mState     As VTHttpaWebsocketState      ' 현재 연결 상태 (0:Closed, 1:Connecting, 2:Establised)
Private mText       As String    ' 마지막으로 수신된 텍스트 메시지
Private mBinary     As Variant   ' 마지막으로 수신된 바이너리 메시지 (Byte 배열)
Private mProtocol   As String    ' 선택된 서브 프로토콜


Public Property Get state() As Long
   
   state = mState
End Property

Public Property Let state(ByVal wstate As VTHttpaWebsocketState)

    mState = wstate
    
End Property



' 마지막 수신 텍스트
Public Property Get LastText() As String

   LastText = mText
End Property

Public Property Let LastText(ByVal text As String)


 mText = text
End Property


' 마지막 수신 바이너리
Public Property Get LastBinary() As Variant
  LastBinary = mBinary

End Property
Public Property Let LastBinary(ByRef v As Variant)

' 1. 이전 바이너리 메모리를 확실히 비우기 (메모리 파편화 방지)
    If Not IsEmpty(mBinary) Then Erase mBinary
    
    ' 2. 새로운 데이터 대입 (배열 복사 발생)
    mBinary = v

End Property

' subprotocols
Public Property Get Protocol() As String

    Protocol = mProtocol
End Property
Public Property Let Protocol(ByVal proto As String)


    mProtocol = proto
End Property


' --- [Methods] ---

' 텍스트 전송
' 클래스 상단에 상수로 상태 정의 (아까 논의한 방식)

' 보완된 SendText
Public Function SendText(ByVal text As String) As Boolean
    ' 클래스가 이미 mUziid(또는 s.cid)를 보유하고 있다고 가정
    If mState = VTHTTPA_WEBSOCKET_STATE_CONNECTED And mCid <> 0 Then
        SendText = VtHttpaWsSendText(mCid, text)
        
    Else
        ' 로그나 디버그 창에 상태 오류 출력 가능
        Debug.Print "SendText Failed: 상태가 연결됨이 아니거나 ID가 없습니다."
        SendText = False
    End If
End Function

' 바이너리 전송
Public Function SendBinary(ByRef Data() As Byte) As Boolean
    ' 1. 상태 및 세션 ID 검증
    If (mState = VTHTTPA_WEBSOCKET_STATE_CONNECTED) And (mCid <> 0) Then
        
        ' 2. 유효하지 않은 배열(초기화되지 않음) 체크
        ' Not Not 구문은 VBA에서 초기화되지 않은 동적 배열을 체크하는 가장 빠른 테크닉입니다.
        If (Not Not Data) = 0 Then
            Debug.Print "SendBinary Skip: Array is empty."
            Exit Function
        End If

        ' 3. 엔진 호출 (C 엔진의 원시 포인터로 데이터가 전달됨)
        ' VtHttpaWsSendBinary 내부에서 SafeArray의 포인터를 직접 참조하도록 설계되어 있을 것입니다.
        SendBinary = VtHttpaWsSendBinary(mCid, Data)
        
    Else
        Debug.Print "SendBinary Failed [" & mCid & "]: Invalid State"
        SendBinary = False
    End If
End Function

' request upgrade  of client
Public Function OpenSession(ByVal s As CVtHttpaSession, _
                        ByVal Path As String, _
                        Optional ByVal Protocol As String = "", _
                        Optional ByVal Query As CVtHttpaQuery = Nothing, _
                        Optional ByVal extraHeaders As CVtHttpaHeaders = Nothing) As Boolean

    On Error GoTo EH

    Dim optionsId As LongLong: optionsId = 0
    Dim FullPath As String: FullPath = Path
    
    ' 1. 기본 방어 (Session Validation)
    If s Is Nothing Then GoTo CleanUp
    
    ' 클라이언트 역할 확인 (WebSocket 클라이언트는 Client 세션에서만 가능)
    If s.Role <> VTHTTPA_SESSION_ROLE_CLIENT Then
        Debug.Print "OpenSession() Skip: Not Client Role"
        GoTo CleanUp
    End If
    
    ' 중복 요청 방지
    If s.IsRequestSent Then
        Debug.Print "OpenSession() Skip: WebSocket upgrade already requested."
        GoTo CleanUp
    End If
    
    ' 트랜잭션 시작 가능 상태 확인
    If s.EventType <> VTHTTPA_EVENT_HT_BEGIN_TRANSACTION Then
        Debug.Print "OpenSession() Skip: Not ready for WS upgrade (Event: " & s.EventType & ")"
        GoTo CleanUp
    End If

    ' 2. Headers 처리 (Request Options 생성)
    If Not extraHeaders Is Nothing Then
        optionsId = extraHeaders.ToRequestOptionsId(s.CID)
        If optionsId = 0 Then GoTo CleanUp
    End If

    ' 3. Query String 결합
    If Not Query Is Nothing Then
        Dim qstring As String: qstring = Query.ToSummary()
        If Len(qstring) > 0 Then
            FullPath = FullPath & "?" & qstring
        End If
    End If

    ' 4. 엔진 호출 (WebSocket Upgrade Request)
    ' s.path 대신 인자로 받은 fullPath를 사용하여 유연성을 높입니다.
    s.IsRequestSent = True
    
    ' VtHttpaWsOpen 래퍼 함수는 내부적으로 path와 protocol을 UTF-8로 변환하여 호출해야 합니다.
    OpenSession = VtHttpaWsOpen(s.CID, FullPath, Protocol, optionsId)

CleanUp:
    ' 생성된 optionsId 핸들 해제
    If optionsId <> 0 Then
        Call VtHttpaRequestOptions_Free(optionsId)
        optionsId = 0
    End If
    
    ' 실패 시 다시 시도할 수 있도록 상태 롤백
    If Not OpenSession And Not s Is Nothing Then
        s.IsRequestSent = False
    End If
    Exit Function

EH:
    Debug.Print "OpenSession() Error: " & Err.Number & " - " & Err.Description
    OpenSession = False
    Resume CleanUp
End Function
' 연결 종료
Public Function CloseSession(Optional ByVal code As Long = 1000, Optional ByVal Reason As String = "") As Boolean
    ' 1. 현재 연결된 상태인지 확인
    If mState = VTHTTPA_WEBSOCKET_STATE_CONNECTED And mCid <> 0 Then
        
        ' 2. 엔진에 Close Frame 전송 요청
        CloseSession = VtHttpaWsSendClose(mCid, code, Reason)
        
        ' 3. [보완] 전송 성공 여부와 상관없이, 닫기 요청을 했으므로
        ' 내부 상태를 '종료 중' 또는 '닫힘'으로 변경하여 중복 호출 방지
        mState = VTHTTPA_WEBSOCKET_STATE_CLOSING ' (예: 상수값 3)
        
    End If
End Function

' --- [Events] ---
Private Sub Class_Initialize()
    mState = VTHTTPA_WEBSOCKET_STATE_CLOSED
    mText = vbNullString
    mBinary = Empty
    
End Sub

Private Sub Class_Terminate()

    mText = vbNullString
   If Not IsEmpty(mBinary) Then Erase mBinary
End Sub


' 1. 초기화 메서드를 통해 세션 ID를 딱 한 번만 주입받음
Public Sub Initialize(ByVal CID As LongLong)
    mCid = CID
End Sub
