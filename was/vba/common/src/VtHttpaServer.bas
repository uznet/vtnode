Attribute VB_Name = "VtHttpaServer"
Option Explicit

#If VBA7 And Win64 Then
  
#If VTNODE_DEBUG = 1 Then  ' Debug Mode

' Poll events from engine and enqueue them to corresponding context event queues.
' Return number of events dequeued.
Private Declare PtrSafe Function vtnode_httpa_server_event_dequeue Lib "vtnode64d.dll" ( _
    ByVal OutEvents As LongPtr, _
    ByVal Capacity As Long _
) As Long

' Send HTTP response.
' InstanceId identifies the HTTP connection.
' Response is the response handle created by vtnode_httpa_response_create_xxx().
' EOut is an optional UTF-8 error message buffer.
' ESize is the size of EOut in bytes.
' Return TRUE on success, FALSE on failure.
Private Declare PtrSafe Function vtnode_httpa_server_send_response Lib "vtnode64d.dll" ( _
    ByVal InstanceId As LongLong, _
    ByVal Response As LongPtr, _
    ByVal eout As LongPtr, _
    ByVal esize As Long _
) As Long



#Else ' Released Mode
    
' Poll events from engine and enqueue them to corresponding context event queues.
' Return number of events dequeued.
Private Declare PtrSafe Function vtnode_httpa_server_event_dequeue Lib "vtnode64.dll" ( _
    ByVal OutEvents As LongPtr, _
    ByVal Capacity As Long _
) As Long

' Send HTTP response.
' InstanceId identifies the HTTP connection.
' Response is the response handle created by vtnode_httpa_response_create_xxx().
' EOut is an optional UTF-8 error message buffer.
' ESize is the size of EOut in bytes.
' Return TRUE on success, FALSE on failure.
Private Declare PtrSafe Function vtnode_httpa_server_send_response Lib "vtnode64.dll" ( _
    ByVal InstanceId As LongLong, _
    ByVal Response As LongPtr, _
    ByVal eout As LongPtr, _
    ByVal esize As Long _
) As Long


#End If
    
#End If


''' <summary>
''' 서버 이벤트 큐에서 Event ID 하나를 가져옵니다.
''' 이벤트가 없거나 오류가 발생하면 0을 반환합니다.
''' </summary>
Public Function VtHttpaServerEventDequeue() As LongLong

    Dim EventId As LongLong
    Dim eventCount As Long

    On Error GoTo EH

    VtHttpaServerEventDequeue = 0
    EventId = 0

    eventCount = vtnode_httpa_server_event_dequeue(VarPtr(EventId), 1)

    If eventCount <= 0 Then
        Exit Function
    End If

    VtHttpaServerEventDequeue = EventId
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaServerEventDequeue] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaServerEventDequeue = 0

End Function



''' <summary>
''' HTTP Response를 클라이언트에 전송합니다.
''' </summary>
''' <param name="InstanceId">
''' HTTP 연결을 식별하는 Instance ID
''' </param>
''' <param name="Response">
''' VtHttpaResponseCreateXXX()로 생성한 Response 핸들
''' </param>
''' <param name="ErrorMessage">
''' 실패 시 Native 엔진이 반환한 UTF-8 오류 문자열
''' </param>
''' <returns>
''' 전송에 성공하면 True, 실패하면 False
''' </returns>
Public Function VtHttpaServerSendResponse( _
    ByVal InstanceId As LongLong, _
    ByVal Response As LongPtr, _
    ByRef ErrorMessage As String _
) As Boolean

    Const ERROR_BUFFER_SIZE As Long = 1024

    Dim ErrorBuffer(0 To ERROR_BUFFER_SIZE - 1) As Byte
    Dim nativeResult As Long

    On Error GoTo EH

    VtHttpaServerSendResponse = False
    ErrorMessage = vbNullString

    If InstanceId = 0 Then
        ErrorMessage = "Invalid HTTP instance ID."

        Debug.Print _
            "[VtHttpaServerSendResponse] " & _
            ErrorMessage

        Exit Function
    End If

    If Response = 0 Then
        ErrorMessage = "Invalid HTTP response handle."

        Debug.Print _
            "[VtHttpaServerSendResponse] " & _
            ErrorMessage

        Exit Function
    End If

    nativeResult = vtnode_httpa_server_send_response( _
            InstanceId, _
            Response, _
            VarPtr(ErrorBuffer(0)), _
            ERROR_BUFFER_SIZE)

    If nativeResult = 0 Then

        ErrorMessage = VtCharUtils.VtUtf8BytesToString(ErrorBuffer)

        If Len(ErrorMessage) = 0 Then
            ErrorMessage = "Failed to send HTTP response."
        End If

        Debug.Print _
            "[VtHttpaServerSendResponse] " & _
            ErrorMessage

        Exit Function
    End If

    VtHttpaServerSendResponse = True
    Exit Function

EH:
    ErrorMessage = _
        "VBA error " & Err.Number & ": " & Err.Description

    Debug.Print _
        "[VtHttpaServerSendResponse] " & _
        ErrorMessage

    VtHttpaServerSendResponse = False

End Function



