Attribute VB_Name = "VtHttpaWebsocket"


Option Explicit

#If VBA7 And Win64 Then

#If VTNODE_DEBUG = 1 Then

'============================================================
' WEBSOCKET EVENT ACCESS
'============================================================

' Get selected subprotocol from WS_ESTABLISHED event.
' Returned string is read-only UTF-8 and valid until event is freed.
Private Declare PtrSafe Function vtnode_websocket_get_selected_protocol_from_event Lib "vtnode64d.dll" ( _
    ByVal EventId As LongLong _
) As LongPtr

' Get UTF-8 text payload from WS_TEXT event.
' Returned text is read-only UTF-8 and valid until event is freed.
Private Declare PtrSafe Function vtnode_websocket_get_utf8text_from_event Lib "vtnode64d.dll" ( _
    ByVal EventId As LongLong, _
    ByRef text As LongPtr, _
    ByRef TextLen As Long _
) As Long

' Get binary payload from WS_BINARY event.
' Returned data is read-only and valid until event is freed.
Private Declare PtrSafe Function vtnode_websocket_get_binary_from_event Lib "vtnode64d.dll" ( _
    ByVal EventId As LongLong, _
    ByRef Data As LongPtr, _
    ByRef DataLen As Long _
) As Long


'============================================================
' WEBSOCKET SEND / CLOSE
'============================================================

' Send UTF-8 text message.
' TextLen < 0 means strlen(Text).
' Valid only after websocket is established.
Private Declare PtrSafe Function vtnode_websocket_send_utf8text Lib "vtnode64d.dll" ( _
    ByVal InstanceId As LongLong, _
    ByVal text As LongPtr, _
    ByVal TextLen As Long _
) As Long

' Send binary message.
' Valid only after websocket is established.
Private Declare PtrSafe Function vtnode_websocket_send_binary Lib "vtnode64d.dll" ( _
    ByVal InstanceId As LongLong, _
    ByVal Data As LongPtr, _
    ByVal DataLen As Long _
) As Long

' Start websocket close handshake.
' CloseCode : Standard websocket close code such as 1000.
' Reason    : Optional UTF-8 string.
' ReasonLen < 0 means strlen(Reason).
' WS_CLOSED event will be generated when close is completed.
Private Declare PtrSafe Function vtnode_websocket_send_close Lib "vtnode64d.dll" ( _
    ByVal InstanceId As LongLong, _
    ByVal CloseCode As Long, _
    ByVal Reason As LongPtr, _
    ByVal ReasonLen As Long _
) As Long


#Else  ' Released Mode



' Get selected subprotocol from WS_ESTABLISHED event.
' Returned string is read-only UTF-8 and valid until event is freed.
Private Declare PtrSafe Function vtnode_websocket_get_selected_protocol_from_event Lib "vtnode64.dll" ( _
    ByVal EventId As LongLong _
) As LongPtr

' Get UTF-8 text payload from WS_TEXT event.
' Returned text is read-only UTF-8 and valid until event is freed.
Private Declare PtrSafe Function vtnode_websocket_get_utf8text_from_event Lib "vtnode64.dll" ( _
    ByVal EventId As LongLong, _
    ByRef text As LongPtr, _
    ByRef TextLen As Long _
) As Long

' Get binary payload from WS_BINARY event.
' Returned data is read-only and valid until event is freed.
Private Declare PtrSafe Function vtnode_websocket_get_binary_from_event Lib "vtnode64.dll" ( _
    ByVal EventId As LongLong, _
    ByRef Data As LongPtr, _
    ByRef DataLen As Long _
) As Long


'============================================================
' WEBSOCKET SEND / CLOSE
'============================================================

' Send UTF-8 text message.
' TextLen < 0 means strlen(Text).
' Valid only after websocket is established.
Private Declare PtrSafe Function vtnode_websocket_send_utf8text Lib "vtnode64.dll" ( _
    ByVal InstanceId As LongLong, _
    ByVal text As LongPtr, _
    ByVal TextLen As Long _
) As Long

' Send binary message.
' Valid only after websocket is established.
Private Declare PtrSafe Function vtnode_websocket_send_binary Lib "vtnode64.dll" ( _
    ByVal InstanceId As LongLong, _
    ByVal Data As LongPtr, _
    ByVal DataLen As Long _
) As Long

' Start websocket close handshake.
' CloseCode : Standard websocket close code such as 1000.
' Reason    : Optional UTF-8 string.
' ReasonLen < 0 means strlen(Reason).
' WS_CLOSED event will be generated when close is completed.
Private Declare PtrSafe Function vtnode_websocket_send_close Lib "vtnode64.dll" ( _
    ByVal InstanceId As LongLong, _
    ByVal CloseCode As Long, _
    ByVal Reason As LongPtr, _
    ByVal ReasonLen As Long _
) As Long
#End If
#End If





Public Function VtHttpaWsOpen(ByVal CID As LongLong, _
                             ByVal Path As String, _
                             ByVal protocols As String, _
                             ByVal optionsId As LongLong) As Boolean
    Dim uriUtf8() As Byte
    Dim protocolsUtf8() As Byte
    Dim uriUtf8Ptr As LongPtr: uriUtf8Ptr = 0
    Dim protocolsUtf8Ptr As LongPtr: protocolsUtf8Ptr = 0

    ' 1. URI 변환 및 포인터 획득 (URI는 필수인 경우가 많음)
    uriUtf8 = VtStringToUtf8Bytes(Path)
    If (Not Not uriUtf8) Then ' 배열이 초기화되었는지 확인
        uriUtf8Ptr = VarPtr(uriUtf8(LBound(uriUtf8)))
    End If

    ' 2. Protocols 변환 (선택 사항일 수 있으므로 Null 체크)
    If Len(protocols) > 0 Then
        protocolsUtf8 = VtStringToUtf8Bytes(protocols)
        If (Not Not protocolsUtf8) Then
            protocolsUtf8Ptr = VarPtr(protocolsUtf8(LBound(protocolsUtf8)))
        End If
    End If

    ' 3. 엔진 호출
    ' vtnode_websocket_open이 내부적으로 문자열을 복사(Copy)해서 사용한다면
    ' 이 함수가 종료되어 uriUtf8 배열이 소멸되어도 안전합니다.
    VtHttpaWsOpen = (vtnode_websocket_open(CID, uriUtf8Ptr, protocolsUtf8Ptr, optionsId) <> 0)

    
End Function

Public Function VtHttpaWsSendText(ByVal CID As LongLong, ByVal text As String) As Boolean
    Dim textUtf8() As Byte
    Dim textUtf8Ptr As LongPtr
    Dim TextLen As Long

    ' 1. 변환: VtStringToUtf8Bytes가 끝에 vbNullChar를 붙여서 반환한다고 가정
    textUtf8 = VtStringToUtf8Bytes(text)
    
    ' 2. 유효성 체크
    If (Not Not textUtf8) = 0 Then
       VtHttpaWsSendText = False
       Exit Function
    End If

    ' 3. 포인터 획득
    textUtf8Ptr = VarPtr(textUtf8(0))
    
    ' 4. 정확한 길이 계산
    ' UBound + 1은 Null을 포함한 전체 바이트 수입니다.
    ' WebSocket 텍스트 프레임은 보통 Null을 제외한 순수 UTF-8 데이터만 실어 보냅니다.
    TextLen = (UBound(textUtf8) - LBound(textUtf8) + 1) - 1
    
    ' 데이터가 아예 없는 경우(빈 문자열) 엔진 호출 방지
    If TextLen > 0 Then
      ' 5. 엔진 호출
      VtHttpaWsSendText = (vtnode_websocket_send_utf8text(CID, textUtf8Ptr, TextLen) <> 0)
    Else
      VtHttpaWsSendText = False
      
    End If
End Function

Public Function VtHttpaWsSendBinary(ByVal CID As LongLong, ByRef Data() As Byte) As Boolean
    Dim DataPtr As LongPtr
    Dim DataLen As Long
    
    ' 1. 유효성 체크 (배열이 비어있는지 확인)
    On Error Resume Next
    DataLen = UBound(Data) - LBound(Data) + 1
    If Err.Number <> 0 Or DataLen <= 0 Then
        VtHttpaWsSendBinary = False
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' 2. 포인터 획득 (LBound를 사용하여 시작 지점 유연하게 대응)
    DataPtr = VarPtr(Data(LBound(Data)))

    ' 3. 엔진 전송 호출
    ' 바이너리는 널 문자가 없으므로 계산된 전체 길이를 그대로 보냅니다.
    VtHttpaWsSendBinary = (vtnode_websocket_send_binary(CID, DataPtr, DataLen) <> 0)
End Function

Public Function VtHttpaWsSendClose(ByVal CID As LongLong, _
                                  ByVal CloseCode As Long, _
                                  ByVal Reason As String) As Boolean
    Dim reasonUtf8() As Byte
    Dim reasonUtf8Ptr As LongPtr: reasonUtf8Ptr = 0
    Dim ReasonLen As Long: ReasonLen = 0

    ' 1. 사유(Reason)가 있을 때만 UTF-8 변환
    If Len(Reason) > 0 Then
        reasonUtf8 = VtStringToUtf8Bytes(Reason)
        If (Not Not reasonUtf8) Then
            reasonUtf8Ptr = VarPtr(reasonUtf8(0))
            ' Null Terminated이므로 실제 데이터 길이는 전체 크기 - 1
            ReasonLen = (UBound(reasonUtf8) - LBound(reasonUtf8) + 1) - 1
        End If
    End If

    ' 2. 엔진 호출
    ' closeCode: 보통 정상 종료는 1000, 비정상 종료는 1001 등을 사용
    VtHttpaWsSendClose = (vtnode_websocket_send_close(CID, CloseCode, reasonUtf8Ptr, ReasonLen) <> 0)
    
End Function

Public Function VtHttpaWsEventGetSelectedProtocol(ByVal eid As LongLong) As String
    Dim protocolPtr As LongPtr
    
    ' 1. 엔진으로부터 프로토콜 문자열의 포인터를 획득
    protocolPtr = vtnode_websocket_get_selected_protocol_from_event(eid)
    
    ' 2. NULL 포인터 체크 (안전 장치)
    If protocolPtr = 0 Then
        VtHttpaWsEventGetSelectedProtocol = ""
        Exit Function
    End If
    
    ' 3. UTF-8 포인터를 VBA String으로 변환
    VtHttpaWsEventGetSelectedProtocol = VtUtf8PtrToString(protocolPtr)
    
End Function

Public Function VtHttpaWsEventGetText(ByVal eid As LongLong) As String
    Dim TextPtr As LongPtr
    Dim TextLen As Long
    
    ' 1. 엔진으로부터 포인터와 길이를 직접 획득
    ' 성공 시 0이 아닌 값(True)을 반환한다고 가정
    If vtnode_websocket_get_utf8text_from_event(eid, TextPtr, TextLen) <> 0 Then
        
        ' 2. 유효성 검사: 포인터가 NULL이 아니고 길이가 0보다 큰지 확인
        If TextPtr <> 0 And TextLen > 0 Then
            ' 길이(textLen)를 명시적으로 전달하여 NULL 종료 문자에 의존하지 않음
            VtHttpaWsEventGetText = VtUtf8PtrToStringWithLen(TextPtr, TextLen)
        Else
            VtHttpaWsEventGetText = ""
        End If
        
    Else
        VtHttpaWsEventGetText = ""
    End If
End Function
Public Function VtHttpaWsEventGetBinary(ByVal eid As LongLong) As Byte()
    Dim DataPtr As LongPtr
    Dim DataLen As Long
    Dim Data() As Byte

    ' 1. 엔진으로부터 포인터와 데이터 길이 획득
    If vtnode_websocket_get_binary_from_event(eid, DataPtr, DataLen) <> 0 Then
        
        If DataPtr <> 0 And DataLen > 0 Then
            ' 2. 받은 길이만큼 VBA 배열 공간 확보
            ReDim Data(0 To DataLen - 1) As Byte
            
            ' 3. C 엔진 메모리 -> VBA 메모리로 고속 복사
            ' Win32 API인 RtlMoveMemory를 사용하여 포인터 직접 참조
            Call CopyMemory(VarPtr(Data(0)), DataPtr, DataLen)
            
            ' 4. 복사된 배열 반환
            VtHttpaWsEventGetBinary = Data
        Else
            ' 데이터가 0바이트인 경우 빈 배열 반환 (Nothing 대신)
            VtHttpaWsEventGetBinary = Split("", "")
        End If
        
    Else
        ' 실패 시 초기화되지 않은 상태 혹은 빈 배열 반환
        ' 사용 측에서 (Not Not arrayName)으로 체크 가능하도록 설계
    End If
End Function

