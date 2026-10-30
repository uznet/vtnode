Attribute VB_Name = "VtHttpaParams"
' =========================================================================
'  VTNode Edge HTTPa Engine API Declarations for VBA7
' =========================================================================
Option Explicit

#If VBA7 And Win64 Then

#If VTNODE_DEBUG = 1 Then
    ' 64비트 및 32비트 Office 2010 이상 호환 선언
    Private Declare PtrSafe Function vtnode_httpa_get_origin_port Lib "vtnode64d.dll" () As Long
    ' const char* 리턴 값은 메모리 주소(포인터)이므로 LongPtr로 안전하게 받습니다.
    Private Declare PtrSafe Function vtnode_httpa_get_origin_scheme Lib "vtnode64d.dll" () As LongPtr
    Private Declare PtrSafe Function vtnode_httpa_get_http_info Lib "vtnode64d.dll" (ByRef outPort As Long) As Long

    Private Declare PtrSafe Function vtnode_httpa_get_https_info Lib "vtnode64d.dll" ( _
        ByRef outPort As Long, _
        ByRef outEnabled As Long, _
        ByVal outCertFile As LongPtr, _
        ByVal certFileSize As Long, _
        ByVal outKeyFile As LongPtr, _
        ByVal keyFileSize As Long _
    ) As Long
    
    Private Declare PtrSafe Function vtnode_httpa_get_websocket_info Lib "vtnode64d.dll" ( _
        ByRef outEnabled As Long, _
        ByVal outSubprotocols As LongPtr, _
        ByVal subprotocolsSize As Long _
    ) As Long
    
#Else
    ' 64비트 및 32비트 Office 2010 이상 호환 선언
    Private Declare PtrSafe Function vtnode_httpa_get_origin_port Lib "vtnode64.dll" () As Long
    ' const char* 리턴 값은 메모리 주소(포인터)이므로 LongPtr로 안전하게 받습니다.
    Private Declare PtrSafe Function vtnode_httpa_get_origin_scheme Lib "vtnode64.dll" () As LongPtr
    
    Private Declare PtrSafe Function vtnode_httpa_get_http_info Lib "vtnode64.dll" (ByRef outPort As Long) As Long

    Private Declare PtrSafe Function vtnode_httpa_get_https_info Lib "vtnode64.dll" ( _
        ByRef outPort As Long, _
        ByRef outEnabled As Long, _
        ByVal outCertFile As LongPtr, _
        ByVal certFileSize As Long, _
        ByVal outKeyFile As LongPtr, _
        ByVal keyFileSize As Long _
    ) As Long
    Private Declare PtrSafe Function vtnode_httpa_get_websocket_info Lib "vtnode64.dll" ( _
        ByRef outEnabled As Long, _
        ByVal outSubprotocols As LongPtr, _
        ByVal subprotocolsSize As Long _
    ) As Long

#End If

    
#End If

' -------------------------------------------------------------------------
'  C 코어 엔진으로부터 현재 구동 중인 오리진 포트 번호를 획득
' -------------------------------------------------------------------------
Public Function VtHttpaGetOriginPort() As Long
    
    On Error GoTo ErrorHandler
    
    ' C 엔진의 __stdcall API를 직접 호출하여 포트 값을 리턴받습니다.
    VtHttpaGetOriginPort = vtnode_httpa_get_origin_port()
    Exit Function

ErrorHandler:
    ' DLL을 찾을 수 없거나 예외 발생 시 안전하게 0 또는 기본값 리턴
    Debug.Print "vtnode_httpa_get_origin_port 호출 실패: " & Err.Description
    VtHttpaGetOriginPort = 0

End Function

' -------------------------------------------------------------------------
'  C 코어의 포인터 문자열을 안전하게 유니코드 String으로 변환하여 반환
' -------------------------------------------------------------------------
Public Function VtHttpaGetOriginScheme() As String
    #If VBA7 Then
        Dim pStr As LongPtr
    #Else
        Dim pStr As Long
    #End If
    
    Dim Length As Long
    Dim buffer As String
    
    ' 1. C 엔진으로부터 메모리 주소를 획득 (vtnode 내부 static/global 메모리 영역)
    pStr = vtnode_httpa_get_origin_scheme()
    
    If pStr <> 0 Then
        VtHttpaGetOriginScheme = VtUtf8PtrToString(pStr)
    Else
       VtHttpaGetOriginScheme = ""
    End If
    
End Function



' -------------------------------------------------------------------------
'  테스트 서브루틴: 이그제큐션 창에서 포트와 스키마 확인
' -------------------------------------------------------------------------
' -------------------------------------------------------------------------
'  인자로 받은 host와 C 코어의 scheme, port를 조합하여 표준 Origin URL을 생성
'  (예: GetOrigin("localhost") -> "https://localhost:443")
' -------------------------------------------------------------------------
Public Function VtHttpaGetOrigin(ByVal host As String) As String
    Dim Port As Long
    Dim scheme As String
    Dim cleanHost As String
    
    ' 1. C 엔진 코어로부터 네트워크 설정 값 로드
    Port = VtHttpaGetOriginPort()
    scheme = VtHttpaGetOriginScheme()
    
'    Debug.Print "port=" & port
'    Debug.Print "scheme=" & scheme
    
    ' 예외 처리: 값의 유효성 검증
    If Len(scheme) = 0 Or Port <= 0 Then
        VtHttpaGetOrigin = ""
        Exit Function
    End If
    
    ' 2. 인자로 들어온 host 문자열 정제 (공백 제거 및 끝자리 슬래시'/' 제거)
    cleanHost = Trim$(host)
    If Right$(cleanHost, 1) = "/" Then
        cleanHost = left$(cleanHost, Len(cleanHost) - 1)
    End If
    
    ' 만약 호스트에 이미 스키마(http:// 등)가 포함되어 들어왔다면 해당 부분 제거
    If InStr(cleanHost, "://") > 0 Then
        cleanHost = Mid$(cleanHost, InStr(cleanHost, "://") + 3)
    End If
    
    ' 3. 데이터 터널링 및 웹 표준 규격에 맞는 최종 주소 조합
    '    (표준 80/443 포트일 경우 포트 표시 생략 옵션을 넣거나, 명시적으로 표기)
    If (scheme = "http" And Port = 80) Or (scheme = "https" And Port = 443) Then
        VtHttpaGetOrigin = scheme & "://" & cleanHost
    Else
        VtHttpaGetOrigin = scheme & "://" & cleanHost & ":" & CStr(Port)
    End If
End Function



Public Function VtHttpaGetHttpInfo(ByRef outPort As Long) As Boolean

    Dim ret As Long

    ret = vtnode_httpa_get_http_info(outPort)

    VtHttpaGetHttpInfo = (ret <> 0)

End Function

Public Function VtHttpaGetHttpsInfo( _
    ByRef outPort As Long, _
    ByRef outEnabled As Boolean, _
    ByRef outCertFile As String, _
    ByRef outKeyFile As String) As Boolean

    Const BUF_SIZE As Long = 1024

    Dim Enabled As Long
    Dim certBuf(0 To BUF_SIZE - 1) As Byte
    Dim keyBuf(0 To BUF_SIZE - 1) As Byte
    Dim ret As Long

    ret = vtnode_httpa_get_https_info( _
              outPort, _
              Enabled, _
              VarPtr(certBuf(0)), BUF_SIZE, _
              VarPtr(keyBuf(0)), BUF_SIZE)

    If ret = 0 Then
        Exit Function
    End If

    outEnabled = (Enabled <> 0)

    outCertFile = VtCharUtils.VtUtf8BytesToString(certBuf)
    outKeyFile = VtCharUtils.VtUtf8BytesToString(keyBuf)

    VtHttpaGetHttpsInfo = True

End Function


Public Function VtHttpaGetWebsocketInfo( _
    ByRef outEnabled As Boolean, _
    ByRef outSubprotocols As String) As Boolean

    Const BUF_SIZE As Long = 1024

    Dim Enabled As Long
    Dim subprotocolBuf(0 To BUF_SIZE - 1) As Byte
    Dim ret As Long

    ret = vtnode_httpa_get_websocket_info( _
              Enabled, _
              VarPtr(subprotocolBuf(0)), _
              BUF_SIZE)

    If ret = 0 Then
        Exit Function
    End If

    outEnabled = (Enabled <> 0)
    outSubprotocols = VtCharUtils.VtUtf8BytesToString(subprotocolBuf)


    VtHttpaGetWebsocketInfo = True

End Function

