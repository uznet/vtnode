Attribute VB_Name = "VtHttpaNet"

Option Explicit

' VTNode.io Httpa Client API 선언
' serverAddr: 서버 주소 (e.g., "127.0.0.1")
' serverPort: 포트 번호 (e.g., 80, 443)
' secure: SSL/TLS 사용 여부 (VBA True는 -1이므로 엔진 전달 시 주의 필요)
' timeoutSec: 타임아웃 (초 단위)

#If VBA7 And Win64 Then

#If VTNODE_DEBUG = 1 Then
    ' 엔진 연결: 새로운 인스턴스 ID(LongLong)를 반환
    Private Declare PtrSafe Function vtnode_httpa_connect Lib "vtnode64d.dll" ( _
        ByVal serverAddr As LongPtr, _
        ByVal serverPort As Long, _
        ByVal secure As Long, _
        ByVal timeoutSec As Long) As LongLong

    ' 엔진 연결 해제: 인스턴스 ID를 기반으로 통로를 닫음
    Private Declare PtrSafe Function vtnode_httpa_disconnect Lib "vtnode64d.dll" ( _
        ByVal UziId As LongLong) As Long
                
                
    ' 디버그 빌드 DLL (vtnode64d.dll) 바인딩
    Private Declare PtrSafe Function vtnode_httpa_get_origin Lib "vtnode64d.dll" ( _
            ByVal host As LongPtr, _
            ByVal outBuffer As LongPtr, _
            ByVal osize As Long) As Long
                
#Else

    ' 엔진 연결: 새로운 인스턴스 ID(LongLong)를 반환
    Private Declare PtrSafe Function vtnode_httpa_connect Lib "vtnode64.dll" ( _
        ByVal serverAddr As LongPtr, _
        ByVal serverPort As Long, _
        ByVal secure As Long, _
        ByVal timeoutSec As Long) As LongLong

    ' 엔진 연결 해제: 인스턴스 ID를 기반으로 통로를 닫음
    Private Declare PtrSafe Function vtnode_httpa_disconnect Lib "vtnode64.dll" ( _
        ByVal UziId As LongLong) As Long
        
      Private Declare PtrSafe Function vtnode_httpa_get_origin Lib "vtnode64.dll" ( _
            ByVal host As LongPtr, _
            ByVal outBuffer As LongPtr, _
            ByVal osize As Long) As Long
        

#End If
#End If



Public Function VtHttpaConnect(ByVal serverAddr As String, _
                               ByVal serverPort As Long, _
                               ByVal secure As Boolean, _
                               ByVal timeoutSec As Long) As LongLong
    Dim Res As LongLong
    Dim utf8Addr() As Byte
    Dim cSecure As Long

    ' 1. UTF-8 변환 (엔진이 요구하는 char* 규격 준수)
    utf8Addr = VtStringToUtf8Bytes(serverAddr)

    ' 2. Boolean 처리: C의 BOOL(int)은 보통 1(True), 0(False)입니다.
    ' VBA의 True는 -1이므로, Abs()를 취해 1로 넘겨주는 것이 안전합니다.
    cSecure = IIf(secure, 1, 0)

    ' 3. API 호출
    ' VarPtr(utf8Addr(0))을 통해 바이트 배열의 첫 번째 메모리 주소를 전달합니다.
    Res = vtnode_httpa_connect(VarPtr(utf8Addr(0)), serverPort, cSecure, timeoutSec)

    ' 4. 결과 반환
    VtHttpaConnect = Res
End Function


' uziid를 ByRef로 전달받아 해제 후 0으로 초기화합니다.
Public Sub VtHttpaDisconnect(ByRef CID As LongLong)
    If CID <> 0 Then
        Debug.Print "[VTNode] Disconnecting Instance: " & CID
        
        ' 1. 로우레벨 API 호출
        vtnode_httpa_disconnect CID
        
        ' 2. [핵심] 핸들 무효화
        ' 호출한 곳의 변수 값을 0으로 만들어 다음번 호출 시 If문에 걸리게 합니다.
        CID = 0
    Else
        Debug.Print "[VTNode] Disconnect Skip: Instance ID is already 0"
    End If
End Sub


''' <summary>
''' VTNode HTTP Asynchronous 엔진의 오리진 경로 주소를 네이티브 포인터 제어로 획득합니다.
''' </summary>
Public Function VtHttpaGetOrigin(ByVal hostStr As String, ByRef outOrigin As String) As Boolean
    Dim ret As Long
    Const OUT_BUFFER_SIZE As Long = 256 ' 오리진 문자열을 담을 충분한 메모리 버퍼 할당
    
    ' 1. 입력 문자열(host)을 C 가 인지할 수 있는 ANSI(멀티바이트) 바이트 배열로 변환
    Dim hostBytes() As Byte
    hostBytes = StrConv(hostStr & vbNullChar, vbFromUnicode) ' C-String 종료 문자(\0) 결합 후 캐스팅
    
    ' 2. 출력 문자열(out)을 담을 고정 ANSI 바이트 힙 스택 생성 (malloc 효과)
    Dim outBytes() As Byte
    ReDim outBytes(OUT_BUFFER_SIZE - 1)
    
    ' 3. VarPtr()를 사용하여 바이트 배열의 첫 번째 원소 물리 주소를 포인터로 다이렉트 덤프
    ret = vtnode_httpa_get_origin(VarPtr(hostBytes(0)), VarPtr(outBytes(0)), OUT_BUFFER_SIZE)
    
    ' 4. C의 BOOL 반환값(통상 TRUE = 1, FALSE = 0) 검증 및 마샬링 파싱
    If ret <> 0 Then
        ' ANSI 바이트 스트림을 다시 VBA용 유니코드 스트링으로 리버스 캐스팅
        Dim rawStr As String: rawStr = StrConv(outBytes, vbUnicode)
        
        ' NULL 종료 문자 앞까지만 정확하게 잘라내어 유효 데이터 획득
        outOrigin = left$(rawStr, InStr(rawStr, vbNullChar) - 1)
        VtHttpaGetOrigin = True
    Else
        outOrigin = ""
        VtHttpaGetOrigin = False
    End If
End Function
