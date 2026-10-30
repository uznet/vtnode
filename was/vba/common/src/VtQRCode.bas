Attribute VB_Name = "VtQRCode"

' =========================================================================
'  VTNode Edge QR-Code Generator API Declarations for VBA7
' =========================================================================

Option Explicit

#If VBA7 And Win64 Then
  
#If VTNODE_DEBUG = 1 Then

    ' C의 BOOL은 4바이트 정수(int)이므로 VBA에서는 Long과 매칭됩니다.
    ' const char* 입력 인자는 ByVal As String으로 선언하면 자동으로 ANSI 변환됩니다.
    Private Declare PtrSafe Function vtnode_qrcode_save_bmp Lib "vtnode64d.dll" ( _
        ByVal text As LongPtr, _
        ByVal moduleSize As Long, _
        ByVal fileName As LongPtr _
    ) As Long
#Else
    Private Declare PtrSafe Function vtnode_qrcode_save_bmp Lib "vtnode64.dll" ( _
        ByVal text As LongPtr, _
        ByVal moduleSize As Long, _
        ByVal fileName As LongPtr _
    ) As Long

#End If

    
#End If


' -------------------------------------------------------------------------
'  VTNode 전용 QR 코드 생성 함수 (LongPtr 텍스트 포인터 제어 버전)
' -------------------------------------------------------------------------
' moduleSize
'   Size in pixels of one QR module.
'   Range : 1 ~ 32
'   Default : 4
'    <=0 : default value 4

Public Function VtGenerateQRCode(ByVal text As String, ByVal moduleSize As Long, ByVal savePath As String) As Boolean
    Dim cleanText As String
    Dim ansiBytes() As Byte
    Dim cResult As Long
    
    On Error GoTo ErrorHandler
    
    ' 1. 가변형 text 인자를 문자열로 안전하게 변환
    cleanText = Trim$(CStr(text))
    
    ' 방어적 코드: 필수값 검증
    If Len(cleanText) = 0 Or Len(savePath) = 0 Then
        'Debug.Print "VtGenerateQRCode 실패: 잘못된 파라미터"
        VtGenerateQRCode = False
        Exit Function
    End If
    
    
    Dim utf8Text() As Byte
    Dim Utf8Path() As Byte
    
    
    ' 1. 텍스트 UTF-8 변환 및 길이 체크
    utf8Text = VtStringToUtf8Bytes(cleanText)         ' null terminated
    Utf8Path = VtStringToUtf8Bytes(savePath)     ' null terminated
    
    
    ' 3. C API 호출: 바이트 배열의 첫 번째 요소의 메모리 주소(VarPtr)를 LongPtr로 강제 전달
    '    ByVal text As LongPtr 구조이므로 주소 값을 값(Value)으로 넘겨야 C에서 포인터로 받습니다.
    cResult = vtnode_qrcode_save_bmp(VarPtr(utf8Text(0)), moduleSize, VarPtr(Utf8Path(0)))
    
    ' 4. C의 BOOL(1) 반환 값을 VBA Boolean으로 파싱
    If cResult = 1 Then
        VtGenerateQRCode = True
        Debug.Print "VTNode QR 생성 성공 [LongPtr 통신] -> " & savePath
    Else
        VtGenerateQRCode = False
        Debug.Print "VTNode QR 생성 실패 (C 코어 리턴 에러)"
    End If
    Exit Function

ErrorHandler:
    Debug.Print "VtGenerateQRCode 런타임 에러: " & Err.Description
    VtGenerateQRCode = False
End Function
