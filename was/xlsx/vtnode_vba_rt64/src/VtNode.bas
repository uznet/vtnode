Attribute VB_Name = "VtNode"
Option Explicit

#If VBA7 And Win64 Then


    ' 인자 타입을 ByVal As LongPtr로 변경하여 VarPtr 주소값이 그대로 전달되도록 수정
#If VTNODE_DEBUG = 1 Then
    Private Declare PtrSafe Function vtnode_init Lib "vtnode64d.dll" ( _
        ByVal pEout As LongPtr, _
        ByVal esize As Long _
    ) As Long
    
    Private Declare PtrSafe Sub vtnode_uninit Lib "vtnode64d.dll" ()

    Private Declare PtrSafe Function vtnode_verify Lib "vtnode64d.dll" ( _
        ByVal pVtnodeHomeDir As LongPtr, _
        ByVal pEout As LongPtr, _
        ByVal esize As Long _
    ) As Long

#Else

    Private Declare PtrSafe Function vtnode_init Lib "vtnode64.dll" ( _
        ByVal pEout As LongPtr, _
        ByVal esize As Long _
    ) As Long
    Private Declare PtrSafe Sub vtnode_uninit Lib "vtnode64.dll" ()

    Private Declare PtrSafe Function vtnode_verify Lib "vtnode64.dll" ( _
        ByVal pVtnodeHomeDir As LongPtr, _
        ByVal pEout As LongPtr, _
        ByVal esize As Long _
    ) As Long


#End If

#End If

Public Type VTNodeInitResult
    success As Boolean
    emsg    As String
End Type

Public Type VTNodeVerifyResult
    success As Boolean
    emsg    As String
End Type


' 엔진을 가동하기 전에 먼저 경로를 잡아주는 함수
Public Sub VtNodeLoadEngineDLL()
    Dim excelPath As String
    
    ' 현재 활성화된 엑셀 파일의 디렉토리 경로 가져오기
    excelPath = ThisWorkbook.Path
    
    ' 작업 디렉토리를 엑셀 파일 위치로 변경 (DLL 검색 경로 최우선 순위로 지정)
    If VtWin32.SetCurrentDirectory(excelPath) = 0 Then
        ' API 실패 시 VBA 기본 명령어로 대체 방어
        ChDrive excelPath
        ChDir excelPath
    End If
    
    Debug.Print "DLL Search Path Set To: " & excelPath
End Sub

'Public Function VtNodeInit() As VTNodeInitResult
'    Dim initResult As VTNodeInitResult
'
'    'Call LoadEngineDLL
'
'    ' 1. 결과 구조체 초기화
'    initResult.success = True
'    initResult.emsg = ""
'
'    VtNodeInit = initResult
'
'
'End Function
'
'Public Sub VtNodeUnInit()
'End Sub


Public Function VtNodeInit() As VTNodeInitResult
    Dim result As Long
    Dim eout(0 To 1023) As Byte
    Dim initResult As VTNodeInitResult


    'Call LoadEngineDLL

    ' 1. 결과 구조체 초기화
    initResult.success = False
    initResult.emsg = ""

    ' 2. C 엔진 호출 (주소값을 명확히 전달)
    result = vtnode_init(VarPtr(eout(0)), 1024)

    ' 3. 결과 매핑 (인자명 startResult -> initResult 오타 수정)
    If result <> 0 Then
        initResult.success = True
        initResult.emsg = "Success"
    Else
        initResult.success = False
        initResult.emsg = VtUtf8PtrToString(VarPtr(eout(0)))

    End If


    VtNodeInit = initResult

End Function

Public Sub VtNodeUnInit()
    Call vtnode_uninit
   ' Debug.Print "VtNode Engine Uninitialized."
End Sub

Public Function VtNodeVerify(ByVal vtnodeHomeDir As String) As VTNodeVerifyResult
    Dim utf8Dir() As Byte
    Dim eout(0 To 1023) As Byte
    Dim pDir As LongPtr
    Dim pFile As LongPtr
    Dim result As Long
    Dim verifyResult As VTNodeVerifyResult
    

    ' 1. 결과 구조체 초기화
    verifyResult.success = False
    verifyResult.emsg = ""
    
    ' 2. binDir 유효성 체크 및 UTF-8 변환
    If vtnodeHomeDir <> "" Then
        utf8Dir = VtStringToUtf8Bytes(vtnodeHomeDir)
        pDir = VarPtr(utf8Dir(0))
    Else
        verifyResult.emsg = "vtnodeHomeDir is null or empty"
        VtNodeVerify = verifyResult
        Exit Function
    End If
    
    
    ' 4. C 엔진 호출
    result = vtnode_verify(pDir, VarPtr(eout(0)), 1024)
    
    ' 5. 결과 매핑
    If result <> 0 Then
        verifyResult.success = True
        verifyResult.emsg = "Success"
    Else
        verifyResult.success = False
        verifyResult.emsg = VtUtf8PtrToString(VarPtr(eout(0)))
        
        Debug.Print "VtNodeVerify Failed: " & verifyResult.emsg
    End If
    
    VtNodeVerify = verifyResult
End Function

