Attribute VB_Name = "VtNode"
Option Explicit




#If VBA7 And Win64 Then


    ' 인자 타입을 ByVal As LongPtr로 변경하여 VarPtr 주소값이 그대로 전달되도록 수정
#If VTNODE_DEBUG = 1 Then
    Private Declare PtrSafe Function vtnode_init Lib "vtnode64d.dll" ( _
        ByVal pAppHomeDir As LongPtr, _
        ByVal pEout As LongPtr, _
        ByVal esize As Long _
    ) As Long
    
    Private Declare PtrSafe Sub vtnode_uninit Lib "vtnode64d.dll" ()

    Private Declare PtrSafe Function vtnode_verify Lib "vtnode64d.dll" ( _
        ByVal pVtnodeHomeDir As LongPtr, _
        ByVal pEout As LongPtr, _
        ByVal esize As Long _
    ) As Long
    
    Private Declare PtrSafe Function vtnode_get_abi_version Lib "vtnode64d.dll" () As Long
    Private Declare PtrSafe Function vtnode_is_initialized Lib "vtnode64d.dll" () As Long

#Else

    Private Declare PtrSafe Function vtnode_init Lib "vtnode64.dll" ( _
        ByVal pAppHomeDir As LongPtr, _
        ByVal pEout As LongPtr, _
        ByVal esize As Long _
    ) As Long
    Private Declare PtrSafe Sub vtnode_uninit Lib "vtnode64.dll" ()

    Private Declare PtrSafe Function vtnode_verify Lib "vtnode64.dll" ( _
        ByVal pVtnodeHomeDir As LongPtr, _
        ByVal pEout As LongPtr, _
        ByVal esize As Long _
    ) As Long
    
    
    Private Declare PtrSafe Function vtnode_get_abi_version Lib "vtnode64.dll" () As Long

    Private Declare PtrSafe Function vtnode_is_initialized Lib "vtnode64.dll" () As Long
    
    
   
    
    
    
#End If

#Else

    Private Function vtnode_get_abi_version() As Long
    
      vtnode_get_abi_version = 0
       
    End Function
    
    Private Function vtnode_is_initialized() As Boolean
       
       vtnode_is_initialized = False
       
      Err.Raise _
        vbObjectError + 1001, _
        "VTNode Runtime", _
        "VTNode requires 64-bit Microsoft Excel." & vbCrLf & _
        "Your current Excel is 32-bit."
    
         
    End Function
    Private Function vtnode_init(ByVal pAppHomeDir As LongPtr, ByVal pEout As LongPtr, ByVal esize As Long) As Long
    
       vtnode_init = 0
    
      
    End Function

#End If

Public Const VTNODE_VBA_VERSION_MAJOR As Long = 0
Public Const VTNODE_VBA_VERSION_MINOR As Long = 1
Public Const VTNODE_VBA_VERSION_PATCH As Long = 0

Public Const VTNODE_REQUIRED_ABI_VERSION As Long = 1


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
    excelPath = ThisWorkbook.path
    
    ' 작업 디렉토리를 엑셀 파일 위치로 변경 (DLL 검색 경로 최우선 순위로 지정)
    If VtWin32.SetCurrentDirectory(excelPath) = 0 Then
        ' API 실패 시 VBA 기본 명령어로 대체 방어
        ChDrive excelPath
        ChDir excelPath
    End If
    
    Debug.Print "DLL Search Path Set To: " & excelPath
End Sub



Private Sub CheckABIVersionCompatibility()


    Dim nativeAbi As Long
     
    nativeAbi = vtnode_get_abi_version()

    If nativeAbi <> VTNODE_REQUIRED_ABI_VERSION Then
    
       
        Err.Raise _
            vbObjectError + 1002, _
            "VTNode Runtime", _
            "VTNode Native Runtime ABI mismatch." & vbCrLf & _
            "Required: " & VTNODE_REQUIRED_ABI_VERSION & vbCrLf & _
            "Installed: " & nativeAbi

    End If

End Sub




Public Function VtNodeInit(ByVal appHomeDir As String) As VTNodeInitResult

    Dim utf8Home() As Byte
    Dim eout(0 To 1024) As Byte
    Dim result As Long
    Dim initResult As VTNodeInitResult

    initResult.success = False
    initResult.emsg = vbNullString
    
    
    
 
    
    ' 2. Check Native Runtime ABI compatibility.
    CheckABIVersionCompatibility
 

    ' 3. Validate application home directory.
    If Len(appHomeDir) = 0 Then
        initResult.emsg = "VBA Error: appHomeDir is required."
        VtNodeInit = initResult
        Exit Function
    End If


    ' 4. Convert path to NUL-terminated UTF-8.
    utf8Home = VtStringToUtf8Bytes(appHomeDir)

    ' 5. Initialize VTNode Native Runtime.
    result = vtnode_init( _
                VarPtr(utf8Home(0)), _
                VarPtr(eout(0)), _
                1024)

    ' 6. Process initialization result.
    If result <> 0 Then
        initResult.success = True
        initResult.emsg = "Success"
    Else
        initResult.emsg = VtUtf8PtrToString(VarPtr(eout(0)))
    End If

    VtNodeInit = initResult

End Function

Public Sub VtNodeUnInit()
    Call vtnode_uninit
End Sub

Public Function VtNodeIsInitialized() As Boolean


  CheckABIVersionCompatibility
     
  If vtnode_is_initialized() <> 0 Then
      VtNodeIsInitialized = True
  
  Else
      VtNodeIsInitialized = False
   End If
End Function
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



