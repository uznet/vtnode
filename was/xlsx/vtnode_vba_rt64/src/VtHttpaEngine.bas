Attribute VB_Name = "VtHttpaEngine"
' --- [ modVtnode.bas ] ---

Option Explicit

#If VBA7 And Win64 Then


#If VTNODE_DEBUG = 1 Then


    ' 1. 고성능 Asynchronous HTTP 엔진 가동 함수
    ' homeDir, iniFile, eout 모두 문자열 포인터 주소(LongPtr)를 넘겨받습니다.
    Private Declare PtrSafe Function vtnode_httpa_start Lib "vtnode64d.dll" ( _
        ByVal pHomeDir As LongPtr, _
        ByVal pIniFile As LongPtr, _
        ByVal pEout As LongPtr, _
        ByVal esize As Long _
    ) As Long

    ' 2. 엔진 정지 서브루틴 (반환값 void이므로 Sub로 선언)
    Private Declare PtrSafe Sub vtnode_httpa_stop Lib "vtnode64d.dll" ()
    
    ' 엔진이 현재 구동 중인지 여부를 확인하는 함수 (BOOL 반환형 -> Long 매핑)
    Private Declare PtrSafe Function vtnode_httpa_is_run Lib "vtnode64d.dll" () As Long
#Else  ' released mode

    ' 1. 고성능 Asynchronous HTTP 엔진 가동 함수
    ' homeDir, iniFile, eout 모두 문자열 포인터 주소(LongPtr)를 넘겨받습니다.
    Private Declare PtrSafe Function vtnode_httpa_start Lib "vtnode64.dll" ( _
        ByVal pHomeDir As LongPtr, _
        ByVal pIniFile As LongPtr, _
        ByVal pEout As LongPtr, _
        ByVal esize As Long _
    ) As Long

    ' 2. 엔진 정지 서브루틴 (반환값 void이므로 Sub로 선언)
    Private Declare PtrSafe Sub vtnode_httpa_stop Lib "vtnode64.dll" ()
    
    ' 엔진이 현재 구동 중인지 여부를 확인하는 함수 (BOOL 반환형 -> Long 매핑)
    Private Declare PtrSafe Function vtnode_httpa_is_run Lib "vtnode64.dll" () As Long

#End If

#End If



Public Function VtHttpaStart(ByVal homeDir As String, ByVal iniFile As String) As VTHttpaEngineStartResult
    Dim utf8Home() As Byte
    Dim utf8Ini() As Byte
    Dim eout(0 To 1023) As Byte
    Dim pHome As LongPtr
    Dim pIni As LongPtr
    Dim result As Long
    Dim startResult As VTHttpaEngineStartResult

    ' 1. 로컬 구조체 초기값 설정
    startResult.success = False
    startResult.emsg = ""
    
    ' 2. homeDir UTF-8 주소 변환 (필수)
    If homeDir <> "" Then
        utf8Home = VtStringToUtf8Bytes(homeDir)
        pHome = VarPtr(utf8Home(0))
    Else
        startResult.emsg = "VBA Error: homeDir 경로가 누락되었습니다."
        ' 탈출하기 전에 현재까지의 실패 결과 구조체를 최종 바인딩
        VtHttpaStart = startResult
        Exit Function
    End If
    
    ' 3. iniFile UTF-8 주소 변환 (빈 값일 경우 C 엔진에 NULL 포인터 전달)
    If iniFile <> "" Then
        utf8Ini = VtStringToUtf8Bytes(iniFile)
        pIni = VarPtr(utf8Ini(0))
    Else
        pIni = 0 ' C의 NULL
    End If
    
    ' 4. 네이티브 C 엔진 가동 호출
    ' BOOL(__stdcall) 반환값은 0이 False, 0이 아니면 True입니다.
    result = vtnode_httpa_start(pHome, pIni, VarPtr(eout(0)), 1024)
    
    ' 5. 결과 매핑
    If result <> 0 Then
        startResult.success = True
        startResult.emsg = "Success"
    Else
        ' vtnode package 자체 검증 등 내부에서 에러 발생 시 버퍼 파싱
        startResult.success = False ' 명시적 지정
        startResult.emsg = VtUtf8PtrToString(VarPtr(eout(0)))
        'Debug.Print "vtnode_httpa_start Failed: " & startResult.emsg
    End If
    
    ' 6. [중요] 모든 처리가 완료된 최종 구조체를 함수 반환값에 대입
    VtHttpaStart = startResult
End Function



Public Sub VtHttpaStop()
    On Error GoTo ErrorHandler

    ' 1. 백그라운드에서 비동기로 돌고 있는 C 엔진 정지 API 호출
    ' (상단 Declare문에서 void 매핑인 Sub로 선언했으므로 Call로 호출)
    Call vtnode_httpa_stop
    
    ' 2. 정상 정지 로그 기록
    Debug.Print "vtnode_httpa_stop: 비동기 HTTP 엔진이 성공적으로 정지되었습니다."
    
    Exit Sub

ErrorHandler:
    ' 예기치 못한 크래시나 누수 방지를 위한 예외 처리
    Debug.Print "VtHttpaStop 오류 발생 (Err #" & Err.Number & "): " & Err.Description
End Sub


Public Function VtHttpaIsRun() As Boolean
    ' C 엔진이 0이 아닌 값을 리턴하면 현재 가동 중(True)인 상태입니다.
    If vtnode_httpa_is_run() <> 0 Then
        VtHttpaIsRun = True
    Else
        VtHttpaIsRun = False
    End If
End Function
'#If VBA7 Then
'
'    Private Declare PtrSafe Function vtnode_httpa_init Lib "vtnode64d.dll" ( _
'        ByVal homeDir As LongPtr, _
'        ByVal iniFile As LongPtr, _
'        ByVal eout As LongPtr, _
'        ByVal esize As Long) As Long
'
'    Private Declare PtrSafe Sub vtnode_httpa_uninit Lib "vtnode64d.dll" ()
'
'    Private Declare PtrSafe Function vtnode_httpa_is_initialized Lib "vtnode64d.dll" () As Long
'
'    Private Declare PtrSafe Function vtnode_httpa_start Lib "vtnode64d.dll" ( _
'        ByVal eout As LongPtr, _
'        ByVal esize As Long) As Long
'
'    Private Declare PtrSafe Function vtnode_httpa_is_run Lib "vtnode64d.dll" () As Long
'
'    Private Declare PtrSafe Sub vtnode_httpa_stop Lib "vtnode64d.dll" ()
'
'#Else
'
'    ' 32bit 필요 시 xhttpa32d.dll 선언 추가
'
'#End If
'
'
'' =========================================================
'' Low / Mid Level Engine Wrapper
'' =========================================================
'
'
'' ByRef로 전달된 startResult는 호출 측의 메모리 주소를 직접 가리킵니다.
'Private Function VtHttpaEngineInit(ByVal homeDir As String, _
'                                  ByVal iniFile As String, _
'                                  ByRef startResult As VTHttpaEngineStartReult) As Boolean
'
'    Dim utf8Home() As Byte
'    Dim utf8Ini() As Byte
'    Dim eout(0 To 1023) As Byte
'    Dim pIni As LongPtr
'    Dim result As Long
'
'    ' 1. 결과 구조체 초기화
'    startResult.success = False
'    startResult.emsg = ""
'
'    ' 이미 초기화된 경우 처리
'    If VtHttpaEngineIsInitialized() Then
'        startResult.success = True
'        startResult.emsg = "Already initialized"
'        VtHttpaEngineInit = True
'        Exit Function
'    End If
'
'    ' 2. 문자열 UTF-8 변환
'    utf8Home = VtStringToUtf8Bytes(homeDir)
'
'    If iniFile <> "" Then
'        utf8Ini = VtStringToUtf8Bytes(iniFile)
'        pIni = VarPtr(utf8Ini(0))
'    Else
'        pIni = 0
'    End If
'
'    ' 3. C 엔진 호출 (eout 버퍼의 포인터를 넘겨 에러 메시지 수신)
'    result = vtnode_httpa_init(VarPtr(utf8Home(0)), pIni, VarPtr(eout(0)), 1024)
'
'    ' 4. 결과 매핑
'    If result <> 0 Then
'        startResult.success = True
'        startResult.emsg = "Success"
'        VtHttpaEngineInit = True
'    Else
'        ' 엔진에서 넘어온 UTF-8 에러 메시지를 VBA String으로 변환하여 구조체에 저장
'        startResult.success = False
'        startResult.emsg = VtUtf8PtrToString(VarPtr(eout(0)))
'
'        Debug.Print "VtHttpaEngineInit Failed: " & startResult.emsg
'        VtHttpaEngineInit = False
'    End If
'
'End Function
'
'
'Private Sub VtHttpaEngineUninit()
'    If Not VtHttpaEngineIsInitialized() Then
'        Exit Sub
'    End If
'
'    Call vtnode_httpa_uninit
'    Debug.Print "VtHttpa Engine Uninitialized."
'End Sub
'
'
'Private Function VtHttpaEngineIsInitialized() As Boolean
'    VtHttpaEngineIsInitialized = (vtnode_httpa_is_initialized() <> 0)
'End Function
'
'
'Private Function VtHttpaEngineStart(ByRef startResult As VTHttpaEngineStartReult) As Boolean
'    Dim eout(0 To 1023) As Byte
'    Dim result As Long
'
'    ' 1. 초기 상태 설정
'    startResult.success = False
'    startResult.emsg = ""
'
'    ' 이미 실행 중인 경우
'    If VtHttpaEngineIsRun() Then
'        startResult.success = True
'        startResult.emsg = "Engine is already running."
'        VtHttpaEngineStart = True
'        Exit Function
'    End If
'
'    ' 초기화 여부 확인
'    If Not VtHttpaEngineIsInitialized() Then
'        startResult.success = False
'        startResult.emsg = "Engine is not initialized."
'        Debug.Print "VtHttpaEngineStart Failed: " & startResult.emsg
'        VtHttpaEngineStart = False
'        Exit Function
'    End If
'
'    ' 2. C 엔진 시작 호출
'    result = vtnode_httpa_start(VarPtr(eout(0)), 1024)
'
'    ' 3. 결과 판별 및 구조체 업데이트
'    If result <> 0 Then
'        startResult.success = True
'        startResult.emsg = "Success"
'        VtHttpaEngineStart = True
'        Debug.Print "VtHttpaEngineStart: Engine started."
'    Else
'        ' C 엔진에서 넘어온 에러 메시지 캡처
'        startResult.success = False
'        startResult.emsg = VtUtf8PtrToString(VarPtr(eout(0)))
'
'        Debug.Print "VtHttpaEngineStart Failed: " & startResult.emsg
'        VtHttpaEngineStart = False
'    End If
'
'End Function
'
'
'Public Function VtHttpaEngineIsRun() As Boolean
'    VtHttpaEngineIsRun = (vtnode_httpa_is_run() <> 0)
'End Function
'
'
'Public Sub VtHttpaEngineStop()
'    If Not VtHttpaEngineIsRun() Then
'        Exit Sub
'    End If
'
'    Call vtnode_httpa_stop
'    Debug.Print "VtHttpa Engine Stopped."
'End Sub
'
''=====================Public Function ==================
'
'' =========================================================
'' High Level Engine Lifecycle API
'' Init -> Start
'' =========================================================
'
'' =========================================================
'' High Level Engine Lifecycle API
'' Init -> Start 과정을 하나로 묶어 구조체로 결과 반환
'' =========================================================
'Public Function VtHttpaEngineRun(ByVal homeDir As String, ByVal iniFile As String) As VTHttpaEngineStartReult
'
'    Dim ret As VTHttpaEngineStartReult
'
'    ' 1. 초기값 설정
'    ret.success = False
'    ret.emsg = ""
'
'    ' 2. 이미 실행 중인지 확인
'    If VtHttpaEngineIsRun() Then
'        ret.success = True
'        ret.emsg = "Engine is already running."
'        VtHttpaEngineRun = ret
'        Exit Function
'    End If
'
'    ' 3. 초기화 여부 확인 및 실행 (Init)
'    If Not VtHttpaEngineIsInitialized() Then
'        ' 내부 함수에 ret를 넘겨 상세 에러를 받아옴
'        If Not VtHttpaEngineInit(homeDir, iniFile, ret) Then
'            ' Init 실패 시 ret에 담긴 에러와 함께 반환
'            VtHttpaEngineRun = ret
'            Exit Function
'        End If
'    End If
'
'    ' 4. 엔진 시작 (Start)
'    ' 이미 초기화가 성공했다면 다시 ret를 넘겨 Start 결과를 받아옴
'    If Not VtHttpaEngineStart(ret) Then
'        ' Start 실패 시 결과 반환
'        VtHttpaEngineRun = ret
'        Exit Function
'    End If
'
'    ' 5. 최종 성공 상태 반환
'    ret.success = True
'    ret.emsg = "Engine started successfully."
'    VtHttpaEngineRun = ret
'
'End Function
'
'
'' =========================================================
'' High Level Engine Lifecycle API
'' Stop -> Uninit
'' =========================================================
'
'Public Sub VtHttpaEngineShutdown()
'
'    If VtHttpaEngineIsRun() Then
'        Call VtHttpaEngineStop
'    End If
'
'    If VtHttpaEngineIsInitialized() Then
'        Call VtHttpaEngineUninit
'    End If
'
'End Sub
'
