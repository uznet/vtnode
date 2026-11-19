Attribute VB_Name = "MyHttpa"
'Attribute VB_Name = "MyHttpaMain"

Option Explicit

Private Const MODULE_NAME As String = "MyHttpa"

'Private mWorkbookHomeDirectory As String
'Private mVTNodeHomeDirectory As String



'' =========================================================
'' Excel Application 이벤트 모니터 해제
'' =========================================================
'Public Sub EndExcelCloseMonitor()
'
'    On Error Resume Next
'
'    If Not mExcelEvents Is Nothing Then
'        Set mExcelEvents.App = Nothing
'        Set mExcelEvents = Nothing
'    End If
'
'End Sub
'' =========================================================
'' VTNode Native Runtime 초기화
'' =========================================================
'Public Function HttpaInit() As Boolean
'
'    Const PROC_NAME As String = "HttpaInit"
'
'    Dim LogDirectory As String
'    Dim LogFilePrefix As String
'    Dim initResult As VTNodeInitResult
'    Dim verifyResult As VTNodeVerifyResult
'
'    On Error GoTo EH
'
'    HttpaInit = VtNode.VtNodeIsInitialized()
'    ' 이미 초기화되어 있으면 재초기화하지 않는다.
'    If HttpaInit Then
'        Exit Function
'    End If
'
'
'    mWorkbookHomeDirectory = ThisWorkbook.path
'    mVTNodeHomeDirectory = VtEnv.VtGetEnv("VTNODE_HOME", vbNullString)
'
'    ' Application Home 확인
'    If Len(Trim$(mWorkbookHomeDirectory)) = 0 Then
'        MsgBox "Application Home directory is empty.", _
'               vbCritical, "VTNode Initialization"
'        Exit Function
'    End If
'
'    ' VTNode Home 확인
'    If Len(Trim$(mVTNodeHomeDirectory)) = 0 Then
'        MsgBox "VTNODE_HOME 환경변수가 설정되지 않았습니다.", _
'               vbCritical, "VTNode Initialization"
'        Exit Function
'    End If
'
'    LogDirectory = mWorkbookHomeDirectory & Application.PathSeparator & "logs"
'
'    LogFilePrefix = ThisWorkbook.Name
'    If InStrRev(LogFilePrefix, ".") > 1 Then
'        LogFilePrefix = left$(LogFilePrefix, InStrRev(LogFilePrefix, ".") - 1)
'    End If
'
'    VtLog.VtLogInitialize LogDirectory, LogFilePrefix
'
'    VtLogInfo MODULE_NAME, PROC_NAME, _
'              "HttpaInit() started. VTNODE_HOME=" & _
'              mVTNodeHomeDirectory
'
'    ' Native Runtime 초기화
'    initResult = VtNode.VtNodeInit(mWorkbookHomeDirectory)
'
'    If Not initResult.success Then
'
'        VtLogError MODULE_NAME, PROC_NAME, _
'                   "VtNodeInit() failed. reason=" & _
'                   initResult.emsg
'
'        MsgBox "VtNodeInit() failed." & vbCrLf & _
'               initResult.emsg, vbCritical, "VTNode"
'
'        GoTo InitFailed
'    End If
'
'
'    VtLogInfo MODULE_NAME, PROC_NAME, _
'              "VtNodeInit() success"
'
'    ' Runtime 설치 및 환경 검증
'    verifyResult = VtNode.VtNodeVerify(mVTNodeHomeDirectory)
'
'    If Not verifyResult.success Then
'
'        VtLogError MODULE_NAME, PROC_NAME, _
'                   "VtNodeVerify() failed. reason=" & _
'                   verifyResult.emsg
'
'        MsgBox "VtNodeVerify() failed." & vbCrLf & _
'               verifyResult.emsg, vbCritical, "VTNode"
'
'        GoTo InitFailed
'    End If
'
'    VtLogInfo MODULE_NAME, PROC_NAME, _
'              "VtNodeVerify() success"
'
'    HttpaInit = VtNode.VtNodeIsInitialized()
'
'
'    VtLogInfo MODULE_NAME, PROC_NAME, _
'              "HttpaInit() completed successfully"
'
'    Exit Function
'
'InitFailed:
'    On Error Resume Next
'
'    If VtNode.VtNodeIsInitialized() Then
'        VtNode.VtNodeUnInit
'    End If
'
'    VtLogClose
'    Exit Function
'
'EH:
'    Dim errNumber As Long
'    Dim errDescription As String
'
'    errNumber = Err.number
'    errDescription = Err.description
'
'    MsgBox "HttpaInit() unexpected error." & vbCrLf & _
'           "Number: " & CStr(errNumber) & vbCrLf & _
'           "Description: " & errDescription, _
'           vbCritical, "VTNode"
'
'    Resume InitFailed
'
'End Function
'
'
'Public Sub WasUninit()
'
'    Debug.Print "WasUninit:Begin"
'    If MyHttpaLooper.IsLooping() Then
'
'        MsgBox "Stop HTTP Event Loop first.", _
'               vbExclamation, "VTNode"
'
'        Exit Sub
'
'    End If
'
'    VtNode.VtNodeUnInit
'
'    Debug.Print "WasUninit:End"
'
'
'End Sub

' =========================================================
' VTNode HTTPA Service 시작
' =========================================================
Public Function StartHttpa() As Boolean

    Const PROC_NAME As String = "StartHttpa"

    Dim httpaConfigFile As String
    Dim ErrorMessage As String
    Dim loopAccepted As Boolean
    Dim errNumber As Long
    Dim errDescription As String
    
    
    On Error GoTo EH

    StartHttpa = False

    VtLogInfo MODULE_NAME, PROC_NAME, "Start"

    ' Native Runtime 초기화 확인
   
    If Not MyVTNode.IsInitialized Then
        MsgBox "VTNode  is not initialized  !!!"
        Exit Function
    End If


    ' Application Home 확인
    If Len(Trim$(MyVTNode.WorkbookHomeDirectory)) = 0 Then

        VtLogError MODULE_NAME, PROC_NAME, _
                   "Application home directory is empty."

        Exit Function
    End If

    ' VTNode Home 확인
    If Len(Trim$(MyVTNode.VTNodeHomeDirectory)) = 0 Then

        VtLogError MODULE_NAME, PROC_NAME, _
                   "VTNODE_HOME is empty."

        Exit Function
    End If

    ' HTTPA 설정 파일 경로
    httpaConfigFile = VtPathUtils.VtPathCombine( _
                          MyVTNode.WorkbookHomeDirectory, _
                          "vtnode_httpa_params.ini")

    If Len(dir$(httpaConfigFile, vbNormal)) = 0 Then

        ErrorMessage = _
            "HTTPA configuration file not found." & _
            vbCrLf & httpaConfigFile

        VtLogError MODULE_NAME, PROC_NAME, ErrorMessage

        MsgBox ErrorMessage, vbCritical, "VTNode"
        Exit Function
    End If

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "---------------------------------------------"

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "Requesting VTNode HTTPA Service start."

    VtLogInfo MODULE_NAME, PROC_NAME, _
              " - Application Home : " & MyVTNode.WorkbookHomeDirectory

    VtLogInfo MODULE_NAME, PROC_NAME, _
              " - VTNode Home      : " & MyVTNode.VTNodeHomeDirectory

    VtLogInfo MODULE_NAME, PROC_NAME, _
              " - Config File      : " & httpaConfigFile

    ' HTTPA Service 시작 요청
    loopAccepted = MyHttpaLooper.BeginLoop( _
                          VTHttpaEventLoopType.VTHTTPA_EVENT_LOOP_CONTINUOUS, _
                          httpaConfigFile, _
                          ErrorMessage)

    If Not loopAccepted Then

        VtLogError MODULE_NAME, PROC_NAME, _
                   "HTTPA Service start request failed. reason=" & ErrorMessage

        MsgBox "HTTPA Service start failed." & vbCrLf & ErrorMessage, vbCritical, "VTNode"

        Exit Function
    End If

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "HTTPA Service start request accepted. message=" & _
              ErrorMessage

'    VtLogInfo MODULE_NAME, PROC_NAME, _
'              "Httpa status=" & _
'              WasHttpaServiceStatusToString( _
'                  WasHttpaService.HttpaStatus)
'
'    VtLogInfo MODULE_NAME, PROC_NAME, _
'              "Loop status=" & _
'              WasHttpaLoopStatusToString( _
'                  WasHttpaService.LoopStatus)


    'WasHttpaService.TraceState "StartHttpa"

    ' True는 시작 요청이 수락되었음을 의미한다.
    

   ' UpdateRuntimeStates
    
    StartHttpa = True
    VtLogInfo MODULE_NAME, PROC_NAME, "End"
    Exit Function

EH:
    errNumber = Err.number
    errDescription = Err.description

    On Error Resume Next

    VtLogError MODULE_NAME, PROC_NAME, _
               "Unexpected error. number=" & _
               CStr(errNumber) & _
               ", description=" & errDescription

    MsgBox "StartHttpa() failed." & vbCrLf & _
           errDescription, vbCritical, "VTNode"

    StartHttpa = False

End Function

Public Sub StopHttpa()

    Const PROC_NAME As String = "StopHttpa"

    Dim ErrorMessage As String
    Dim requestAccepted As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH


    VtLogInfo MODULE_NAME, PROC_NAME, "Start"
    
    'Debug.Print "WasStopHttpa:1"
    
    '
    ' 비정상 상태라면 정상 종료 요청 대신
    ' 사용자에게 강제 종료 여부를 확인한다.
    '
    Call MyHttpaLooper.EndLoop
        
    VtLogInfo MODULE_NAME, PROC_NAME, "End"
'    MyAudio.Play AUDIO_WAS_STOPPED
'    VtWin32.Sleep 1000
'
'    UpdateRuntimeStates
    'Debug.Print "WasStopHttpa:2"

    
    Exit Sub

EH:
    errNumber = Err.number
    errDescription = Err.description

'    VtLogError MODULE_NAME,
'               PROC_NAME,
'               "Stop request failed. number=" _
'               & CStr(errNumber) _
'               & ", description=" & errDescription

End Sub
'Public Function WasConfirmForceTerminate() As Boolean
'    Const PROC_NAME As String = "WasConfirmForceTerminate"

'    Dim prompt As String
'    Dim answer As VbMsgBoxResult
'    Dim ErrorMessage As String

'    On Error GoTo EH

'    WasConfirmForceTerminate = False

'    If Not WasHttpaService.RequiresForceTerminate Then
'        VtLogInfo MODULE_NAME, _
'                  PROC_NAME, _
'                  "Force termination is not required."

'        Exit Function
'    End If

'    prompt = _
'        "VTNode HTTPA Service가 비정상 상태입니다." _
'        & vbCrLf & vbCrLf _
'        & WasHttpaService.GetStatusDescription() _
'        & vbCrLf & vbCrLf _
'        & "강제로 종료하시겠습니까?" _
'        & vbCrLf _
'        & "처리 중인 연결과 요청이 즉시 종료될 수 있습니다."

'    answer = MsgBox( _
'                prompt, _
'                vbYesNo Or vbExclamation Or vbDefaultButton2, _
'                "VTNode 강제 종료 확인")

'    If answer <> vbYes Then
'        VtLogInfo MODULE_NAME, _
'                  PROC_NAME, _
'                  "Force termination cancelled by user."

'        Exit Function
'    End If

'    If Not WasHttpaService.ForceTerminate(ErrorMessage) Then
'        MsgBox _
'            "강제 종료에 실패했습니다." _
'            & vbCrLf & vbCrLf _
'            & ErrorMessage, _
'            vbCritical Or vbOKOnly, _
'            "VTNode 강제 종료 실패"

'        Exit Function
'    End If

'    MsgBox _
'        "VTNode HTTPA Service가 강제로 종료되었습니다.", _
'        vbInformation Or vbOKOnly, _
'        "VTNode 종료 완료"

'    WasConfirmForceTerminate = True
'    Exit Function

'EH:
'    VtLogError MODULE_NAME, _
'               PROC_NAME, _
'               "Unexpected error. number=" _
'               & CStr(Err.Number) _
'               & ", description=" & Err.description

'    WasConfirmForceTerminate = False
'End Function



Public Function WWWRootDirectory() As String

  WWWRootDirectory = VtPathUtils.VtPathCombine(MyVTNode.WorkbookHomeDirectory, "www")
  
End Function

'Public Function WorkbookHomeDirectory() As String
'    WorkbookHomeDirectory = mWorkbookHomeDirectory
'
'
'
'End Function
'
'





