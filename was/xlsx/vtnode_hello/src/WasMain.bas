Attribute VB_Name = "WasMain"

Option Explicit

Private Const MODULE_NAME As String = "WasMain"
Private mHomeDirectory As String
Private mVTNodeHomeDirectory As String
Private mInitialized As Boolean

Private mExcelEvents As CWasExcelEvents



' =========================================================
' Excel Application 이벤트 모니터 시작
' =========================================================
Public Sub BeginExcelCloseMonitor()

    On Error GoTo EH

    If mExcelEvents Is Nothing Then
        Set mExcelEvents = New CWasExcelEvents
        Set mExcelEvents.App = Application
    End If

    Debug.Print "BeginExcelCloseMonitor: success"
    Exit Sub

EH:
    MsgBox "Excel event monitor failed." & vbCrLf & _
           Err.description, vbExclamation, "VTNode"

End Sub


' =========================================================
' Excel Application 이벤트 모니터 해제
' =========================================================
Public Sub EndExcelCloseMonitor()

    On Error Resume Next

    If Not mExcelEvents Is Nothing Then
        Set mExcelEvents.App = Nothing
        Set mExcelEvents = Nothing
    End If

End Sub
' =========================================================
' VTNode Native Runtime 초기화
' =========================================================
Public Function WasInit() As Boolean

    Const PROC_NAME As String = "WasInit"

    Dim LogDirectory As String
    Dim initResult As VTNodeInitResult
    Dim verifyResult As VTNodeVerifyResult
    Dim nativeInitialized As Boolean

    On Error GoTo EH

    WasInit = False

    ' 이미 초기화되어 있으면 재초기화하지 않는다.
    If mInitialized Then
        WasInit = True
        Exit Function
    End If

    mHomeDirectory = ThisWorkbook.Path
    mVTNodeHomeDirectory = _
        VtEnv.VtGetEnv("VTNODE_HOME", vbNullString)

    ' Application Home 확인
    If Len(Trim$(mHomeDirectory)) = 0 Then
        MsgBox "Application Home directory is empty.", _
               vbCritical, "VTNode Initialization"
        Exit Function
    End If

    ' VTNode Home 확인
    If Len(Trim$(mVTNodeHomeDirectory)) = 0 Then
        MsgBox "VTNODE_HOME 환경변수가 설정되지 않았습니다.", _
               vbCritical, "VTNode Initialization"
        Exit Function
    End If

    LogDirectory = mHomeDirectory & _
                   Application.PathSeparator & "logs"

    VtLogInitialize LogDirectory

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "WasInit() started. VTNODE_HOME=" & _
              mVTNodeHomeDirectory

    ' Native Runtime 초기화
    initResult = VtNode.VtNodeInit()

    If Not initResult.success Then

        VtLogError MODULE_NAME, PROC_NAME, _
                   "VtNodeInit() failed. reason=" & _
                   initResult.emsg

        MsgBox "VtNodeInit() failed." & vbCrLf & _
               initResult.emsg, vbCritical, "VTNode"

        GoTo InitFailed
    End If

    nativeInitialized = True

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "VtNodeInit() success"

    ' Runtime 설치 및 환경 검증
    verifyResult = VtNode.VtNodeVerify(mVTNodeHomeDirectory)

    If Not verifyResult.success Then

        VtLogError MODULE_NAME, PROC_NAME, _
                   "VtNodeVerify() failed. reason=" & _
                   verifyResult.emsg

        MsgBox "VtNodeVerify() failed." & vbCrLf & _
               verifyResult.emsg, vbCritical, "VTNode"

        GoTo InitFailed
    End If

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "VtNodeVerify() success"

    mInitialized = True
    WasInit = True

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "WasInit() completed successfully"

    Exit Function

InitFailed:
    On Error Resume Next

    If nativeInitialized Then
        VtNode.VtNodeUnInit
    End If

    mInitialized = False
    VtLogClose
    Exit Function

EH:
    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.Number
    errDescription = Err.description

    MsgBox "WasInit() unexpected error." & vbCrLf & _
           "Number: " & CStr(errNumber) & vbCrLf & _
           "Description: " & errDescription, _
           vbCritical, "VTNode"

    Resume InitFailed

End Function


' =========================================================
' VTNode HTTPA Service 시작
' =========================================================
Public Function WasStartHttpa() As Boolean

    Const PROC_NAME As String = "WasStartHttpa"

    Dim httpaConfigFile As String
    Dim ErrorMessage As String
    Dim requestAccepted As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    WasStartHttpa = False

    ' Native Runtime 초기화 확인
    If Not mInitialized Then

        If Not WasInit() Then
            Exit Function
        End If

    End If

    ' Application Home 확인
    If Len(Trim$(mHomeDirectory)) = 0 Then

        VtLogError MODULE_NAME, PROC_NAME, _
                   "Application home directory is empty."

        Exit Function
    End If

    ' VTNode Home 확인
    If Len(Trim$(mVTNodeHomeDirectory)) = 0 Then

        VtLogError MODULE_NAME, PROC_NAME, _
                   "VTNODE_HOME is empty."

        Exit Function
    End If

    ' HTTPA 설정 파일 경로
    httpaConfigFile = VtPathUtils.VtPathCombine( _
                          mHomeDirectory, _
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
              " - Application Home : " & mHomeDirectory

    VtLogInfo MODULE_NAME, PROC_NAME, _
              " - VTNode Home      : " & mVTNodeHomeDirectory

    VtLogInfo MODULE_NAME, PROC_NAME, _
              " - Config File      : " & httpaConfigFile

    ' HTTPA Service 시작 요청
    requestAccepted = WasHttpaService.RequestStart( _
                          mVTNodeHomeDirectory, _
                          mHomeDirectory, _
                          httpaConfigFile, _
                          ErrorMessage)

    If Not requestAccepted Then

        VtLogError MODULE_NAME, PROC_NAME, _
                   "HTTPA Service start request failed. reason=" & _
                   ErrorMessage

        MsgBox "HTTPA Service start failed." & vbCrLf & _
               ErrorMessage, vbCritical, "VTNode"

        Exit Function
    End If

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "HTTPA Service start request accepted. message=" & _
              ErrorMessage

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "Httpa status=" & _
              WasHttpaServiceStatusToString( _
                  WasHttpaService.HttpaStatus)

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "Loop status=" & _
              WasHttpaLoopStatusToString( _
                  WasHttpaService.LoopStatus)

    WasHttpaService.TraceState "WasStartHttpa"

    ' True는 시작 요청이 수락되었음을 의미한다.
    WasStartHttpa = True
    Exit Function

EH:
    errNumber = Err.Number
    errDescription = Err.description

    On Error Resume Next

    VtLogError MODULE_NAME, PROC_NAME, _
               "Unexpected error. number=" & _
               CStr(errNumber) & _
               ", description=" & errDescription

    MsgBox "WasStartHttpa() failed." & vbCrLf & _
           errDescription, vbCritical, "VTNode"

    WasStartHttpa = False

End Function
Public Function WasStopHttpa() As Boolean
    Const PROC_NAME As String = "WasStopHttpa"

    Dim ErrorMessage As String
    Dim requestAccepted As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    WasStopHttpa = False

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Stop requested. serviceStatus=" _
              & WasHttpaServiceStatusToString( _
                    WasHttpaService.HttpaStatus) _
              & ", loopStatus=" _
              & WasHttpaLoopStatusToString( _
                    WasHttpaService.LoopStatus)

    '
    ' 비정상 상태라면 정상 종료 요청 대신
    ' 사용자에게 강제 종료 여부를 확인한다.
    '
    If WasHttpaService.RequiresForceTerminate Then
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "Abnormal HTTPA Service state detected."

        WasStopHttpa = WasConfirmForceTerminate()
        Exit Function
    End If

    '
    ' 정상적인 비동기 종료 요청
    '
    
    requestAccepted = _
        WasHttpaService.RequestStop(ErrorMessage)

   
    
    If Not requestAccepted Then
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "Stop request rejected. reason=" _
                   & ErrorMessage

        '
        ' RequestStop 처리 과정에서 상태가 FAULTED로
        ' 변경됐을 가능성을 다시 검사한다.
        '
        If WasHttpaService.RequiresForceTerminate Then
            WasStopHttpa = WasConfirmForceTerminate()
        End If

        Exit Function
    End If

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Stop request accepted. message=" _
              & ErrorMessage

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Current httpaStatus=" _
              & WasHttpaServiceStatusToString(WasHttpaService.HttpaStatus) _
              & ", loopStatus=" _
              & WasHttpaLoopStatusToString(WasHttpaService.LoopStatus)

   
    WasStopHttpa = True
    Exit Function

EH:
    errNumber = Err.Number
    errDescription = Err.description

    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "Stop request failed. number=" _
               & CStr(errNumber) _
               & ", description=" & errDescription

    If WasHttpaService.RequiresForceTerminate Then
        WasStopHttpa = WasConfirmForceTerminate()
    Else
        WasStopHttpa = False
    End If
End Function
Public Function WasConfirmForceTerminate() As Boolean
    Const PROC_NAME As String = "WasConfirmForceTerminate"

    Dim prompt As String
    Dim answer As VbMsgBoxResult
    Dim ErrorMessage As String

    On Error GoTo EH

    WasConfirmForceTerminate = False

    If Not WasHttpaService.RequiresForceTerminate Then
        VtLogInfo MODULE_NAME, _
                  PROC_NAME, _
                  "Force termination is not required."

        Exit Function
    End If

    prompt = _
        "VTNode HTTPA Service가 비정상 상태입니다." _
        & vbCrLf & vbCrLf _
        & WasHttpaService.GetStatusDescription() _
        & vbCrLf & vbCrLf _
        & "강제로 종료하시겠습니까?" _
        & vbCrLf _
        & "처리 중인 연결과 요청이 즉시 종료될 수 있습니다."

    answer = MsgBox( _
                prompt, _
                vbYesNo Or vbExclamation Or vbDefaultButton2, _
                "VTNode 강제 종료 확인")

    If answer <> vbYes Then
        VtLogInfo MODULE_NAME, _
                  PROC_NAME, _
                  "Force termination cancelled by user."

        Exit Function
    End If

    If Not WasHttpaService.ForceTerminate(ErrorMessage) Then
        MsgBox _
            "강제 종료에 실패했습니다." _
            & vbCrLf & vbCrLf _
            & ErrorMessage, _
            vbCritical Or vbOKOnly, _
            "VTNode 강제 종료 실패"

        Exit Function
    End If

    MsgBox _
        "VTNode HTTPA Service가 강제로 종료되었습니다.", _
        vbInformation Or vbOKOnly, _
        "VTNode 종료 완료"

    WasConfirmForceTerminate = True
    Exit Function

EH:
    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "Unexpected error. number=" _
               & CStr(Err.Number) _
               & ", description=" & Err.description

    WasConfirmForceTerminate = False
End Function



Public Function WasWWWRootDirectory() As String

  WasWWWRootDirectory = VtPathUtils.VtPathCombine(mHomeDirectory, "www")
  
End Function

Public Function WasHomeDirectory() As String
    WasHomeDirectory = mHomeDirectory
    
 
  
End Function






