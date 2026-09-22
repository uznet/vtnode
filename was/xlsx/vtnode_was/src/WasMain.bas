Attribute VB_Name = "WasMain"

Option Explicit

Private Const MODULE_NAME As String = "WasMain"
Private mHomeDirectory As String
Private mVTNodeHomeDirectory As String

Private mExcelEvents As CWasExcelEvents




Public Sub BeginExcelCloseMonitor()

      Debug.Print "BeginExcelCloseMonitor:1"
    If mExcelEvents Is Nothing Then
      Debug.Print "BeginExcelCloseMonitor:2"
      Set mExcelEvents = New CWasExcelEvents
      Set mExcelEvents.App = Application
    End If
    
End Sub
Public Function WasInit() As Boolean
    Const PROC_NAME As String = "WasInit"

    Dim logDirectory As String
    Dim initResult As VTNodeInitResult
    Dim verifyResult As VTNodeVerifyResult

    On Error GoTo EH

    
    Call BeginExcelCloseMonitor
    
    
    WasInit = False
    WasUILogger.InitLog
    
    mHomeDirectory = ThisWorkbook.path
    mVTNodeHomeDirectory = VtEnv.VtGetEnv("VTNODE_HOME", vbNullString)

    If Len(Trim$(mVTNodeHomeDirectory)) = 0 Then
        MsgBox "VTNODE_HOME 환경변수가 설정되지 않았습니다.", _
               vbCritical, _
               "VTNode Initialization"

        Exit Function
    End If

    logDirectory = mHomeDirectory _
                 & Application.PathSeparator _
                 & "logs"

    VtLogInitialize logDirectory

    VtLogInfo MODULE_NAME, _
                PROC_NAME, _
                "WasInit() started. VTNODE_HOME=" & mVTNodeHomeDirectory

    initResult = VtNode.VtNodeInit()

    If Not initResult.success Then
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "VtNodeInit() failed. reason=" & initResult.emsg

        Exit Function
    End If

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "VtNodeInit() success"

    verifyResult = VtNode.VtNodeVerify(mVTNodeHomeDirectory)

    If Not verifyResult.success Then
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "VtNodeVerify() failed. reason=" & verifyResult.emsg

        Exit Function
    End If

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "VtNodeVerify() success"

    WasInit = True
    Exit Function

EH:
    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "Unexpected error. number=" & CStr(Err.Number) _
               & ", description=" & Err.description

    WasInit = False
End Function

Public Sub WasUninit()
    Const PROC_NAME As String = "WasUninit"

    On Error GoTo EH

    VtLogInfo MODULE_NAME, PROC_NAME, "WasUninit() started"

    VtNode.VtNodeUnInit

    VtLogInfo MODULE_NAME, PROC_NAME, "VtNodeUnInit() success"

CleanUp:
    VtLogClose
    Exit Sub

EH:
    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "Unexpected error. number=" & CStr(Err.Number) _
               & ", description=" & Err.description

    Resume CleanUp
End Sub

Public Function WasStartHttpa() As Boolean
    Const PROC_NAME As String = "WasStartHttpa"

    Dim httpaConfigFile As String
    Dim errorMessage As String
    Dim requestAccepted As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    WasStartHttpa = False

    WasUILogger.InitLog

    '
    ' Application Home 검증
    '
    If Len(Trim$(mHomeDirectory)) = 0 Then
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "Application home directory is empty."

        Exit Function
    End If

    '
    ' VTNode Home 검증
    '
    If Len(Trim$(mVTNodeHomeDirectory)) = 0 Then
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "VTNODE_HOME is empty."

        Exit Function
    End If

    '
    ' HTTPA 설정 파일 경로 생성
    '
    httpaConfigFile = VtPathUtils.VtPathCombine( _
                            mHomeDirectory, _
                            "vtnode_httpa_params.ini")

    If Len(dir$(httpaConfigFile, vbNormal)) = 0 Then
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "HTTPA configuration file not found. path=" _
                   & httpaConfigFile

        Exit Function
    End If

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "---------------------------------------------"

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Requesting VTNode HTTPA Service start."

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              " - Application Home : " & mHomeDirectory

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              " - VTNode Home      : " & mVTNodeHomeDirectory

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              " - Config File      : " & httpaConfigFile

    '
    ' RequestStart의 True는 HTTPA Engine 시작 및
    ' Event Loop 예약 요청이 정상 처리됐다는 의미다.
    '
    
    requestAccepted = WasHttpaService.RequestStart( _
                            mVTNodeHomeDirectory, _
                            mHomeDirectory, _
                            httpaConfigFile, _
                            errorMessage)

    If Not requestAccepted Then
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "HTTPA Service start request failed. reason=" _
                   & errorMessage

        Exit Function
    End If

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "HTTPA Service start request accepted. message=" _
              & errorMessage

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Httpa status=" _
              & WasHttpaServiceStatusToString( _
                    WasHttpaService.HttpaStatus)

    VtLogInfo MODULE_NAME, _
              PROC_NAME, _
              "Loop status=" _
              & WasHttpaLoopStatusToString( _
                    WasHttpaService.LoopStatus)

  '  WasHttpaService.UpdateRuntimeStates
    
    WasHttpaService.TraceState "WasStartHttpa"
    WasStartHttpa = True
    Exit Function

EH:
    errNumber = Err.Number
    errDescription = Err.description

    VtLogError MODULE_NAME, _
               PROC_NAME, _
               "Unexpected error. number=" & CStr(errNumber) _
               & ", description=" & errDescription

    WasStartHttpa = False
End Function


Public Function WasStopHttpa() As Boolean
    Const PROC_NAME As String = "WasStopHttpa"

    Dim errorMessage As String
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
        WasHttpaService.RequestStop(errorMessage)

   
    
    If Not requestAccepted Then
        VtLogError MODULE_NAME, _
                   PROC_NAME, _
                   "Stop request rejected. reason=" _
                   & errorMessage

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
              & errorMessage

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
    Dim errorMessage As String

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

    If Not WasHttpaService.ForceTerminate(errorMessage) Then
        MsgBox _
            "강제 종료에 실패했습니다." _
            & vbCrLf & vbCrLf _
            & errorMessage, _
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






