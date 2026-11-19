Attribute VB_Name = "WasVTNode"
Option Explicit

Private Const MODULE_NAME As String = "WasVTNode"


Private mWorkbookHomeDirectory As String
Private mVTNodeHomeDirectory As String



' =========================================================
' VTNode Native Runtime 초기화
' =========================================================

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

Public Function IsSupportedOfficeArch() As Boolean

  
#If Win64 Then
   IsSupportedOfficeArch = True
#Else
   IsSupportedOfficeArch = False
#End If


     
End Function

Public Function Initialize() As Boolean

    Const PROC_NAME As String = "Initialize"

    Dim LogDirectory As String
    Dim LogFilePrefix As String
    Dim initResult As VTNodeInitResult
    Dim verifyResult As VTNodeVerifyResult

    On Error GoTo EH

    Initialize = VtNode.VtNodeIsInitialized()
    ' 이미 초기화되어 있으면 재초기화하지 않는다.
    If Initialize Then
        Exit Function
    End If


    mWorkbookHomeDirectory = ThisWorkbook.path
    mVTNodeHomeDirectory = VtEnv.VtGetEnv("VTNODE_HOME", vbNullString)

    ' Application Home 확인
    If Len(Trim$(mWorkbookHomeDirectory)) = 0 Then
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

    LogDirectory = mWorkbookHomeDirectory & Application.PathSeparator & "logs"

    LogFilePrefix = ThisWorkbook.Name
    If InStrRev(LogFilePrefix, ".") > 1 Then
        LogFilePrefix = left$(LogFilePrefix, InStrRev(LogFilePrefix, ".") - 1)
    End If
    
    VtLog.VtLogInitialize LogDirectory, LogFilePrefix

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "Init() started. VTNODE_HOME=" & _
              mVTNodeHomeDirectory

    ' Native Runtime 초기화
    initResult = VtNode.VtNodeInit(mWorkbookHomeDirectory)

    If Not initResult.success Then

        VtLogError MODULE_NAME, PROC_NAME, _
                   "VtNodeInit() failed. reason=" & _
                   initResult.emsg

        MsgBox "VtNodeInit() failed." & vbCrLf & _
               initResult.emsg, vbCritical, "VTNode"

        GoTo InitFailed
    End If


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

    Initialize = VtNode.VtNodeIsInitialized()
    

    VtLogInfo MODULE_NAME, PROC_NAME, _
              "Init() completed successfully"

    Exit Function

InitFailed:
    On Error Resume Next

    If VtNode.VtNodeIsInitialized() Then
        VtNode.VtNodeUnInit
    End If

    VtLogClose
    Exit Function

EH:
    Dim errNumber As Long
    Dim errDescription As String

    errNumber = Err.number
    errDescription = Err.description

    MsgBox "Init() unexpected error." & vbCrLf & _
           "Number: " & CStr(errNumber) & vbCrLf & _
           "Description: " & errDescription, _
           vbCritical, "VTNode"

    Resume InitFailed

End Function

Public Function IsInitialized() As Boolean


    IsInitialized = VtNode.VtNodeIsInitialized()
    
End Function

Public Sub Finalize()
    
    Const PROC_NAME As String = "Finalize"
    VtLogInfo MODULE_NAME, PROC_NAME, "Finalizing..."
    If VtNode.VtNodeIsInitialized() Then
    
        VtNode.VtNodeUnInit
        VtLogInfo MODULE_NAME, PROC_NAME, "Finalized"
    Else
        VtLogInfo MODULE_NAME, PROC_NAME, "Not Initialized"
    
    End If
    
    VtLogClose
                 

End Sub



Public Function WorkbookHomeDirectory() As String

    WorkbookHomeDirectory = mWorkbookHomeDirectory
End Function

Public Function VTNodeHomeDirectory() As String
   
   VTNodeHomeDirectory = mVTNodeHomeDirectory
   

End Function





