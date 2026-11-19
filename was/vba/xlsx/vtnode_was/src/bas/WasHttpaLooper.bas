Attribute VB_Name = "WasHttpaLooper"

Option Explicit


Private Const MODULE_NAME As String = "WasHttpaLooper"


Private mStartedAt As LongLong
Private mLoopType As WASHttpaEventLoopType
Private mLoopStop As Boolean
Private mLoopRunning As Boolean

Private mPollPending As Boolean
Private mPollProcedure As String
Private mPollNext As Date

Private mLastError As String

Public Enum WASHttpaEventLoopType

        WAS_HTTPA_EVENT_LOOP_UNSET = 0
        WAS_HTTPA_EVENT_LOOP_CONTINUOUS = 1
        WAS_HTTPA_EVENT_LOOP_SCHEDULED = 2
End Enum

Private Const MAX_EVENTS_PER_PUMP As Long = 1000


Private Sub Started(ByVal loopType As WASHttpaEventLoopType)
    
    mLoopStop = False
    mLoopType = loopType
    mStartedAt = VtTimeUtils.VtGetCurrentCTime()
    mLoopRunning = False
    
    mPollPending = False
    mPollProcedure = ""
    mPollNext = 0

    
    WasHttpaEventDispatch.Init
    

    WasAudio.Play AUDIO_WAS_STARTED
    UpdateRuntimeStates

End Sub


Private Sub Stopped()
    
    mLoopStop = True
    mLoopType = WAS_HTTPA_EVENT_LOOP_UNSET
    mStartedAt = 0
    mLoopRunning = False
    
    
    WasAudio.Play AUDIO_WAS_STOPPED
    WasAudio.WaitForComplete
    
    
    UpdateRuntimeStates

    WasHttpaEventDispatch.Uninit
    
    'mPollPending = True
    'mPollProcedure = ""
    'mPollNext = 0
End Sub
Private Sub PumpEvents()
        Dim session As CVtHttpaSession
        Dim count As Long
        Dim number As Long
        Dim description As String


        On Error GoTo EH
        For count = 1 To MAX_EVENTS_PER_PUMP
          Set session = VtHttpaService.HttpaGetServerEvent()

          If session Is Nothing Then Exit For
              WasHttpaEventDispatch.DispatchHttpaEvent session
          Set session = Nothing
          If mLoopStop Then Exit For
        Next count
        
        Exit Sub
EH:
        number = Err.number
        description = Err.description
        Err.Raise number, MODULE_NAME & ".PumpEvents", description

End Sub

' =========================================================
' Continuous Event Loop
' =========================================================
Private Sub RunContinuousLoop()
    
    Const PROC_NAME As String = "RunContinuousLoop"
    
    Dim errNumber As Long
    Dim errDescription As String

    'If mLoopRunning Then Exit Sub

    On Error GoTo EH

    mLoopRunning = True

    'Debug.Print "RunContinuousLoop:Begin"
    VtLogInfo MODULE_NAME, PROC_NAME, "Begin"


    Do While Not mLoopStop

        PumpEvents

        'If mLoopStop Then Exit Do

        DoEvents

        'If mLoopStop Then Exit Do

        VtWin32.Sleep 1
        

    Loop

CleanExit:

    VtLogInfo MODULE_NAME, PROC_NAME, "End"
  '  Stopped

'    If mLoopStop Then
'        Sopped
'    End If

    Exit Sub

EH:
    errNumber = Err.number
    errDescription = Err.description

    On Error Resume Next
  '  Stopped
    
    On Error GoTo 0

    Err.Raise errNumber, MODULE_NAME & ".RunContinuousLoop", errDescription

End Sub

' =========================================================
' Scheduled Timer Callback
' =========================================================
Public Sub StatusTimer()
    mPollPending = False
    If mLoopStop Then Exit Sub
    Call PollStatus

End Sub


Private Sub NextScheduledPoll()

    If mLoopStop Then Exit Sub
    If mPollPending Then Exit Sub

    If mLoopType = WASHttpaEventLoopType.WAS_HTTPA_EVENT_LOOP_SCHEDULED Then
    
        mPollNext = Now + TimeSerial(0, 0, 1)

        mPollProcedure = "'" & Replace(ThisWorkbook.Name, "'", "''") & _
                         "'!WasHttpaLooper.StatusTimer"

        Application.OnTime _
            EarliestTime:=mPollNext, _
            procedure:=mPollProcedure, _
            Schedule:=True

       mPollPending = True

    End If

End Sub

Private Sub CancelScheduledPoll()

    If Not mPollPending Then Exit Sub

    On Error Resume Next

    If mLoopType = WASHttpaEventLoopType.WAS_HTTPA_EVENT_LOOP_SCHEDULED Then

        Application.OnTime _
            EarliestTime:=mPollNext, _
            procedure:=mPollProcedure, _
            Schedule:=False

    End If

    On Error GoTo 0

    mPollPending = False

End Sub


' =========================================================
' Event Polling
' =========================================================
Private Sub PollStatus()

    Dim errNumber As Long
    Dim errDescription As String

    If mLoopStop Then Exit Sub
    If mLoopRunning Then Exit Sub

    '
    If mLoopType = WAS_HTTPA_EVENT_LOOP_CONTINUOUS Then
        RunContinuousLoop
        Exit Sub
    End If

    If mLoopType <> WAS_HTTPA_EVENT_LOOP_SCHEDULED Then Exit Sub

    On Error GoTo EH



    ' do scheduled job
    mLoopRunning = True
    PumpEvents
    If Not mLoopStop Then
        NextScheduledPoll
    End If
    mLoopRunning = False
    ' end of scheduled job

'    If mLoopStop Then
'        Stopped
'    End If

    Exit Sub
    

EH:
    errNumber = Err.number
    errDescription = Err.description

    On Error Resume Next
   ' Stopped
    On Error GoTo 0

    Err.Raise errNumber, MODULE_NAME & ".PollStatus", errDescription

End Sub


Private Function IsValidVariablesInitialized() As Boolean
    ' Check if the state variables are initialized
        
    Debug.Print "IsValidVariablesInitialized=" & mStartedAt & "," & mLoopType
         
        
    If mStartedAt = 0 Or mLoopType = WAS_HTTPA_EVENT_LOOP_UNSET Then
        IsValidVariablesInitialized = False
    Else
        IsValidVariablesInitialized = True
    End If
    
End Function

'Private Sub InitModules()
'        ' Initialize the dependent modules
'        WasHttpaClient.Init
'        WasHttpaError.Init
'        WasHttpaNet.Init
'        WasHttpaServerEvent.Init
'        WasHttpaWebSocket.Init
'End Sub
'
'Private Sub SoppedModules()
'    ' Soppedialize the dependent modules
'    WasHttpaWebSocket.Sopped
'    WasHttpaServerEvent.Sopped
'    WasHttpaNet.Sopped
'    WasHttpaError.Sopped
'    WasHttpaClient.Sopped
'End Sub


Public Function BeginLoop(ByVal loopType As WASHttpaEventLoopType, ByVal configFile As String, ByRef eout As String) As Boolean


    Const PROC_NAME As String = "BeginLoop"
 
    Dim httpaStartResult As VTHttpaEngineStartResult
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH


    BeginLoop = False

    ' Check if the required variables are initialized
    If Len(Trim$(configFile)) = 0 Then
        eout = "Configuration file path is empty."
        Exit Function
    End If

    ' ceck if the configuration file exists
    If Len(dir$(configFile, vbNormal)) = 0 Then
        eout = "Configuration file not found. path=" & configFile
        Exit Function
    End If


    ' check if the loopType is valid
    If loopType <> WAS_HTTPA_EVENT_LOOP_CONTINUOUS And loopType <> WAS_HTTPA_EVENT_LOOP_SCHEDULED Then
        eout = "Invalid loop type. loopType=" & CStr(loopType)
        Exit Function
    End If


'        VtLogInfo MODULE_NAME, _
'                          PROC_NAME, _
'                          "Start requested. vtnodeHomeDir=" & vtnodeHomeDir _
'                          & ", vbaHomeDir=" & vbaHomeDir _
'                          & ", configFile=" & configFile


    If VtHttpaService.HttpaIsRunning Then
        If IsValidVariablesInitialized() Then
            BeginLoop = True
            eout = "Service is already running"
            Exit Function
        End If
        ' Uncondintionally reset the state variables to avoid infinite loop
        Call VtHttpaService.HttpaStop
    End If


    httpaStartResult = VtHttpaService.HttpaStart(configFile)

    If Not httpaStartResult.success Then
        BeginLoop = False
        eout = "Service start failed. error=" & httpaStartResult.emsg
        Exit Function
    End If



    Started loopType

    Call PollStatus

    BeginLoop = httpaStartResult.success

    Exit Function

EH:
    errNumber = Err.number
    errDescription = Err.description

    'mHttpaStatus = WAS_HTTPA_SERVICE_FAULTED
    'mLoopStatus = WAS_HTTPA_LOOP_FAULTED
    mLastError = errDescription
    eout = errDescription

    VtLogError MODULE_NAME, _
                           PROC_NAME, _
                           "Unexpected error. number=" & CStr(errNumber) _
                           & ", description=" & errDescription

    BeginLoop = False

End Function

'Private Sub FinalizeLoop()
'
'    Const PROC_NAME As String = "FinalizeLoop"
'
'    mLoopStop = True
'    mLoopType = WAS_HTTPA_EVENT_LOOP_UNSET
'    mStartedAt = 0
'    mLoopRunning = False
'    mPollPending = False
'
'    Debug.Print "FinalizeLoop:1"
'
'    'VtHttpaService.HttpaStop
'    'SoppedModules
'
'    Debug.Print "FinalizeLoop:2"
'    WasAudio.Play AUDIO_WAS_STOPPED
'    VtWin32.Sleep 1000
'
'    VtLogInfo MODULE_NAME, PROC_NAME, "Loop End"
'
'End Sub

Public Sub EndLoop()
        
    Const PROC_NAME As String = "EndLoop"
        
       
    VtLogInfo MODULE_NAME, PROC_NAME, "Begin"
                                
    If IsValidVariablesInitialized() Then

        If mLoopType = WAS_HTTPA_EVENT_LOOP_SCHEDULED Then
            CancelScheduledPoll
        End If
                                                                
    Else
        VtLogWarn MODULE_NAME, PROC_NAME, "LoopEnd called but state variables are not initialized."
    End If
                
                
    ' Force Stop Httpa Service
    VtHttpaService.HttpaStop
    
    Stopped
    
    VtLogInfo MODULE_NAME, PROC_NAME, "End"
        
End Sub


Public Function IsLooping() As Boolean


    IsLooping = VtHttpaService.HttpaIsRunning() Or _
                IsValidVariablesInitialized() Or _
                 mLoopRunning
     
     
'     If IsValidVariablesInitialized() Or VtHttpaService.HttpaIsRunning() Then
'
'       IsLooping = True
'     Else
'       IsLooping = False
'
'     End If
'


End Function


