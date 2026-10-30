Attribute VB_Name = "WasAudio"
Option Explicit

Private Const MODULE_NAME As String = "WasAudio"

Private Const AUDIO_FOLDER As String = "audio"


' ============================================================
' Table Order Audio Type
' ============================================================
Public Enum TableOrderAudioType

    AUDIO_NONE = 0

    ' UI
    AUDIO_TICK
    AUDIO_CLICK_CONFIRM

    ' Order
    AUDIO_ORDER_RECEIVED

    ' Staff
    AUDIO_STAFF_CALL

    ' Payment
    AUDIO_PAYMENT_COMPLETE
    
    AUDIO_WAS_STARTED
    AUDIO_WAS_STOPPED
    
    AUDIO_REQUEST_RECEIVED
    

End Enum




' ============================================================
' Play
' ============================================================
Public Function Play(ByVal audioType As TableOrderAudioType) As Boolean

    Const PROC_NAME As String = "Play"

    On Error GoTo EH

    Dim waveFilePath As String
    Dim errorOut As String

    Play = False

    If audioType = AUDIO_NONE Then
        Exit Function
    End If

    waveFilePath = GetAudioFilePath(audioType)

    If Len(waveFilePath) = 0 Then

        VtLogError _
            MODULE_NAME, _
            PROC_NAME, _
            "Invalid AudioType=" & CStr(audioType)

        Exit Function

    End If

    ' WAV 파일 존재 확인
    If Len(dir$(waveFilePath)) = 0 Then

        VtLogError _
            MODULE_NAME, _
            PROC_NAME, _
            "Audio file not found: " & waveFilePath

        Exit Function

    End If

    '
    ' VTNode Audio Player
    '
    ' 새로운 Play 요청이 들어오면
    ' 현재 재생 중인 Audio는 VTNode wrapper에서 중지 후 새로 재생.
    '
    If Not VtAudioPlayerStart(waveFilePath, errorOut) Then

        VtLogError _
            MODULE_NAME, _
            PROC_NAME, _
            "Audio play failed. File=" & waveFilePath & _
            ", Error=" & errorOut

        Exit Function

    End If

    Play = True

    Exit Function

EH:

    VtLogError _
        MODULE_NAME, _
        PROC_NAME, _
        "Runtime Error=" & Err.Number & _
        ", " & Err.description

    Play = False

End Function




' ============================================================
' Audio Type -> WAV File Path
' ============================================================
Private Function GetAudioFilePath( _
    ByVal audioType As TableOrderAudioType) As String

    Dim fileName As String

    Select Case audioType

        Case AUDIO_TICK
            fileName = "audio_tick.wav"

        Case AUDIO_CLICK_CONFIRM
            fileName = "audio_click_confirm.wav"

        Case AUDIO_ORDER_RECEIVED
            fileName = "audio_order_received.wav"

        Case AUDIO_STAFF_CALL
            fileName = "audio_staff_call.wav"

        Case AUDIO_PAYMENT_COMPLETE
            fileName = "audio_payment_request.wav"

        Case AUDIO_WAS_STARTED
            fileName = "audio_was_started.wav"
        
        Case AUDIO_WAS_STOPPED
            fileName = "audio_was_stopped.wav"
        Case AUDIO_REQUEST_RECEIVED
            fileName = "audio_request_received.wav"
            
        
        Case Else
            GetAudioFilePath = vbNullString
            Exit Function

    End Select

    GetAudioFilePath = _
        ThisWorkbook.path & "\" & _
        AUDIO_FOLDER & "\" & _
        fileName

End Function


Public Function WaitForComplete( _
    Optional ByVal timeoutMs As Long = 5000) As Boolean

    Const PROC_NAME As String = "WaitForComplete"
    Const POLL_INTERVAL_MS As Long = 50

    Dim elapsedMs As Long

    On Error GoTo EH

    WaitForComplete = False

    ' 이미 재생이 끝났으면 즉시 성공
    If Not VtAudioPlayerIsPlaying() Then
        WaitForComplete = True
        Exit Function
    End If

    ' --------------------------------------------------------
    ' Audio 재생 완료 대기
    ' --------------------------------------------------------
    Do While VtAudioPlayerIsPlaying()

        VtWin32.Sleep POLL_INTERVAL_MS
        DoEvents

        elapsedMs = elapsedMs + POLL_INTERVAL_MS

        If timeoutMs > 0 Then

            If elapsedMs >= timeoutMs Then

                Call VtLogError( _
                    MODULE_NAME, _
                    PROC_NAME, _
                    "Audio wait timeout. TimeoutMs=" & CStr(timeoutMs))

                Exit Function

            End If

        End If

    Loop

    WaitForComplete = True

    Exit Function


EH:

    Call VtLogError( _
        MODULE_NAME, _
        PROC_NAME, _
        "Runtime Error=" & Err.Number & _
        ", " & Err.description)

    WaitForComplete = False

End Function
