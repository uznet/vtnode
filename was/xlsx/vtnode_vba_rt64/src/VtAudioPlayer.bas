Attribute VB_Name = "VtAudioPlayer"

Option Explicit

#If VBA7 And Win64 Then

#If VTNODE_DEBUG = 1 Then  ' Debug Mode
Private Declare PtrSafe Function vtnode_audio_player_start Lib "vtnode64d.dll" ( _
    ByVal waveFile As LongPtr, _
    ByVal eout As LongPtr, _
    ByVal esize As Long _
) As Long

Private Declare PtrSafe Function vtnode_audio_player_is_playing Lib "vtnode64d.dll" () As Long

Private Declare PtrSafe Sub vtnode_audio_player_stop Lib "vtnode64d.dll" ()

#Else

Private Declare PtrSafe Function vtnode_audio_player_start Lib "vtnode64.dll" ( _
    ByVal waveFile As LongPtr, _
    ByVal eout As LongPtr, _
    ByVal esize As Long _
) As Long

Private Declare PtrSafe Function vtnode_audio_player_is_playing Lib "vtnode64.dll" () As Long

Private Declare PtrSafe Sub vtnode_audio_player_stop Lib "vtnode64.dll" ()

#End If '
#End If ' VBA7



Public Function VtAudioPlayerStart( _
    ByVal waveFilePath As String, _
    ByRef errorOut As String) As Boolean

    On Error GoTo EH

    Dim Utf8Path() As Byte
    Dim ErrorBuffer() As Byte
    Dim result As Long

    errorOut = vbNullString
    VtAudioPlayerStart = False

    ' --------------------------------------------------------
    ' 1. Stop current playback
    ' --------------------------------------------------------
    Call vtnode_audio_player_stop

    ' --------------------------------------------------------
    ' 2. Validate WAV file path
    ' --------------------------------------------------------
    If Len(waveFilePath) = 0 Then
        errorOut = "Wave file path is empty."
        Exit Function
    End If

    If dir$(waveFilePath) = vbNullString Then
        errorOut = "Wave file not found: " & waveFilePath
        Exit Function
    End If

    ' --------------------------------------------------------
    ' 3. Convert file path to UTF-8
    '    기존 VTNode UTF-8 변환 함수가 있다면 그것을 사용
    ' --------------------------------------------------------
    Utf8Path = VtCharUtils.VtStringToUtf8Bytes(waveFilePath)
    
    ' --------------------------------------------------------
    ' 4. Prepare native error buffer
    ' --------------------------------------------------------
    ReDim ErrorBuffer(0 To 511)

    ' --------------------------------------------------------
    ' 5. Start asynchronous WAV playback
    ' --------------------------------------------------------
    result = vtnode_audio_player_start( _
                VarPtr(Utf8Path(0)), _
                VarPtr(ErrorBuffer(0)), _
                UBound(ErrorBuffer) + 1)

    If result = 0 Then
        errorOut = VtUtf8BytesToString(ErrorBuffer)
        Exit Function
    End If

    VtAudioPlayerStart = True
    Exit Function

EH:
    errorOut = Err.Description
    VtAudioPlayerStart = False

End Function


Public Function VtAudioPlayerIsPlaying() As Boolean

    Dim result As Long
    result = vtnode_audio_player_is_playing()
    
    VtAudioPlayerIsPlaying = result <> 0
    
    
    
    
   

End Function
