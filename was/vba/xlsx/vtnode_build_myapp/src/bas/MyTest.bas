Attribute VB_Name = "MyTest"
Option Explicit

Public Sub MyTest_Init()

  Dim r As Boolean
  
  r = MyVTNode.Initialize()
  
  Debug.Print "MyVTNode.Initialize=" & r
  
  
    
End Sub

Public Sub MyTest_Uninit()

   MyVTNode.Finalize
  Debug.Print "MyVTNode.Finalize"
   
    
End Sub

Public Sub MyTest_StartHttpa()

Dim r As Boolean
   
   r = MyHttpa.StartHttpa()
   
   Debug.Print "Start.httpa=" & r
   

   
End Sub

Public Sub MyTest_StopHttpa()

   MyHttpa.StopHttpa
   
End Sub


Public Sub MyTest_Play()

  'MyAudio.Play AUDIO_WAS_STOPPED
   Dim errorOut As String
   Dim waveFilePath As String
   
   waveFilePath = "E:\xlibc\test\media\16khz.wav"
   
   If Not VtAudioPlayerStart(waveFilePath, errorOut) Then

      Debug.Print "Audio play failed. File=" & waveFilePath & ", Error=" & errorOut


    End If

   If Not VtAudioPlayerStart(waveFilePath, errorOut) Then

      Debug.Print "Audio play failed. File=" & waveFilePath & ", Error=" & errorOut


    End If



End Sub

