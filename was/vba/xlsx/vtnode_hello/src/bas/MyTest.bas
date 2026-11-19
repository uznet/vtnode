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

  MyAudio.Play AUDIO_STAFF_CALL

End Sub

