Attribute VB_Name = "MyHttpaErrorEvent"
Option Explicit

Private Const MODULE_NAME As String = "MyHttpaErrorEvent"
'Private mFormVTNode As frmVTNode
Private mStartedAt As LongLong

Public Sub Init()


   mStartedAt = VtTimeUtils.VtGetCurrentCTime()
   
   
End Sub


Public Sub Uninit()

   mStartedAt = 0
   

End Sub
Public Sub HandleError(ByVal httpaSession As CVtHttpaSession)
     Const PROC_NAME As String = "HandleError"
     
     Call VtLogInfo(MODULE_NAME, PROC_NAME, httpaSession.ErrorMsg)
     
     'Call UITabLog.WriteLog("    ERROR  : " & httpaSession.ErrorMsg)
End Sub

Public Sub HandleNetError(ByVal httpaSession As CVtHttpaSession)

     Const PROC_NAME As String = "HandleNetError"
     
     Dim CustomData As CMyCustomData
     Set CustomData = httpaSession.CustomData
     

     Call VtLogInfo(MODULE_NAME, PROC_NAME, httpaSession.ErrorMsg)

     'Call UITabLog.WriteLog("    ERROR  : " & httpaSession.ErrorMsg)

End Sub

Public Sub HandleHTError(ByVal httpaSession As CVtHttpaSession)

     
     Const PROC_NAME As String = "HandleHTError"
     Dim CustomData As CMyCustomData
     Set CustomData = httpaSession.CustomData
     
     Call VtLogInfo(MODULE_NAME, PROC_NAME, httpaSession.ErrorMsg)
     
    ' Call UITabLog.WriteLog("    ERROR  : " & httpaSession.ErrorMsg)

End Sub

Public Sub HandleTimeout(ByVal httpaSession As CVtHttpaSession)

     Const PROC_NAME As String = "HandleTimeout"
     Dim CustomData As CMyCustomData
     Set CustomData = httpaSession.CustomData
     
     
     Call VtLogInfo(MODULE_NAME, PROC_NAME, httpaSession.TimeoutReason)

      'Call UITabLog.WriteLog("    TIMEOUT  : " & httpaSession.TimeoutReason)

End Sub


