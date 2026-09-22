Attribute VB_Name = "WasHttpaError"
Option Explicit

Private Const MODULE_NAME As String = "WasHttpaError"
'Private mFormVTNode As frmVTNode

Public Sub Init()

   'Set mFormVTNode = frmVTNode
   
End Sub


Public Sub Uninit()

End Sub
Public Sub HandleError(ByVal httpaSession As CVtHttpaSession)
     Const PROC_NAME As String = "HandleError"
     
     Call VtLogInfo(MODULE_NAME, PROC_NAME, httpaSession.ErrorMsg)
     
     'Call UITabLog.WriteLog("    ERROR  : " & httpaSession.ErrorMsg)
End Sub

Public Sub HandleNetError(ByVal httpaSession As CVtHttpaSession)

     Const PROC_NAME As String = "HandleNetError"
     
     Dim CustomData As CWasCustomData
     Set CustomData = httpaSession.CustomData
     

     Call VtLogInfo(MODULE_NAME, PROC_NAME, httpaSession.ErrorMsg)

     'Call UITabLog.WriteLog("    ERROR  : " & httpaSession.ErrorMsg)

End Sub

Public Sub HandleHTError(ByVal httpaSession As CVtHttpaSession)

     
     Const PROC_NAME As String = "HandleHTError"
     Dim CustomData As CWasCustomData
     Set CustomData = httpaSession.CustomData
     
     Call VtLogInfo(MODULE_NAME, PROC_NAME, httpaSession.ErrorMsg)
     
    ' Call UITabLog.WriteLog("    ERROR  : " & httpaSession.ErrorMsg)

End Sub

Public Sub HandleTimeout(ByVal httpaSession As CVtHttpaSession)

     Const PROC_NAME As String = "HandleTimeout"
     Dim CustomData As CWasCustomData
     Set CustomData = httpaSession.CustomData
     
     
     Call VtLogInfo(MODULE_NAME, PROC_NAME, httpaSession.TimeoutReason)

      'Call UITabLog.WriteLog("    TIMEOUT  : " & httpaSession.TimeoutReason)

End Sub


