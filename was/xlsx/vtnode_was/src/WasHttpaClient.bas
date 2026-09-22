Attribute VB_Name = "WasHttpaClient"
Option Explicit


Private Const MODULE_NAME As String = "WasHttpaClient"
'Private mFormVTNode As frmVTNode

Public Sub Init()

   'Set mFormVTNode = frmVTNode
   
End Sub


Public Sub Uninit()

End Sub


Public Sub HandleConnecting(ByVal httpaSession As CVtHttpaSession)


End Sub

Public Sub HandleConnected(ByVal httpaSession As CVtHttpaSession)


End Sub



Public Sub HandleConnectFailed(ByVal httpaSession As CVtHttpaSession)


End Sub

Public Sub HandleDisconnected(ByVal httpaSession As CVtHttpaSession)


End Sub


Public Sub HandleTransactionBegin(ByVal httpaSession As CVtHttpaSession)

End Sub

Public Sub HandleTransactionEnd(ByVal httpaSession As CVtHttpaSession)

End Sub


Public Sub HandleResponse(ByVal httpaSession As CVtHttpaSession)

    Const PROC_NAME As String = "HandleResponse"
    Dim summary As String
    
    summary = httpaSession.response.ToSummary()
                
    Call VtLogInfo(MODULE_NAME, PROC_NAME, "----------------------------------------------------------")
    Call VtLogInfo(MODULE_NAME, PROC_NAME, ">>> HTTP RESPONSE:" & summary)
                
    summary = httpaSession.response.Contents.ToSummary()
    Call VtLogInfo(MODULE_NAME, PROC_NAME, ">>> RESPONSE CONTENTS:" & summary)

End Sub







