Attribute VB_Name = "WasHttpaWebSocket"
'Attribute VB_Name = "WasHttpaWebSocket"
Option Explicit

Private Const MODULE_NAME As String = "WasHttpaWebSocket"

Public Sub Init()
   ' WasUIWebSocketSessions.Init
End Sub


Public Sub Uninit()

End Sub


Public Sub HandleEstablised(ByVal httpaSession As CVtHttpaSession)
  
    Const PROC_NAME As String = "HandleEstablised"
    
    Dim errorOut As String
    Dim result As Boolean


    VtLogInfo MODULE_NAME, PROC_NAME

    result = WasUIWebSocketSessions.OnOpen( _
        httpaSession.websocket, _
        errorOut)

    If Not result Then
       VtLogError MODULE_NAME, PROC_NAME, _
           "WasUIWebSocketSessions.OnOpen() failed->" & errorOut
       httpaSession.websocket.CloseSession

    Else
       VtLogInfo MODULE_NAME, PROC_NAME, _
           "WasUIWebSocketSessions.OnOpen() success"

    End If

        
    
   
  
End Sub

Public Sub HandleReceivedText(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleReceivedText"
    Dim msg As String
    msg = httpaSession.websocket.LastText
        
    
    VtLogInfo MODULE_NAME, PROC_NAME, msg
    
    Call WasUIWebSocketSessions.OnRecv( _
        httpaSession.websocket, _
        msg)
     

End Sub

Public Sub HandleReceivedBinary(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleReceivedBinary"

    ' Binary WebSocket messages are not used by the current MVP.
End Sub

Public Sub HandleClosed(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleClosed"
    
    VtLogInfo MODULE_NAME, PROC_NAME
    
    Call WasUIWebSocketSessions.OnClose(httpaSession.websocket)
    
    
    

End Sub





