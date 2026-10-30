Attribute VB_Name = "WasHttpaWebSocket"
'Attribute VB_Name = "WasHttpaWebSocket"
Option Explicit

Private Const MODULE_NAME As String = "WasHttpaWebSocket"

Public Sub Init()
    WebsocketDemo.WebsocketDemo_Init
End Sub


Public Sub Uninit()
    WebsocketDemo.WebsocketDemo_Uninit

End Sub


Public Sub HandleEstablised(ByVal httpaSession As CVtHttpaSession)
  
    Const PROC_NAME As String = "HandleEstablised"
    
    Dim errorOut As String
    Dim result As Boolean
    
    
    VtLogInfo MODULE_NAME, PROC_NAME
    
    result = WebsocketDemo.WebsocketDemo_OnOpen( _
        httpaSession.Websocket, _
        errorOut)
    
    If Not result Then
       VtLogError MODULE_NAME, PROC_NAME, _
           "WebsocketDemo.WebsocketDemo_OnOpen() failed->" & errorOut
       httpaSession.Websocket.CloseSession
        
    Else
       VtLogInfo MODULE_NAME, PROC_NAME, _
           "WebsocketDemo.WebsocketDemo_OnOpen() success"
    
    End If
    
        
    
   
  
End Sub

Public Sub HandleReceivedText(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleReceivedText"
    Dim msg As String
    msg = httpaSession.Websocket.LastText
        
    
    VtLogInfo MODULE_NAME, PROC_NAME, msg
    
    Call WebsocketDemo.WebsocketDemo_OnRecv( _
        httpaSession.Websocket, _
        msg)
     

End Sub

Public Sub HandleReceivedBinary(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleReceivedBinary"

    ' Binary WebSocket messages are not used by the current MVP.
End Sub

Public Sub HandleClosed(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleClosed"
    
    VtLogInfo MODULE_NAME, PROC_NAME
    
    Call WebsocketDemo.WebsocketDemo_OnClose(httpaSession.Websocket)
    
    
    

End Sub





