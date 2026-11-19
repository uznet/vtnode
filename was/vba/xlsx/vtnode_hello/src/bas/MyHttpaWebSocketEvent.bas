Attribute VB_Name = "MyHttpaWebSocketEvent"
'Attribute VB_Name = "WasHttpaWebSocket"
Option Explicit

Private Const MODULE_NAME As String = "MyHttpaWebSocketEvent"

Private mStartedAt As LongLong

Public Sub Init()

    'Set mFormVTNode = frmVTNode
    mStartedAt = VtTimeUtils.VtGetCurrentCTime()

End Sub


Public Sub Uninit()

    mStartedAt = 0

End Sub


Public Sub HandleEstablised(ByVal httpaSession As CVtHttpaSession)
  
    Const PROC_NAME As String = "HandleEstablised"
    
    Dim errorOut As String
    Dim result As Boolean


    VtLogInfo MODULE_NAME, PROC_NAME

    result = WasUIWebSocketSessions.OnOpen( _
        httpaSession.Websocket, _
        errorOut)

    If Not result Then
       VtLogError MODULE_NAME, PROC_NAME, _
           "WasUIWebSocketSessions.OnOpen() failed->" & errorOut
       httpaSession.Websocket.CloseSession

    Else
       VtLogInfo MODULE_NAME, PROC_NAME, _
           "WasUIWebSocketSessions.OnOpen() success"

    End If

        
    
   
  
End Sub

Public Sub HandleReceivedText(ByVal httpaSession As CVtHttpaSession)
    Const PROC_NAME As String = "HandleReceivedText"
    Dim msg As String
    msg = httpaSession.Websocket.LastText
        
    
    VtLogInfo MODULE_NAME, PROC_NAME, msg
    
    Call WasUIWebSocketSessions.OnRecv( _
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
    
    Call WasUIWebSocketSessions.OnClose(httpaSession.Websocket)
    
    
    

End Sub





