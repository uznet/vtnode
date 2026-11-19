Attribute VB_Name = "WasHttpaRuntimeStates"
Option Explicit

Private Const MODULE_NAME As String = "WasHttpaRuntimeStates"


Private Sub CollectAppInfo(ByRef info As VtHttpaRuntimeAppInfo)
    info.ApplicationName = ThisWorkbook.Name
    info.WorkbookPath = ThisWorkbook.FullName
End Sub

Private Sub CollectNetworkInfo(ByRef info As VtHttpaRuntimeNetworkInfo)


    Dim eout As String
    Dim r As Boolean

    r = VtNetwork.VtNetGetLocalIp(info.LocalIp, eout)
    r = VtNetwork.VtNetGetExternalIp(info.ExternalIp, eout)


End Sub
Private Function CollectRuntimeStatus() As VtHttpaRuntimeStatus

    '
    ' No service instance
    '
    
    If Not VtHttpaService.HasInstance Then
        CollectRuntimeStatus = VT_RUNTIME_STOPPED
        Exit Function
    End If

    '
    ' Service instance exists and is running
    '
    If VtHttpaService.HttpaIsRunning Then
        CollectRuntimeStatus = VT_RUNTIME_RUNNING
    Else
        CollectRuntimeStatus = VT_RUNTIME_STOPPED
    End If

End Function
'Public Type VtNodeRuntimeState
'    Status As VtHttpaRuntimeStatus
'    App As VtHttpaRuntimeAppInfo
'    network As VtHttpaRuntimeNetworkInfo
'    Http As VtHttpaRuntimeHttpInfo
'    Https As VtHttpaRuntimeHttpsInfo
'    Timing As VtHttpaRuntimeTimingInfo
'    Stats As VtHttpaRuntimeStatsInfo
'    LatestActivity As String
'    LastError As String
'End Type
Private Function CollectRuntimeState() As VtNodeRuntimeState

    Dim state As VtNodeRuntimeState

    state.Status = CollectRuntimeStatus()

    Call CollectAppInfo(state.App)
    Call CollectNetworkInfo(state.Network)

    If VtHttpaService.HasInstance Then

        state.Http = VtHttpaService.GetHttpInfo(state.Network)

        state.Https = VtHttpaService.GetHttpsInfo(state.Network)

        state.Timing = VtHttpaService.GetTimingInfo()

        state.Stats = VtHttpaService.GetStatsInfo()

        state.LatestActivity = VtHttpaService.GetLatestActivity()

        state.LastError = VtHttpaService.GetLastError()

    End If

    CollectRuntimeState = state

End Function

Public Sub ShowRuntimeStates()

    Dim state As VtNodeRuntimeState

    On Error GoTo EH

    '
    ' Collect the current runtime state.
    '

    state = CollectRuntimeState()

    '
    ' Update the GET STARTED worksheet.
    '

    ShowRuntimeStatesOnSheet state


    '
    ' Update the console before showing it.
    '
   VtHttpaRuntimeUIControl.UpdateRuntimeState state

    '
    ' Show the Runtime Console modelessly.
    '
    VtHttpaRuntimeUIControl.ShowVtHttpaRuntimeUI

    Exit Sub

EH:
    Err.Raise Err.number, "ShowRuntimeStates", Err.description

End Sub


Public Sub UpdateRuntimeStates()

    Dim state As VtNodeRuntimeState

    On Error GoTo EH

    '
    ' Collect once.
    '
    state = CollectRuntimeState()

    '
    ' Always update the worksheet.
    '
    
    ShowRuntimeStatesOnSheet state

    '
    ' Update the Runtime Console only when it is visible.
    '

    If VtHttpaRuntimeUIControl.IsVtHttpaRuntimeUIVisible Then
        VtHttpaRuntimeUIControl.UpdateRuntimeState state
    End If

    Exit Sub

EH:
    Debug.Print "[UpdateRuntimeStates] " _
              & Err.number & " : " _
              & Err.description

End Sub


Private Sub UpdateHyperLink(ByVal target As Range, ByVal url As String)

    ' 기존 하이퍼링크 제거
    target.Hyperlinks.Delete

    ' 새 주소와 표시 문자열을 함께 등록
    target.Worksheet.Hyperlinks.Add _
        Anchor:=target, _
        Address:=url, _
        TextToDisplay:=url

End Sub

Public Sub SetAccessUrl(ByVal url As String)

   Call UpdateHyperLink(ThisWorkbook.Worksheets(MAIN_SHEET).Range(ACCESS_URL_CELL), url)
'    ThisWorkbook.Worksheets(MAIN_SHEET) _
'        .Range(ACCESS_URL_CELL).Value2 = Trim$(url)
End Sub


Public Sub SetNetwork( _
    ByRef httpInfo As VtHttpaRuntimeHttpInfo, _
    ByRef httpsInfo As VtHttpaRuntimeHttpsInfo)

    Dim httpText As String
    Dim httpsText As String

    If httpInfo.Enabled Then
        httpText = "HTTP " & CStr(httpInfo.Port)
    Else
        httpText = "HTTP OFF"
    End If

    If httpsInfo.Enabled Then
        httpsText = "HTTPS " & CStr(httpsInfo.Port)
    Else
        httpsText = "HTTPS OFF"
    End If

    ThisWorkbook.Worksheets(MAIN_SHEET) _
        .Range(WasUIController.NETWORK_RANGE).Cells(1, 1).Value2 = _
        httpText & "  ·  " & httpsText
End Sub

Private Sub ShowRuntimeStatesOnSheet(ByRef state As VtNodeRuntimeState)

    Const SHEET_NAME As String = WasUIController.MAIN_SHEET

    Dim ws As Worksheet
    Dim statusText As String
    Dim serviceUrl As String
    Dim connectionCount As Long

    On Error GoTo EH

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

    ' Update only the three runtime fields exposed by the VTNode WAS MVP shell.

    If state.Http.Enabled And Len(Trim$(state.Http.HttpUrl)) > 0 Then
        serviceUrl = Trim$(state.Http.HttpUrl)
    ElseIf state.Https.Enabled And Len(Trim$(state.Https.HttpsUrl)) > 0 Then
        serviceUrl = Trim$(state.Https.HttpsUrl)
    Else
        serviceUrl = "-"
    End If

    UpdateHyperLink ws.Range(WasUIController.ACCESS_URL_CELL), serviceUrl


    statusText = VtHttpaRuntimeStatusToString(state.Status)
    With ws.Range("F11:H12")
        .Cells(1, 1).Value2 = statusText

        Select Case state.Status
            Case VT_RUNTIME_RUNNING
                .Font.Color = RGB(21, 128, 61)
            Case VT_RUNTIME_STARTING, VT_RUNTIME_STOPPING
                .Font.Color = RGB(180, 83, 9)
            Case VT_RUNTIME_FAULTED
                .Font.Color = RGB(220, 38, 38)
            Case Else
                .Font.Color = RGB(194, 65, 12)
        End Select
    End With

    '
    ' Synchronize the QR image with the new D27 value.
    '
    SetQRCode SHEET_NAME, WasUIController.QRCODE_CELL, IsRunning:=state.Http.Enabled, serviceUrl:=serviceUrl
    
    SetNetwork state.Http, state.Https
    

    Exit Sub


EH:
    Err.Raise Err.number, "ShowRuntimeStatesOnSheet", Err.description

End Sub


'

Public Function WebsocketDemoAccessUrl() As String
    Dim state As VtNodeRuntimeState
    Dim baseUrl As String
    Dim hostStart As Long
    Dim pathStart As Long

    state = CollectRuntimeState()
    WebsocketDemo.WebsocketDemo_UpdateHeaderState (state.Status = VT_RUNTIME_RUNNING)
    If state.Http.Enabled Then baseUrl = Trim$(state.Http.HttpUrl)
    If Len(baseUrl) = 0 And state.Https.Enabled Then baseUrl = Trim$(state.Https.HttpsUrl)
    If Len(baseUrl) = 0 Or baseUrl = "-" Then Exit Function
    If LCase$(left$(baseUrl, 7)) <> "http://" And _
       LCase$(left$(baseUrl, 8)) <> "https://" Then Exit Function
    hostStart = InStr(1, baseUrl, "://", vbBinaryCompare) + 3
    pathStart = InStr(hostStart, baseUrl, "/", vbBinaryCompare)
    If pathStart > 0 Then baseUrl = left$(baseUrl, pathStart - 1)
    WebsocketDemoAccessUrl = baseUrl & "/websocket.html"
End Function


