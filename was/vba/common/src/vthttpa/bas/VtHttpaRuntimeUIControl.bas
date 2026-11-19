Attribute VB_Name = "VtHttpaRuntimeUIControl"
Option Explicit

' Public UI API for VtHttpaRuntimeConsole.
' Showing or hiding this form never starts or stops the VTNode Runtime.

Private Const MODULE_NAME As String = "VtHttpaRuntimeUIControl"
Private Const FORM_TYPE_NAME As String = "VtHttpaRuntimeConsole"

Public Sub ShowVtHttpaRuntimeUI()
    On Error GoTo EH

    If IsVtHttpaRuntimeUIVisible Then Exit Sub

    VtHttpaRuntimeConsole.Show vbModeless
    Exit Sub

EH:
    Err.Raise Err.number, _
              MODULE_NAME & ".ShowVtHttpaRuntimeUI", _
              Err.description
End Sub

'Public Sub ShowDemoUI()
'    On Error GoTo EH
'
'    With VtHttpaRuntimeConsole
'        .SetRuntimeStatus "RUNNING"
'        .SetRuntimeOwner "get-started.xlsm", _
'                         "C:\VTNode\examples\get-started\get-started.xlsm"
'        .SetNetworkAddresses "192.168.0.4", _
'                             "211.123.124.2"
'        .SetHttpServer 12345, _
'                       "http://192.168.0.4:12345", _
'                       "ws://192.168.0.4:12345"
'        .SetHttpsServer True, _
'                        8080, _
'                        "https://192.168.0.4:8080", _
'                        "wss://192.168.0.4:8080"
'        .SetRuntimeTiming "2026-08-28 09:32:10", _
'                          "00:12:35"
'        .SetStats 127, 2
'        .SetLatestActivity _
'            "15:42:35  200 OK" & vbCrLf _
'            & "15:42:35  GET /favicon.ico" & vbCrLf _
'            & "15:42:31  200 OK" & vbCrLf _
'            & "15:42:31  GET /index.html"
'        .Show vbModeless
'    End With
'
'    Exit Sub
'
'EH:
'    Err.Raise Err.Number, _
'              MODULE_NAME & ".ShowDemoUI", _
'              Err.Description
'End Sub

Public Sub HideVtHttpaRuntimeUI()
    Dim runtimeForm As Object

    On Error GoTo EH

    Set runtimeForm = GetLoadedVtHttpaRuntimeForm()
    If runtimeForm Is Nothing Then Exit Sub

    runtimeForm.Hide
    Exit Sub

EH:
    Err.Raise Err.number, _
              MODULE_NAME & ".HideVtHttpaRuntimeUI", _
              Err.description
End Sub

Public Sub ToggleVtHttpaRuntimeUI()
    If IsVtHttpaRuntimeUIVisible Then
        HideVtHttpaRuntimeUI
    Else
        ShowVtHttpaRuntimeUI
    End If
End Sub

Public Sub UpdateRuntimeStatus(ByVal value As VtHttpaRuntimeStatus)
    VtHttpaRuntimeConsole.SetRuntimeStatus VtHttpaRuntimeStatusToString(value)
End Sub

Public Sub UpdateAppInfo(ByRef value As VtHttpaRuntimeAppInfo)
    VtHttpaRuntimeConsole.SetRuntimeOwner _
        value.ApplicationName, _
        value.WorkbookPath
End Sub

Public Sub UpdateNetworkInfo(ByRef value As VtHttpaRuntimeNetworkInfo)
    VtHttpaRuntimeConsole.SetNetworkAddresses _
        value.LocalIp, _
        value.ExternalIp
End Sub

Public Sub UpdateHttpInfo(ByRef value As VtHttpaRuntimeHttpInfo)
    If value.Enabled Then
        VtHttpaRuntimeConsole.SetHttpServer _
            value.Port, _
            value.HttpUrl, _
            value.WebSocketUrl
    Else
        VtHttpaRuntimeConsole.SetHttpServer "-", "-", "-"
    End If
End Sub

Public Sub UpdateHttpsInfo(ByRef value As VtHttpaRuntimeHttpsInfo)
    VtHttpaRuntimeConsole.SetHttpsServer _
        value.Enabled, _
        value.Port, _
        value.HttpsUrl, _
        value.SecureWebSocketUrl
End Sub

Public Sub UpdateTimingInfo(ByRef value As VtHttpaRuntimeTimingInfo)
    Dim startedAtText As String

    If value.StartedAt = 0 Then
        startedAtText = "-"
    Else
        startedAtText = Format$(value.StartedAt, "yyyy-mm-dd hh:nn:ss")
    End If

    VtHttpaRuntimeConsole.SetRuntimeTiming _
        startedAtText, _
        FormatUptime(value.UptimeSeconds)
End Sub

Public Sub UpdateStatsInfo(ByRef value As VtHttpaRuntimeStatsInfo)
    VtHttpaRuntimeConsole.SetStats _
        value.TotalRequests, _
        value.WebSocketClients
End Sub

Public Sub UpdateLatestActivity(ByVal activityText As String)
    VtHttpaRuntimeConsole.SetLatestActivity activityText
End Sub

Public Sub AddLatestActivity(ByVal activityText As String)
    VtHttpaRuntimeConsole.AddLatestActivity activityText
End Sub

Public Sub UpdateRuntimeState(ByRef value As VtNodeRuntimeState)
    UpdateRuntimeStatus value.Status
    UpdateAppInfo value.App
    UpdateNetworkInfo value.Network
    UpdateHttpInfo value.Http
    UpdateHttpsInfo value.Https
    UpdateTimingInfo value.Timing
    UpdateStatsInfo value.Stats
    UpdateLatestActivity value.LatestActivity
End Sub

Public Property Get IsVtHttpaRuntimeUIVisible() As Boolean
    Dim runtimeForm As Object

    Set runtimeForm = GetLoadedVtHttpaRuntimeForm()

    If runtimeForm Is Nothing Then
        IsVtHttpaRuntimeUIVisible = False
    Else
        IsVtHttpaRuntimeUIVisible = runtimeForm.Visible
    End If
End Property

Public Property Get IsVtHttpaRuntimeUILoaded() As Boolean
    Dim runtimeForm As Object

    Set runtimeForm = GetLoadedVtHttpaRuntimeForm()
    IsVtHttpaRuntimeUILoaded = Not runtimeForm Is Nothing
End Property

Private Function GetLoadedVtHttpaRuntimeForm() As Object
    Dim candidate As Object

    For Each candidate In VBA.UserForms
        If StrComp(TypeName(candidate), FORM_TYPE_NAME, vbTextCompare) = 0 Then
            Set GetLoadedVtHttpaRuntimeForm = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Function FormatUptime(ByVal totalSeconds As Long) As String
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long

    If totalSeconds < 0 Then totalSeconds = 0

    hours = totalSeconds \ 3600
    minutes = (totalSeconds Mod 3600) \ 60
    seconds = totalSeconds Mod 60

    FormatUptime = Pad2(hours) _
                   & ":" & Pad2(minutes) _
                   & ":" & Pad2(seconds)
End Function

Private Function Pad2(ByVal value As Long) As String
    If value < 10 Then
        Pad2 = "0" & CStr(value)
    Else
        Pad2 = CStr(value)
    End If
End Function
