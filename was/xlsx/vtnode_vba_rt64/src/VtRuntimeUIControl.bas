Attribute VB_Name = "VtRuntimeUIControl"
Option Explicit

' Public UI API for VtRuntimeConsole.
' Showing or hiding this form never starts or stops the VTNode Runtime.

Private Const MODULE_NAME As String = "VtRuntimeUIControl"
Private Const FORM_TYPE_NAME As String = "VtRuntimeConsole"

Public Sub ShowVtRuntimeUI()
    On Error GoTo EH

    If IsVtRuntimeUIVisible Then Exit Sub

    VtRuntimeConsole.Show vbModeless
    Exit Sub

EH:
    Err.Raise Err.Number, _
              MODULE_NAME & ".ShowVtRuntimeUI", _
              Err.Description
End Sub

Public Sub ShowDemoUI()
    On Error GoTo EH

    With VtRuntimeConsole
        .SetRuntimeStatus "RUNNING"
        .SetRuntimeOwner "get-started.xlsm", _
                         "C:\VTNode\examples\get-started\get-started.xlsm"
        .SetNetworkAddresses "192.168.0.4", _
                             "211.123.124.2"
        .SetHttpServer 12345, _
                       "http://192.168.0.4:12345", _
                       "ws://192.168.0.4:12345"
        .SetHttpsServer True, _
                        8080, _
                        "https://192.168.0.4:8080", _
                        "wss://192.168.0.4:8080"
        .SetRuntimeTiming "2026-08-28 09:32:10", _
                          "00:12:35"
        .SetStats 127, 2
        .SetLatestActivity _
            "15:42:35  200 OK" & vbCrLf _
            & "15:42:35  GET /favicon.ico" & vbCrLf _
            & "15:42:31  200 OK" & vbCrLf _
            & "15:42:31  GET /index.html"
        .Show vbModeless
    End With

    Exit Sub

EH:
    Err.Raise Err.Number, _
              MODULE_NAME & ".ShowDemoUI", _
              Err.Description
End Sub

Public Sub HideVtRuntimeUI()
    Dim runtimeForm As Object

    On Error GoTo EH

    Set runtimeForm = GetLoadedVtRuntimeForm()
    If runtimeForm Is Nothing Then Exit Sub

    runtimeForm.Hide
    Exit Sub

EH:
    Err.Raise Err.Number, _
              MODULE_NAME & ".HideVtRuntimeUI", _
              Err.Description
End Sub

Public Sub ToggleVtRuntimeUI()
    If IsVtRuntimeUIVisible Then
        HideVtRuntimeUI
    Else
        ShowVtRuntimeUI
    End If
End Sub

Public Sub UpdateRuntimeStatus(ByVal value As VtRuntimeStatus)
    VtRuntimeConsole.SetRuntimeStatus VtRuntimeStatusToString(value)
End Sub

Public Sub UpdateAppInfo(ByRef value As VtRuntimeAppInfo)
    VtRuntimeConsole.SetRuntimeOwner _
        value.ApplicationName, _
        value.WorkbookPath
End Sub

Public Sub UpdateNetworkInfo(ByRef value As VtRuntimeNetworkInfo)
    VtRuntimeConsole.SetNetworkAddresses _
        value.LocalIp, _
        value.ExternalIp
End Sub

Public Sub UpdateHttpInfo(ByRef value As VtRuntimeHttpInfo)
    If value.Enabled Then
        VtRuntimeConsole.SetHttpServer _
            value.Port, _
            value.HttpUrl, _
            value.WebSocketUrl
    Else
        VtRuntimeConsole.SetHttpServer "-", "-", "-"
    End If
End Sub

Public Sub UpdateHttpsInfo(ByRef value As VtRuntimeHttpsInfo)
    VtRuntimeConsole.SetHttpsServer _
        value.Enabled, _
        value.Port, _
        value.HttpsUrl, _
        value.SecureWebSocketUrl
End Sub

Public Sub UpdateTimingInfo(ByRef value As VtRuntimeTimingInfo)
    Dim startedAtText As String

    If value.StartedAt = 0 Then
        startedAtText = "-"
    Else
        startedAtText = Format$(value.StartedAt, "yyyy-mm-dd hh:nn:ss")
    End If

    VtRuntimeConsole.SetRuntimeTiming _
        startedAtText, _
        FormatUptime(value.UptimeSeconds)
End Sub

Public Sub UpdateStatsInfo(ByRef value As VtRuntimeStatsInfo)
    VtRuntimeConsole.SetStats _
        value.TotalRequests, _
        value.WebSocketClients
End Sub

Public Sub UpdateLatestActivity(ByVal activityText As String)
    VtRuntimeConsole.SetLatestActivity activityText
End Sub

Public Sub AddLatestActivity(ByVal activityText As String)
    VtRuntimeConsole.AddLatestActivity activityText
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

Public Property Get IsVtRuntimeUIVisible() As Boolean
    Dim runtimeForm As Object

    Set runtimeForm = GetLoadedVtRuntimeForm()

    If runtimeForm Is Nothing Then
        IsVtRuntimeUIVisible = False
    Else
        IsVtRuntimeUIVisible = runtimeForm.Visible
    End If
End Property

Public Property Get IsVtRuntimeUILoaded() As Boolean
    Dim runtimeForm As Object

    Set runtimeForm = GetLoadedVtRuntimeForm()
    IsVtRuntimeUILoaded = Not runtimeForm Is Nothing
End Property

Private Function GetLoadedVtRuntimeForm() As Object
    Dim candidate As Object

    For Each candidate In VBA.UserForms
        If StrComp(TypeName(candidate), FORM_TYPE_NAME, vbTextCompare) = 0 Then
            Set GetLoadedVtRuntimeForm = candidate
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
