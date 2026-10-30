VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VtRuntimeConsole 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "VtRuntimeConsole.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "VtRuntimeConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Version 5#
'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VtRuntimeConsole
'   caption = "VTNode Runtime"
'   ClientHeight = 10200
'   ClientLeft = 120
'   ClientTop = 465
'   ClientWidth = 5850
'   StartUpPosition = 0    'Manual
'End
'Attribute VB_Name = "VtRuntimeConsole"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
Option Explicit

' VTNode Runtime dashboard
' The application workbook owns START/STOP. This form only presents runtime state.

Private Const FORM_WIDTH As Single = 390
Private Const MARGIN_X As Single = 18
Private Const VALUE_X As Single = 132

Private WithEvents mHideButton As MSForms.CommandButton
Attribute mHideButton.VB_VarHelpID = -1
'Attribute mHideButton.VB_VarHelpID = -1

Private mStatusDot As MSForms.label
Private mStatusValue As MSForms.label
Private mApplicationValue As MSForms.label
Private mWorkbookValue As MSForms.label
Private mLocalIpValue As MSForms.label
Private mExternalIpValue As MSForms.label
Private mHttpPortValue As MSForms.label
Private mHttpUrlValue As MSForms.label
Private mWebSocketUrlValue As MSForms.label
Private mTlsStatusValue As MSForms.label
Private mTlsPortValue As MSForms.label
Private mHttpsUrlValue As MSForms.label
Private mSecureWebSocketUrlValue As MSForms.label
Private mStartedAtValue As MSForms.label
Private mWebSocketClientsValue As MSForms.label
Private mUptimeValue As MSForms.label
Private mRequestsValue As MSForms.label
Private mActivityValue As MSForms.label

Private Sub UserForm_Initialize()
    BuildDashboard
    ResetDashboard
End Sub

Private Sub UserForm_Activate()
    PositionAtExcelRight
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub mHideButton_Click()
    Me.Hide
End Sub

Public Sub ResetDashboard()
    SetRuntimeStatus "STOPPED"
    SetRuntimeOwner "-", "-"
    SetNetworkAddresses "-", "-"
    SetHttpServer "-", "-", "-"
    SetHttpsServer False, "-", "-", "-"
    SetRuntimeTiming "-", "00:00:00"
    SetStats 0, 0
    SetLatestActivity "Waiting for runtime activity..."
End Sub

Public Sub SetRuntimeOwner( _
    ByVal ApplicationName As String, _
    ByVal WorkbookPath As String)

    mApplicationValue.caption = DisplayValue(ApplicationName)
    mApplicationValue.ControlTipText = CStr(ApplicationName)
    mWorkbookValue.caption = DisplayValue(WorkbookPath)
    mWorkbookValue.ControlTipText = CStr(WorkbookPath)
End Sub

Public Sub SetRuntimeStatus(ByVal statusText As String)
    Dim normalized As String

    normalized = UCase$(Trim$(statusText))
    If Len(normalized) = 0 Then normalized = "UNKNOWN"

    mStatusValue.caption = normalized

    Select Case normalized
        Case "RUNNING"
            mStatusDot.foreColor = RGB(52, 211, 153)
            mStatusValue.foreColor = RGB(52, 211, 153)
        Case "STARTING", "STOPPING", "WARNING"
            mStatusDot.foreColor = RGB(245, 158, 11)
            mStatusValue.foreColor = RGB(245, 158, 11)
        Case "FAULTED", "ERROR"
            mStatusDot.foreColor = RGB(239, 68, 68)
            mStatusValue.foreColor = RGB(239, 68, 68)
        Case Else
            mStatusDot.foreColor = RGB(107, 114, 128)
            mStatusValue.foreColor = RGB(156, 163, 175)
    End Select
End Sub

Public Sub SetHttpServer( _
    ByVal Port As Variant, _
    ByVal HttpUrl As String, _
    ByVal WebSocketUrl As String)

    mHttpPortValue.caption = DisplayValue(Port)
    SetLabelValue mHttpUrlValue, HttpUrl
    SetLabelValue mWebSocketUrlValue, WebSocketUrl
End Sub

Public Sub SetNetworkAddresses( _
    ByVal LocalIp As String, _
    ByVal ExternalIp As String)

    SetLabelValue mLocalIpValue, LocalIp
    SetLabelValue mExternalIpValue, ExternalIp
End Sub

Public Sub SetHttpsServer( _
    ByVal tlsEnabled As Boolean, _
    ByVal tlsPort As Variant, _
    ByVal HttpsUrl As String, _
    ByVal SecureWebSocketUrl As String)

    SetEnabledValue mTlsStatusValue, tlsEnabled
    mTlsPortValue.caption = IIf(tlsEnabled, DisplayValue(tlsPort), "-")

    If tlsEnabled Then
        SetLabelValue mHttpsUrlValue, HttpsUrl
        SetLabelValue mSecureWebSocketUrlValue, SecureWebSocketUrl
    Else
        SetLabelValue mHttpsUrlValue, "-"
        SetLabelValue mSecureWebSocketUrlValue, "-"
    End If
End Sub

Public Sub SetRuntimeTiming( _
    ByVal StartedAt As String, _
    ByVal uptime As String)

    mStartedAtValue.caption = DisplayValue(StartedAt)
    mUptimeValue.caption = DisplayValue(uptime)
End Sub

Public Sub SetStats( _
    ByVal TotalRequests As Long, _
    ByVal WebSocketClients As Long)

    mRequestsValue.caption = Format$(TotalRequests, "#,##0")
    mWebSocketClientsValue.caption = Format$(WebSocketClients, "#,##0")
End Sub

Public Sub SetLatestActivity(ByVal activityText As String)
    If Len(Trim$(activityText)) = 0 Then
        mActivityValue.caption = "Waiting for runtime activity..."
    Else
        mActivityValue.caption = activityText
    End If
End Sub

Public Sub AddLatestActivity(ByVal activityText As String)
    Dim currentText As String

    currentText = mActivityValue.caption
    If currentText = "Waiting for runtime activity..." Then currentText = vbNullString

    If Len(currentText) > 0 Then
        mActivityValue.caption = activityText & vbCrLf & KeepFirstLines(currentText, 3)
    Else
        mActivityValue.caption = activityText
    End If
End Sub

Private Sub BuildDashboard()
    Dim y As Single

    Me.caption = "VTNode Runtime"
    Me.width = FORM_WIDTH
    Me.height = 710
    Me.BackColor = RGB(24, 26, 29)
    Me.BorderStyle = fmBorderStyleSingle
    Me.ScrollBars = fmScrollBarsNone
    Me.KeepScrollBarsVisible = fmScrollBarsNone

    AddText "lblTitle", "VTNode Runtime", MARGIN_X, 14, 210, 24, 16, RGB(232, 232, 232), True
    AddText "lblSubtitle", "Excel Web Application Runtime", MARGIN_X, 39, 250, 17, 9, RGB(156, 163, 170), False

    Set mStatusDot = AddText("lblStatusDot", ChrW$(&H25CF), 298, 17, 14, 18, 12, RGB(107, 114, 128), True)
    Set mStatusValue = AddText("lblStatusValue", "STOPPED", 314, 18, 58, 17, 9, RGB(156, 163, 175), True)

    y = 66
    AddSection "RUNTIME OWNER", y
    Set mApplicationValue = AddRow("Application", "-", y + 30)
    Set mWorkbookValue = AddRow("Workbook", "-", y + 52)

    y = 136
    AddSection "NETWORK", y
    Set mLocalIpValue = AddRow("Local IP", "-", y + 30)
    Set mExternalIpValue = AddRow("External IP", "-", y + 52)

    y = 206
    AddSection "HTTP SERVER", y
    Set mHttpPortValue = AddRow("Port", "-", y + 30)
    Set mHttpUrlValue = AddRow("HTTP URL", "-", y + 52)
    Set mWebSocketUrlValue = AddRow("WebSocket URL", "-", y + 74)

    y = 298
    AddSection "HTTPS SERVER", y
    Set mTlsStatusValue = AddRow("TLS", "Disabled", y + 30)
    Set mTlsPortValue = AddRow("Port", "-", y + 52)
    Set mHttpsUrlValue = AddRow("HTTPS URL", "-", y + 74)
    Set mSecureWebSocketUrlValue = AddRow("Secure WS URL", "-", y + 96)

    y = 412
    AddSection "RUNTIME", y
    Set mStartedAtValue = AddRow("Started At", "-", y + 30)
    Set mUptimeValue = AddRow("Uptime", "00:00:00", y + 52)

    y = 482
    AddSection "STATS", y
    Set mRequestsValue = AddRow("Total Requests", "0", y + 30)
    Set mWebSocketClientsValue = AddRow("WS Clients", "0", y + 52)

    y = 552
    AddSection "LATEST ACTIVITY", y
    Set mActivityValue = AddText("lblActivityValue", "", MARGIN_X, y + 30, 354, 58, 9, RGB(209, 213, 219), False)
    mActivityValue.BackColor = RGB(34, 37, 42)
    mActivityValue.BackStyle = fmBackStyleOpaque
    mActivityValue.WordWrap = True
    mActivityValue.SpecialEffect = fmSpecialEffectFlat

    Set mHideButton = Me.Controls.Add("Forms.CommandButton.1", "cmdHideRuntime", True)
    With mHideButton
        .caption = "HIDE RUNTIME"
        .left = MARGIN_X
        .top = 644
        .width = 354
        .height = 28
        .BackColor = RGB(42, 46, 52)
        .foreColor = RGB(232, 232, 232)
        .TakeFocusOnClick = False
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
    End With
End Sub

Private Sub AddSection(ByVal caption As String, ByVal top As Single)
    Dim divider As MSForms.label

    AddText "section" & Replace$(caption, " ", vbNullString), caption, MARGIN_X, top, 180, 18, 9, RGB(230, 198, 106), True
    Set divider = AddText("line" & Replace$(caption, " ", vbNullString), vbNullString, MARGIN_X, top + 21, 354, 1, 1, RGB(90, 75, 40), False)
    divider.BackColor = RGB(90, 75, 40)
    divider.BackStyle = fmBackStyleOpaque
End Sub

Private Function AddRow( _
    ByVal labelText As String, _
    ByVal valueText As String, _
    ByVal top As Single) As MSForms.label

    Dim keyName As String

    keyName = Replace$(Replace$(labelText, " ", vbNullString), "/", vbNullString) _
              & CStr(Me.Controls.count)
    AddText "lblKey" & keyName, labelText, MARGIN_X, top, 106, 17, 9, RGB(156, 163, 170), False
    Set AddRow = AddText("lblValue" & keyName, valueText, VALUE_X, top, 240, 17, 9, RGB(232, 232, 232), False)
End Function

Private Function AddText( _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal left As Single, _
    ByVal top As Single, _
    ByVal width As Single, _
    ByVal height As Single, _
    ByVal fontSize As Single, _
    ByVal foreColor As Long, _
    ByVal isBold As Boolean) As MSForms.label

    Dim item As MSForms.label

    Set item = Me.Controls.Add("Forms.Label.1", controlName, True)
    With item
        .caption = caption
        .left = left
        .top = top
        .width = width
        .height = height
        .BackStyle = fmBackStyleTransparent
        .foreColor = foreColor
        .Font.Name = "Segoe UI"
        .Font.Size = fontSize
        .Font.Bold = isBold
    End With

    Set AddText = item
End Function

Private Sub SetEnabledValue(ByVal target As MSForms.label, ByVal Enabled As Boolean)
    If Enabled Then
        target.caption = "Enabled"
        target.foreColor = RGB(52, 211, 153)
    Else
        target.caption = "Disabled"
        target.foreColor = RGB(156, 163, 170)
    End If
End Sub

Private Sub SetLabelValue(ByVal target As MSForms.label, ByVal value As String)
    target.caption = DisplayValue(value)
    target.ControlTipText = CStr(value)
End Sub

Private Function DisplayValue(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        DisplayValue = "-"
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        DisplayValue = "-"
    Else
        DisplayValue = CStr(value)
    End If
End Function

Private Function KeepFirstLines(ByVal value As String, ByVal maxLines As Long) As String
    Dim lines() As String
    Dim index As Long
    Dim result As String

    lines = Split(value, vbCrLf)
    For index = LBound(lines) To UBound(lines)
        If index - LBound(lines) >= maxLines Then Exit For
        If Len(result) > 0 Then result = result & vbCrLf
        result = result & lines(index)
    Next index

    KeepFirstLines = result
End Function

Private Sub PositionAtExcelRight()
    On Error Resume Next

    Me.StartUpPosition = 0
    Me.left = Application.left + Application.width - Me.width - 18
    Me.top = Application.top + 90
End Sub


