Attribute VB_Name = "WebsocketDemo"
Option Explicit


' WebSocket session-table controller for the Websocket worksheet.
' Session table: B17:H24. Row 17 is reserved for WAITING; sessions use 18:24.
'
' Sheet event integration:
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Not WebsocketDemo.WebsocketDemo_IsWebSocketSessionsRange(Target) Then
'           Exit Sub
'       End If
'
'       WebsocketDemo.WebsocketDemo_OnShellTextChanged Target
'   End Sub
'
' CLEAR CLOSED and SEND are Rectangle Shapes created by Init().
' Their OnAction macros call WebsocketDemo_OnClearClosed() and WebsocketDemo_OnSend().
'
' Each session row keeps its CVtHttpaWebsocket object reference.
' Sending calls CVtHttpaWebsocket.SendText directly.

Private Const SESSION_SHEET As String = "Websocket"
Private Const SESSION_AREA As String = "B13:H24"
Private Const SEND_TEXT_AREA As String = "F17:F24"
Private Const CLEAR_CLOSED_AREA As String = "G13:H13"
Private Const SEND_BUTTON_AREA As String = "G17:G24"
Private Const FIRST_SESSION_ROW As Long = 17
Private Const LAST_SESSION_ROW As Long = 24

Private Const COL_ID As String = "B"
Private Const COL_STATUS As String = "C"
Private Const COL_CONNECTED_AT As String = "D"
Private Const COL_RECEIVED As String = "E"
Private Const COL_SEND_TEXT As String = "F"
Private Const COL_SEND_BUTTON As String = "G"
Private Const COL_LAST_ACTIVITY As String = "H"

Private Const STATUS_CONNECTED As String = "CONNECTED"
Private Const STATUS_CLOSED As String = "CLOSED"
Private Const STATUS_WAITING As String = "WAITING"

Private Const SEND_ENABLED_TEXT As String = "SEND"
Private Const SEND_ENABLED_COLOR As Long = 15230766   ' RGB(46, 103, 232)
Private Const SEND_DISABLED_COLOR As Long = 16447477  ' RGB(245, 247, 250)
Private Const TEXT_ENABLED_COLOR As Long = 16777215   ' RGB(255, 255, 255)
Private Const TEXT_DISABLED_COLOR As Long = 8947848   ' RGB(136, 136, 136)
Private Const CLEAR_CLOSED_COLOR As Long = 4334356     ' RGB(20, 35, 66)
Private Const CLEAR_CLOSED_SHAPE As String = "btnWsClearClosed"
Private Const SEND_SHAPE_PREFIX As String = "btnWsSend_"

Private mWebSockets( _
    FIRST_SESSION_ROW To LAST_SESSION_ROW _
) As CVtHttpaWebsocket

Public Sub WebsocketDemo_Activate()

    On Error GoTo ErrorHandler

    ThisWorkbook.Worksheets("WEBSOCKET").Activate
    Exit Sub

ErrorHandler:
    MsgBox "Failed to activate WEBSOCKET worksheet." & vbCrLf & _
           Err.description, vbExclamation, "VTNode WAS"

End Sub


Public Sub WebsocketDemo_Init()
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    On Error GoTo CleanUp

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    WebsocketDemo_InitSessions
    InitClearClosed
    InitSendButtons

CleanUp:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents

    If Err.Number <> 0 Then
        Debug.Print "[WebsocketDemo.Init] " & _
                    Err.Number & " : " & Err.description
    End If
End Sub

Public Function WebsocketDemo_IsWebSocketSessionsRange( _
    ByVal target As Range _
) As Boolean

    Dim ws As Worksheet

    On Error GoTo NotSessionRange

    WebsocketDemo_IsWebSocketSessionsRange = False
    If target Is Nothing Then Exit Function

    Set ws = SessionWorksheet()
    If Not (target.Worksheet Is ws) Then Exit Function

    WebsocketDemo_IsWebSocketSessionsRange = Not Intersect( _
        target, _
        ws.Range(SESSION_AREA)) Is Nothing

    Exit Function

NotSessionRange:
    WebsocketDemo_IsWebSocketSessionsRange = False
End Function

Public Function WebsocketDemo_IsClearClosedRange( _
    ByVal target As Range _
) As Boolean
    WebsocketDemo_IsClearClosedRange = IsTargetInArea(target, CLEAR_CLOSED_AREA)
End Function

Public Function WebsocketDemo_IsSendButtonRange( _
    ByVal target As Range _
) As Boolean
    WebsocketDemo_IsSendButtonRange = IsTargetInArea(target, SEND_BUTTON_AREA)
End Function

Public Function WebsocketDemo_IsSendMessageRange( _
    ByVal target As Range _
) As Boolean
    WebsocketDemo_IsSendMessageRange = IsTargetInArea(target, SEND_TEXT_AREA)
End Function

Public Sub WebsocketDemo_OnClearClosed()
    WebsocketDemo_ClearClosed
End Sub

Public Sub WebsocketDemo_OnSend()
    Dim ws As Worksheet
    Dim callerName As String
    Dim sessionRow As Long

    On Error GoTo ErrorHandler

    callerName = CStr(Application.Caller)
    Set ws = SessionWorksheet()
    sessionRow = ws.Shapes(callerName).TopLeftCell.Row

    If sessionRow < FIRST_SESSION_ROW Or _
       sessionRow > LAST_SESSION_ROW Then Exit Sub

    WebsocketDemo_Send sessionRow
    Exit Sub

ErrorHandler:
    Debug.Print "[WebsocketDemo.OnSend] " & _
                Err.Number & " : " & Err.description
End Sub

Public Sub WebsocketDemo_InitSessions()
    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    On Error GoTo CleanUp

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Set ws = SessionWorksheet()
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        If Not mWebSockets(sessionRow) Is Nothing Then
            Err.Raise vbObjectError + 1701, "WebsocketDemo", "Cannot initialize while sessions are connected."
        End If
    Next sessionRow
    ws.Range(SEND_TEXT_AREA).NumberFormat = "@"
    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        ResetSessionRow ws, sessionRow
    Next sessionRow

    SetWaitingRow ws, FIRST_SESSION_ROW
CleanUp:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents

    If Err.Number <> 0 Then
        Debug.Print "[WebsocketDemo.InitSessions] " & _
                    Err.Number & " : " & Err.description
    End If
End Sub

Public Function WebsocketDemo_OnOpen( _
    ByVal Websocket As CVtHttpaWebsocket, _
    ByRef errorOut As String _
) As Boolean

    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim previousEvents As Boolean

    On Error GoTo ErrorHandler

    WebsocketDemo_OnOpen = False
    errorOut = vbNullString
    previousEvents = Application.EnableEvents

    If Websocket Is Nothing Then
        errorOut = "The WebSocket object is Nothing."
        Exit Function
    End If

    Set ws = SessionWorksheet()

    If FindRowByWebSocket(Websocket) > 0 Then
        errorOut = "The WebSocket object is already registered."
        Exit Function
    End If

    sessionRow = FindAvailableRow(ws)
    If sessionRow = 0 Then
        errorOut = "No available WebSocket session row."
        Exit Function
    End If

    Application.EnableEvents = False

    ResetSessionRow ws, sessionRow
    Set mWebSockets(sessionRow) = Websocket
    SessionField(ws, COL_ID, sessionRow).Cells(1, 1).Value2 = _
        "WS-" & Format$(sessionRow - FIRST_SESSION_ROW, "00")
    SessionField(ws, COL_STATUS, sessionRow).Cells(1, 1).Value2 = _
        STATUS_CONNECTED
    SessionField(ws, COL_CONNECTED_AT, sessionRow).Cells(1, 1).value = Now
    UpdateLastActivity ws, sessionRow

    ' SEND remains disabled until SEND MESSAGE contains text.
    SetSendButtonEnabled ws, sessionRow, False
    EnsureWaitingRow ws

    Application.EnableEvents = previousEvents
    WebsocketDemo_OnOpen = True
    Exit Function

ErrorHandler:
    Application.EnableEvents = previousEvents
    errorOut = "WebsocketDemo_OnOpen failed: " & Err.description
End Function

Public Sub WebsocketDemo_OnClose(ByVal Websocket As CVtHttpaWebsocket)
    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim previousEvents As Boolean

    On Error GoTo CleanUp

    If Websocket Is Nothing Then Exit Sub

    previousEvents = Application.EnableEvents
    Set ws = SessionWorksheet()
    sessionRow = FindRowByWebSocket(Websocket)
    If sessionRow = 0 Then Exit Sub

    Application.EnableEvents = False

    SessionField(ws, COL_STATUS, sessionRow).Cells(1, 1).Value2 = _
        STATUS_CLOSED
    Set mWebSockets(sessionRow) = Nothing
    SetSendButtonEnabled ws, sessionRow, False
    UpdateLastActivity ws, sessionRow
    EnsureWaitingRow ws

CleanUp:
    Application.EnableEvents = previousEvents

    If Err.Number <> 0 Then
        Debug.Print "[WebsocketDemo.OnClose] " & _
                    Err.Number & " : " & Err.description
    End If
End Sub

Public Sub WebsocketDemo_OnRecv( _
    ByVal Websocket As CVtHttpaWebsocket, _
    ByVal msg As String _
)
    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim previousEvents As Boolean

    On Error GoTo CleanUp

    If Websocket Is Nothing Then Exit Sub

    previousEvents = Application.EnableEvents
    Set ws = SessionWorksheet()
    sessionRow = FindRowByWebSocket(Websocket)
    If sessionRow = 0 Then Exit Sub

    Application.EnableEvents = False

    ' RECEIVED MESSAGE contains the latest received text.
    With SessionField(ws, COL_RECEIVED, sessionRow).Cells(1, 1)
        .NumberFormat = "@"
        .Value2 = Left$(msg, 32767)
    End With
    UpdateLastActivity ws, sessionRow

CleanUp:
    Application.EnableEvents = previousEvents

    If Err.Number <> 0 Then
        Debug.Print "[WebsocketDemo.OnRecv] " & _
                    Err.Number & " : " & Err.description
    End If
End Sub

Public Sub WebsocketDemo_OnShellTextChanged(ByVal changedRange As Range)
    Dim ws As Worksheet
    Dim monitoredRange As Range
    Dim sessionRow As Long
    Dim canSend As Boolean

    On Error GoTo ErrorHandler

    If changedRange Is Nothing Then Exit Sub
    If Not WebsocketDemo_IsSendMessageRange(changedRange) Then Exit Sub

    Set ws = SessionWorksheet()

    Set monitoredRange = Intersect( _
        changedRange, _
        ws.Range(SEND_TEXT_AREA))

    If monitoredRange Is Nothing Then Exit Sub

    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        If Not Intersect( _
            monitoredRange, _
            ws.Range(COL_SEND_TEXT & sessionRow)) Is Nothing Then

            canSend = IsConnectedRow(ws, sessionRow) And _
                      HasSendText(ws, sessionRow)

            SetSendButtonEnabled ws, sessionRow, canSend
        End If
    Next sessionRow

    Exit Sub

ErrorHandler:
    Debug.Print "[WebsocketDemo.WebsocketDemo_OnShellTextChanged] " & _
                Err.Number & " : " & Err.description
End Sub

Public Sub WebsocketDemo_Send(ByVal sessionRow As Long)
    Dim ErrorMessage As String

    If Not WebsocketDemo_TrySend(sessionRow, ErrorMessage) Then
        If Len(ErrorMessage) > 0 Then
            MsgBox ErrorMessage, vbExclamation, "WebSocket WebsocketDemo_Send"
        End If
    End If
End Sub

Public Function WebsocketDemo_TrySend( _
    ByVal sessionRow As Long, _
    ByRef errorOut As String _
) As Boolean

    Dim ws As Worksheet
    Dim Websocket As CVtHttpaWebsocket
    Dim message As String
    Dim previousEvents As Boolean

    On Error GoTo ErrorHandler

    WebsocketDemo_TrySend = False
    errorOut = vbNullString
    previousEvents = Application.EnableEvents

    If sessionRow < FIRST_SESSION_ROW Or _
       sessionRow > LAST_SESSION_ROW Then
        errorOut = "Invalid session row: " & CStr(sessionRow)
        Exit Function
    End If

    Set ws = SessionWorksheet()

    If Not IsConnectedRow(ws, sessionRow) Then
        errorOut = "The selected WebSocket session is not connected."
        Exit Function
    End If

    Set Websocket = mWebSockets(sessionRow)
    If Websocket Is Nothing Then
        errorOut = "The selected row does not contain a WebSocket object."
        Exit Function
    End If

    If IsError(ws.Range(COL_SEND_TEXT & sessionRow).Value2) Then
        errorOut = "The message cell contains an Excel error."
        Exit Function
    End If
    message = CStr(ws.Range(COL_SEND_TEXT & sessionRow).Value2)

    If Len(Trim$(message)) = 0 Then
        errorOut = "Enter a message before sending."
        Exit Function
    End If

    If Not Websocket.SendText(message) Then
        errorOut = "CVtHttpaWebsocket.SendText returned False."
        Exit Function
    End If

    Application.EnableEvents = False
    SessionField(ws, COL_SEND_TEXT, sessionRow).ClearContents
    SetSendButtonEnabled ws, sessionRow, False
    UpdateLastActivity ws, sessionRow
    Application.EnableEvents = previousEvents

    WebsocketDemo_TrySend = True
    Exit Function

ErrorHandler:
    Application.EnableEvents = previousEvents
    errorOut = "WebsocketDemo_Send failed: " & Err.description
End Function

Public Sub WebsocketDemo_ClearClosed()
    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    On Error GoTo CleanUp

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Set ws = SessionWorksheet()
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        If UCase$(Trim$(CStr( _
            ws.Range(COL_STATUS & sessionRow).Value2))) = STATUS_CLOSED Then
            ResetSessionRow ws, sessionRow
        End If
    Next sessionRow

    EnsureWaitingRow ws

CleanUp:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents

    If Err.Number <> 0 Then
        Debug.Print "[WebsocketDemo.ClearClosed] " & _
                    Err.Number & " : " & Err.description
    End If
End Sub

Private Function SessionWorksheet() As Worksheet
    Set SessionWorksheet = ThisWorkbook.Worksheets(SESSION_SHEET)
End Function

Private Function SessionField( _
    ByVal ws As Worksheet, _
    ByVal columnName As String, _
    ByVal sessionRow As Long _
) As Range
    Set SessionField = ws.Range(columnName & sessionRow).MergeArea
End Function

Private Sub InitClearClosed()
    Dim ws As Worksheet
    Dim buttonRange As Range
    Dim buttonShape As Shape

    Set ws = SessionWorksheet()
    Set buttonRange = ws.Range(CLEAR_CLOSED_AREA)
    buttonRange.ClearContents

    DeleteShapeIfExists ws, CLEAR_CLOSED_SHAPE

    Set buttonShape = ws.Shapes.AddShape( _
        msoShapeRectangle, _
        buttonRange.left, _
        buttonRange.top, _
        buttonRange.width, _
        buttonRange.height)

    With buttonShape
        .Name = CLEAR_CLOSED_SHAPE
        .OnAction = MacroActionName("WebsocketDemo_OnClearClosed")
        .Placement = xlMoveAndSize
        .Fill.foreColor.RGB = CLEAR_CLOSED_COLOR
        .line.Visible = msoFalse
        .TextFrame2.TextRange.text = "CLEAR CLOSED"
        .TextFrame2.TextRange.Font.Fill.foreColor.RGB = TEXT_ENABLED_COLOR
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
End Sub

Private Sub InitSendButtons()
    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim Enabled As Boolean
    Dim buttonRange As Range
    Dim buttonShape As Shape
    Dim shapeName As String

    Set ws = SessionWorksheet()

    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        shapeName = SendShapeName(sessionRow)
        DeleteShapeIfExists ws, shapeName

        Set buttonRange = ws.Range( _
            COL_SEND_BUTTON & sessionRow)
        Set buttonShape = ws.Shapes.AddShape( _
            msoShapeRectangle, _
            buttonRange.left, _
            buttonRange.top, _
            buttonRange.width, _
            buttonRange.height)

        With buttonShape
            .Name = shapeName
            .Placement = xlMoveAndSize
            .line.Visible = msoFalse
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
        End With

        Enabled = IsConnectedRow(ws, sessionRow) And _
                  HasSendText(ws, sessionRow)
        SetSendButtonEnabled ws, sessionRow, Enabled
    Next sessionRow
End Sub

Private Function IsTargetInArea( _
    ByVal target As Range, _
    ByVal areaAddress As String _
) As Boolean
    Dim ws As Worksheet

    On Error GoTo NotInArea

    IsTargetInArea = False
    If target Is Nothing Then Exit Function

    Set ws = SessionWorksheet()
    If Not (target.Worksheet Is ws) Then Exit Function

    IsTargetInArea = Not Intersect( _
        target, _
        ws.Range(areaAddress)) Is Nothing
    Exit Function

NotInArea:
    IsTargetInArea = False
End Function

Private Function FindRowByWebSocket( _
    ByVal Websocket As CVtHttpaWebsocket _
) As Long

    Dim sessionRow As Long

    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        If Not mWebSockets(sessionRow) Is Nothing Then
            If mWebSockets(sessionRow) Is Websocket Then
                FindRowByWebSocket = sessionRow
                Exit Function
            End If
        End If
    Next sessionRow
End Function

Private Function FindAvailableRow(ByVal ws As Worksheet) As Long
    Dim sessionRow As Long
    Dim statusText As String
    Dim idText As String

    ' First choice: an unused or WAITING row.
    For sessionRow = FIRST_SESSION_ROW + 1 To LAST_SESSION_ROW
        statusText = UCase$(Trim$(CStr( _
            ws.Range(COL_STATUS & sessionRow).Value2)))
        idText = Trim$(CStr(ws.Range(COL_ID & sessionRow).Value2))

        If (Len(idText) = 0 Or idText = "-") And _
           (Len(statusText) = 0 Or statusText = STATUS_WAITING) Then
            FindAvailableRow = sessionRow
            Exit Function
        End If
    Next sessionRow

    ' CLOSED rows are retained until CLEAR CLOSED is selected.
End Function

Private Sub ResetSessionRow( _
    ByVal ws As Worksheet, _
    ByVal sessionRow As Long _
)
    SessionField(ws, COL_ID, sessionRow).ClearContents
    SessionField(ws, COL_STATUS, sessionRow).ClearContents
    SessionField(ws, COL_CONNECTED_AT, sessionRow).ClearContents
    SessionField(ws, COL_RECEIVED, sessionRow).ClearContents
    SessionField(ws, COL_SEND_TEXT, sessionRow).ClearContents
    SessionField(ws, COL_LAST_ACTIVITY, sessionRow).ClearContents
    Set mWebSockets(sessionRow) = Nothing
    SetSendButtonEnabled ws, sessionRow, False
End Sub

Private Sub SetWaitingRow( _
    ByVal ws As Worksheet, _
    ByVal sessionRow As Long _
)
    ResetSessionRow ws, sessionRow
    SessionField(ws, COL_ID, sessionRow).Cells(1, 1).Value2 = "-"
    SessionField(ws, COL_STATUS, sessionRow).Cells(1, 1).Value2 = _
        STATUS_WAITING
    SessionField(ws, COL_CONNECTED_AT, sessionRow).Cells(1, 1).Value2 = "-"
    SessionField(ws, COL_RECEIVED, sessionRow).Cells(1, 1).Value2 = _
        "Waiting for connection..."
    SessionField(ws, COL_LAST_ACTIVITY, sessionRow).Cells(1, 1).Value2 = "-"
End Sub

Private Sub EnsureWaitingRow(ByVal ws As Worksheet)
    Dim sessionRow As Long
    Dim statusText As String
    Dim idText As String

    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        statusText = UCase$(Trim$(CStr( _
            ws.Range(COL_STATUS & sessionRow).Value2)))
        If statusText = STATUS_WAITING Then Exit Sub
    Next sessionRow

    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        statusText = Trim$(CStr( _
            ws.Range(COL_STATUS & sessionRow).Value2))
        idText = Trim$(CStr(ws.Range(COL_ID & sessionRow).Value2))

        If Len(statusText) = 0 And Len(idText) = 0 Then
            SetWaitingRow ws, sessionRow
            Exit Sub
        End If
    Next sessionRow
End Sub

Private Function IsConnectedRow( _
    ByVal ws As Worksheet, _
    ByVal sessionRow As Long _
) As Boolean
    IsConnectedRow = _
        UCase$(Trim$(CStr( _
            ws.Range(COL_STATUS & sessionRow).Value2))) = STATUS_CONNECTED
End Function

Private Sub SetSendButtonEnabled( _
    ByVal ws As Worksheet, _
    ByVal sessionRow As Long, _
    ByVal Enabled As Boolean _
)
    Dim buttonRange As Range
    Dim buttonShape As Shape
    Dim shapeName As String

    Set buttonRange = ws.Range( _
        COL_SEND_BUTTON & sessionRow)

    shapeName = SendShapeName(sessionRow)
    If Not ShapeExists(ws, shapeName) Then
        buttonRange.ClearContents
        Exit Sub
    End If

    Set buttonShape = ws.Shapes(shapeName)
    buttonRange.ClearContents

    If Enabled Then
        buttonShape.OnAction = MacroActionName("WebsocketDemo_OnSend")
        buttonShape.Fill.foreColor.RGB = SEND_ENABLED_COLOR
        buttonShape.TextFrame2.TextRange.text = SEND_ENABLED_TEXT
        buttonShape.TextFrame2.TextRange.Font.Fill.foreColor.RGB = _
            TEXT_ENABLED_COLOR
        buttonShape.TextFrame2.TextRange.Font.Bold = msoTrue
    Else
        buttonShape.OnAction = vbNullString
        buttonShape.Fill.foreColor.RGB = SEND_DISABLED_COLOR
        buttonShape.TextFrame2.TextRange.text = vbNullString
    End If
End Sub

Private Function SendShapeName(ByVal sessionRow As Long) As String
    SendShapeName = SEND_SHAPE_PREFIX & CStr(sessionRow)
End Function

Private Function MacroActionName(ByVal procedureName As String) As String
    MacroActionName = "'" & Replace(ThisWorkbook.Name, "'", "''") & _
                      "'!WebsocketDemo." & procedureName
End Function

Private Function ShapeExists( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String _
) As Boolean
    Dim targetShape As Shape

    On Error Resume Next
    Set targetShape = ws.Shapes(shapeName)
    ShapeExists = Not targetShape Is Nothing
    Set targetShape = Nothing
    On Error GoTo 0
End Function

Private Sub DeleteShapeIfExists( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String _
)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub UpdateLastActivity( _
    ByVal ws As Worksheet, _
    ByVal sessionRow As Long _
)
    SessionField(ws, COL_LAST_ACTIVITY, sessionRow).Cells(1, 1).value = Now
End Sub





Private Function HasSendText(ByVal ws As Worksheet, ByVal sessionRow As Long) As Boolean
    Dim value As Variant
    value = ws.Range(COL_SEND_TEXT & sessionRow).Value2
    If IsError(value) Then Exit Function
    HasSendText = Len(Trim$(CStr(value))) > 0
End Function

Public Sub WebsocketDemo_Uninit()
    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim previousEvents As Boolean
    previousEvents = Application.EnableEvents
    On Error GoTo CleanUp
    Set ws = SessionWorksheet()
    Application.EnableEvents = False
    For sessionRow = FIRST_SESSION_ROW + 1 To LAST_SESSION_ROW
        If Not mWebSockets(sessionRow) Is Nothing Then
            SessionField(ws, COL_STATUS, sessionRow).Value2 = STATUS_CLOSED
            Set mWebSockets(sessionRow) = Nothing
            UpdateLastActivity ws, sessionRow
        End If
        SetSendButtonEnabled ws, sessionRow, False
    Next sessionRow
CleanUp:
    Application.EnableEvents = previousEvents
    If Err.Number <> 0 Then Debug.Print "[WebsocketDemo.Uninit] " & Err.Description
End Sub