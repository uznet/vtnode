Attribute VB_Name = "WasUIWebSocketSessions"
'Attribute VB_Name = "WasUIWebSocketSessions"
Option Explicit

' WebSocket session-table controller for the VTNode WAS worksheet.
' Session table: B30:M35
'
' Sheet event integration:
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Not WasUIWebSocketSessions.IsWebSocketSessionsRange(Target) Then
'           Exit Sub
'       End If
'
'       WasUIWebSocketSessions.OnShellTextChanged Target
'   End Sub
'
' CLEAR CLOSED and SEND are Rectangle Shapes created by Init().
' Their OnAction macros call OnClearClosed() and OnSend().
'
' Each session row keeps its CVtHttpaWebsocket object reference.
' Sending calls CVtHttpaWebsocket.SendText directly.

Private Const SESSION_SHEET As String = "VTNode WAS"
Private Const SESSION_AREA As String = "B27:M35"
Private Const SEND_TEXT_AREA As String = "I30:J35"
Private Const CLEAR_CLOSED_AREA As String = "K27:M28"
Private Const SEND_BUTTON_AREA As String = "K30:K35"
Private Const FIRST_SESSION_ROW As Long = 30
Private Const LAST_SESSION_ROW As Long = 35

Private Const COL_ID As String = "B"
Private Const COL_STATUS As String = "C"
Private Const COL_CONNECTED_AT As String = "E"
Private Const COL_RECEIVED As String = "G"
Private Const COL_SEND_TEXT As String = "I"
Private Const COL_SEND_BUTTON As String = "K"
Private Const COL_LAST_ACTIVITY As String = "L"

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

Public Sub Init()
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    On Error GoTo CleanUp

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    InitSessions
    InitClearClosed
    InitSendButtons

CleanUp:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents

    If Err.Number <> 0 Then
        Debug.Print "[WasUIWebSocketSessions.Init] " & _
                    Err.Number & " : " & Err.description
    End If
End Sub

Public Function IsWebSocketSessionsRange( _
    ByVal target As Range _
) As Boolean

    Dim ws As Worksheet

    On Error GoTo NotSessionRange

    IsWebSocketSessionsRange = False
    If target Is Nothing Then Exit Function

    Set ws = SessionWorksheet()
    If Not (target.Worksheet Is ws) Then Exit Function

    IsWebSocketSessionsRange = Not Intersect( _
        target, _
        ws.Range(SESSION_AREA)) Is Nothing

    Exit Function

NotSessionRange:
    IsWebSocketSessionsRange = False
End Function

Public Function IsClearClosedRange( _
    ByVal target As Range _
) As Boolean
    IsClearClosedRange = IsTargetInArea(target, CLEAR_CLOSED_AREA)
End Function

Public Function IsSendButtonRange( _
    ByVal target As Range _
) As Boolean
    IsSendButtonRange = IsTargetInArea(target, SEND_BUTTON_AREA)
End Function

Public Function IsSendMessageRange( _
    ByVal target As Range _
) As Boolean
    IsSendMessageRange = IsTargetInArea(target, SEND_TEXT_AREA)
End Function

Public Sub OnClearClosed()
    ClearClosed
End Sub

Public Sub OnSend()
    Dim ws As Worksheet
    Dim callerName As String
    Dim sessionRow As Long

    On Error GoTo ErrorHandler

    callerName = CStr(Application.Caller)
    Set ws = SessionWorksheet()
    sessionRow = ws.Shapes(callerName).TopLeftCell.Row

    If sessionRow < FIRST_SESSION_ROW Or _
       sessionRow > LAST_SESSION_ROW Then Exit Sub

    Send sessionRow
    Exit Sub

ErrorHandler:
    Debug.Print "[WasUIWebSocketSessions.OnSend] " & _
                Err.Number & " : " & Err.description
End Sub

Public Sub InitSessions()
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
        ResetSessionRow ws, sessionRow
    Next sessionRow

    SetWaitingRow ws, FIRST_SESSION_ROW
CleanUp:
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEvents

    If Err.Number <> 0 Then
        Debug.Print "[WasUIWebSocketSessions.InitSessions] " & _
                    Err.Number & " : " & Err.description
    End If
End Sub

Public Function OnOpen( _
    ByVal websocket As CVtHttpaWebsocket, _
    ByRef errorOut As String _
) As Boolean

    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim previousEvents As Boolean

    On Error GoTo ErrorHandler

    OnOpen = False
    errorOut = vbNullString
    previousEvents = Application.EnableEvents

    If websocket Is Nothing Then
        errorOut = "The WebSocket object is Nothing."
        Exit Function
    End If

    Set ws = SessionWorksheet()

    If FindRowByWebSocket(websocket) > 0 Then
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
    Set mWebSockets(sessionRow) = websocket
    SessionField(ws, COL_ID, sessionRow).Cells(1, 1).Value2 = _
        "WS-" & Format$(sessionRow - FIRST_SESSION_ROW + 1, "00")
    SessionField(ws, COL_STATUS, sessionRow).Cells(1, 1).Value2 = _
        STATUS_CONNECTED
    SessionField(ws, COL_CONNECTED_AT, sessionRow).Cells(1, 1).Value = Now
    UpdateLastActivity ws, sessionRow

    ' SEND remains disabled until SEND MESSAGE contains text.
    SetSendButtonEnabled ws, sessionRow, False
    EnsureWaitingRow ws

    Application.EnableEvents = previousEvents
    OnOpen = True
    Exit Function

ErrorHandler:
    Application.EnableEvents = previousEvents
    errorOut = "OnOpen failed: " & Err.description
End Function

Public Sub OnClose(ByVal websocket As CVtHttpaWebsocket)
    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim previousEvents As Boolean

    On Error GoTo CleanUp

    If websocket Is Nothing Then Exit Sub

    previousEvents = Application.EnableEvents
    Set ws = SessionWorksheet()
    sessionRow = FindRowByWebSocket(websocket)
    If sessionRow = 0 Then Exit Sub

    Application.EnableEvents = False

    SessionField(ws, COL_STATUS, sessionRow).Cells(1, 1).Value2 = _
        STATUS_CLOSED
    SetSendButtonEnabled ws, sessionRow, False
    UpdateLastActivity ws, sessionRow
    EnsureWaitingRow ws

CleanUp:
    Application.EnableEvents = previousEvents

    If Err.Number <> 0 Then
        Debug.Print "[WasUIWebSocketSessions.OnClose] " & _
                    Err.Number & " : " & Err.description
    End If
End Sub

Public Sub OnRecv( _
    ByVal websocket As CVtHttpaWebsocket, _
    ByVal msg As String _
)
    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim previousEvents As Boolean

    On Error GoTo CleanUp

    If websocket Is Nothing Then Exit Sub

    previousEvents = Application.EnableEvents
    Set ws = SessionWorksheet()
    sessionRow = FindRowByWebSocket(websocket)
    If sessionRow = 0 Then Exit Sub

    Application.EnableEvents = False

    ' RECEIVED MESSAGE contains the latest received text.
    SessionField(ws, COL_RECEIVED, sessionRow).Cells(1, 1).Value2 = msg
    UpdateLastActivity ws, sessionRow

CleanUp:
    Application.EnableEvents = previousEvents

    If Err.Number <> 0 Then
        Debug.Print "[WasUIWebSocketSessions.OnRecv] " & _
                    Err.Number & " : " & Err.description
    End If
End Sub

Public Sub OnShellTextChanged(ByVal changedRange As Range)
    Dim ws As Worksheet
    Dim monitoredRange As Range
    Dim sessionRow As Long
    Dim canSend As Boolean

    On Error GoTo ErrorHandler

    If changedRange Is Nothing Then Exit Sub
    If Not IsSendMessageRange(changedRange) Then Exit Sub

    Set ws = SessionWorksheet()

    Set monitoredRange = Intersect( _
        changedRange, _
        ws.Range(SEND_TEXT_AREA))

    If monitoredRange Is Nothing Then Exit Sub

    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        If Not Intersect( _
            monitoredRange, _
            ws.Range("I" & sessionRow & ":J" & sessionRow)) Is Nothing Then

            canSend = IsConnectedRow(ws, sessionRow) And _
                      Len(Trim$(CStr( _
                          ws.Range(COL_SEND_TEXT & sessionRow).Value2))) > 0

            SetSendButtonEnabled ws, sessionRow, canSend
        End If
    Next sessionRow

    Exit Sub

ErrorHandler:
    Debug.Print "[WasUIWebSocketSessions.OnShellTextChanged] " & _
                Err.Number & " : " & Err.description
End Sub

Public Sub Send(ByVal sessionRow As Long)
    Dim errorMessage As String

    If Not TrySend(sessionRow, errorMessage) Then
        If Len(errorMessage) > 0 Then
            MsgBox errorMessage, vbExclamation, "WebSocket Send"
        End If
    End If
End Sub

Public Function TrySend( _
    ByVal sessionRow As Long, _
    ByRef errorOut As String _
) As Boolean

    Dim ws As Worksheet
    Dim websocket As CVtHttpaWebsocket
    Dim message As String
    Dim previousEvents As Boolean

    On Error GoTo ErrorHandler

    TrySend = False
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

    Set websocket = mWebSockets(sessionRow)
    If websocket Is Nothing Then
        errorOut = "The selected row does not contain a WebSocket object."
        Exit Function
    End If

    message = CStr(ws.Range(COL_SEND_TEXT & sessionRow).Value2)

    If Len(Trim$(message)) = 0 Then
        errorOut = "Enter a message before sending."
        Exit Function
    End If

    If Not websocket.SendText(message) Then
        errorOut = "CVtHttpaWebsocket.SendText returned False."
        Exit Function
    End If

    Application.EnableEvents = False
    SessionField(ws, COL_SEND_TEXT, sessionRow).ClearContents
    SetSendButtonEnabled ws, sessionRow, False
    UpdateLastActivity ws, sessionRow
    Application.EnableEvents = previousEvents

    TrySend = True
    Exit Function

ErrorHandler:
    Application.EnableEvents = previousEvents
    errorOut = "Send failed: " & Err.description
End Function

Public Sub ClearClosed()
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
        Debug.Print "[WasUIWebSocketSessions.ClearClosed] " & _
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
        buttonRange.Left, _
        buttonRange.Top, _
        buttonRange.Width, _
        buttonRange.Height)

    With buttonShape
        .Name = CLEAR_CLOSED_SHAPE
        .OnAction = MacroActionName("OnClearClosed")
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = CLEAR_CLOSED_COLOR
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.text = "CLEAR CLOSED"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = TEXT_ENABLED_COLOR
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
End Sub

Private Sub InitSendButtons()
    Dim ws As Worksheet
    Dim sessionRow As Long
    Dim enabled As Boolean
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
            buttonRange.Left, _
            buttonRange.Top, _
            buttonRange.Width, _
            buttonRange.Height)

        With buttonShape
            .Name = shapeName
            .Placement = xlMoveAndSize
            .Line.Visible = msoFalse
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
        End With

        enabled = IsConnectedRow(ws, sessionRow) And _
                  Len(Trim$(CStr( _
                      ws.Range(COL_SEND_TEXT & sessionRow).Value2))) > 0
        SetSendButtonEnabled ws, sessionRow, enabled
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
    ByVal websocket As CVtHttpaWebsocket _
) As Long

    Dim sessionRow As Long

    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        If Not mWebSockets(sessionRow) Is Nothing Then
            If mWebSockets(sessionRow) Is websocket Then
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
    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        statusText = UCase$(Trim$(CStr( _
            ws.Range(COL_STATUS & sessionRow).Value2)))
        idText = Trim$(CStr(ws.Range(COL_ID & sessionRow).Value2))

        If (Len(idText) = 0 Or idText = "-") And _
           (Len(statusText) = 0 Or statusText = STATUS_WAITING) Then
            FindAvailableRow = sessionRow
            Exit Function
        End If
    Next sessionRow

    ' Second choice: recycle the oldest CLOSED row.
    For sessionRow = FIRST_SESSION_ROW To LAST_SESSION_ROW
        statusText = UCase$(Trim$(CStr( _
            ws.Range(COL_STATUS & sessionRow).Value2)))
        If statusText = STATUS_CLOSED Then
            FindAvailableRow = sessionRow
            Exit Function
        End If
    Next sessionRow
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
    ByVal enabled As Boolean _
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

    If enabled Then
        buttonShape.OnAction = MacroActionName("OnSend")
        buttonShape.Fill.ForeColor.RGB = SEND_ENABLED_COLOR
        buttonShape.TextFrame2.TextRange.text = SEND_ENABLED_TEXT
        buttonShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            TEXT_ENABLED_COLOR
        buttonShape.TextFrame2.TextRange.Font.Bold = msoTrue
    Else
        buttonShape.OnAction = vbNullString
        buttonShape.Fill.ForeColor.RGB = SEND_DISABLED_COLOR
        buttonShape.TextFrame2.TextRange.text = vbNullString
    End If
End Sub

Private Function SendShapeName(ByVal sessionRow As Long) As String
    SendShapeName = SEND_SHAPE_PREFIX & CStr(sessionRow)
End Function

Private Function MacroActionName(ByVal procedureName As String) As String
    MacroActionName = "'" & ThisWorkbook.Name & _
                      "'!WasUIWebSocketSessions." & procedureName
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
    SessionField(ws, COL_LAST_ACTIVITY, sessionRow).Cells(1, 1).Value = Now
End Sub


