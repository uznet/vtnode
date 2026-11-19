Attribute VB_Name = "TableOrderDemoRouter"
Option Explicit
Public Function Handle(ByVal session As CVtHttpaSession, ByVal custom As CWasCustomData) As Boolean
    Dim path As String, payload As Object, result As Object, Content As Object
    Dim errorNumber As Long, errorText As String, errorSource As String, code As Long
    path = LCase$(session.Request.path)
    If path <> "/order" And path <> "/order/" And left$(path, 7) <> "/order/" Then Exit Function
    Handle = True
    On Error GoTo Failed
    If path = "/order" Or path = "/order/" Then
        If session.Request.Method <> VTHTTPA_METHOD_GET Then
            Set result = Failure("Method not allowed.")
            Reply session, custom, result, 405
        Else
            WasHttpaServerEvent.RespondFile session, "/", "table_order.html", custom
        End If
        Exit Function
    End If
    Select Case path
        Case "/order/table/assign", "/order/table/select", "/order/cart/update", "/order/submit", "/order/call"
        Case Else
            Set result = Failure("Unknown order route.")
            Reply session, custom, result, 404
            Exit Function
    End Select
    If session.Request.Method <> VTHTTPA_METHOD_POST Then
        Set result = Failure("POST is required.")
        Reply session, custom, result, 405
        Exit Function
    End If
    Set Content = session.Request.Body
    If Content Is Nothing Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Request body is missing."
    If InStr(1, Content.ContentType, "application/json", vbTextCompare) = 0 Then
        Set result = Failure("JSON content type is required.")
        Reply session, custom, result, 415
        Exit Function
    End If
    Dim text As String
    text = Content.GetText()
    If Len(text) > 262144 Then
        Set result = Failure("Request is too large.")
        Reply session, custom, result, 413
        Exit Function
    End If
    Set payload = JsonConverter.ParseJson(text)
    If TypeName(payload) <> "Dictionary" Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "JSON object is required."
    Set result = TableOrderDemoHandler.Process(path, payload)
    Reply session, custom, result, 200
    Exit Function
Failed:
    errorNumber = Err.number: errorText = Err.description: errorSource = Err.Source
    code = 400
    If errorSource = "TableOrderDemoData" Then code = 500
    If errorNumber = vbObjectError + 2701 Or errorNumber = vbObjectError + 2704 Or errorNumber = vbObjectError + 2711 Then code = 409
    Set result = Failure(errorText)
    If errorNumber = vbObjectError + 2703 Then
        result("eventType") = "SESSION_TIMEOUT"
        code = 409
    End If
    Reply session, custom, result, code
End Function
Private Function Failure(ByVal message As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("result") = "FAIL"
    result("message") = message
    Set Failure = result
End Function
Private Sub Reply(ByVal session As CVtHttpaSession, ByVal custom As CWasCustomData, ByVal value As Object, ByVal code As Long)
    Dim Response As CVtHttpaResponse, errorText As String
    Set Response = VtHttpaFactory.VtHttpaCreateResponse()
    Response.StatusCode = code
    Response.Headers.SetHeader "Content-Type", "application/json; charset=utf-8"
    Response.Headers.SetHeader "Cache-Control", "no-store"
    If Response.SendText(session, JsonConverter.ConvertToJson(value), errorText) Then
        custom.ResponseStatus = "TableOrder JSON(" & CStr(code) & ") -> OK"
    Else
        custom.ResponseStatus = "TableOrder JSON(" & CStr(code) & ") -> Failed: " & errorText
        VtLogError "TableOrderDemoRouter", "Reply", errorText
    End If
    WasUILogger.OnResponseLog custom.RequestTimestamp, Response, custom.ResponseStatus
End Sub

