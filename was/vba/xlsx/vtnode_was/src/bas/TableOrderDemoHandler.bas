Attribute VB_Name = "TableOrderDemoHandler"
Option Explicit
Public Function Process(ByVal path As String, ByVal payload As Object) As Object
    Dim result As Object, tableNo As Long, oldNo As Long, sessionId As String
    Dim added As Boolean, Total As Currency, people As Long
    Dim audioType As TableOrderAudioType
    Set result = CreateObject("Scripting.Dictionary")
    If payload.Exists("orderSessionId") Then sessionId = CStr(payload("orderSessionId"))
    If path = "/order/table/assign" Then
        tableNo = TableOrderDemoSession.Assign(sessionId)
        sessionId = TableOrderDemoSession.Token(tableNo)
    Else
        If Not payload.Exists("tableNo") Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Missing table number."
        If CDbl(payload("tableNo")) <> Fix(CDbl(payload("tableNo"))) Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Invalid table number."
        tableNo = CLng(payload("tableNo"))
        If path = "/order/table/select" Then
            oldNo = tableNo
            If payload.Exists("previousTableNo") Then
                If Not IsNull(payload("previousTableNo")) Then oldNo = CLng(payload("previousTableNo"))
            End If
            TableOrderDemoSession.Move oldNo, tableNo, sessionId
        Else
            TableOrderDemoSession.Require tableNo, sessionId
        End If
    End If
    Select Case path
        Case "/order/table/assign", "/order/table/select"
            result("orderSessionId") = sessionId
            result("sessionId") = sessionId
            result("maxAge") = 1800
            result("peopleCount") = 2
        Case "/order/cart/update"
            Total = TableOrderDemoData.ValidateItems(payload, True)
            TableOrderDemoSession.Storage.Cells(tableNo + 1, 27).Value2 = Total
        Case "/order/submit"
            added = TableOrderDemoData.AppendOrder(payload, tableNo, sessionId)
            If added Then
                TableOrderDemoData.SetStatus tableNo, sessionId, "PREPARING"
                TableOrderDemoSession.SetState tableNo, 2
                audioType = AUDIO_ORDER_RECEIVED
            End If
            TableOrderDemoSession.Storage.Cells(tableNo + 1, 27).ClearContents
            result("duplicate") = Not added
            result("orderId") = CStr(payload("orderId"))
        Case "/order/call"
            If Val(CStr(TableOrderDemoSession.Storage.Cells(tableNo + 1, 26).Value2)) = 0 Then audioType = AUDIO_STAFF_CALL
            TableOrderDemoSession.Storage.Cells(tableNo + 1, 26).Value2 = 1
        Case Else
            Err.Raise vbObjectError + 2710, "TableOrderDemo", "Unknown order action."
    End Select
    result("result") = "OK"
    result("tableNo") = tableNo
    TableOrderDemoPanel.Refresh
    If audioType <> AUDIO_NONE Then WasAudio.Play audioType
    Set Process = result
End Function

