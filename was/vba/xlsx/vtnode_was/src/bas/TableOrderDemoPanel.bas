Attribute VB_Name = "TableOrderDemoPanel"
Option Explicit
Private mRefreshing As Boolean
Public Sub SelectTable()
    TableOrderDemo_Show CLng(Right$(CStr(Application.caller), 2))
    WasAudio.Play AUDIO_TICK
End Sub
Public Sub TableOrderDemo_Show(ByVal tableNo As Long)
    If tableNo < 1 Or tableNo > 8 Then Exit Sub
    TableOrderDemoSession.Storage.Range("W1").Value2 = tableNo
    Refresh
End Sub
Public Function StateLabel(ByVal value As Long) As String
    Select Case value
        Case 1: StateLabel = "주문중"
        Case 2: StateLabel = "조리중"
        Case 3: StateLabel = "식사중"
        Case Else: StateLabel = "빈자리"
    End Select
End Function
Public Sub Refresh()
    If mRefreshing Then Exit Sub
    mRefreshing = True
    Dim ws As Worksheet, Orders As Worksheet, n As Long, selected As Long, state As Long
    Dim card As Shape, action As Shape, actionName As String, sid As String
    Dim amount As Currency, active As Long, calls As Long, totalSales As Currency
    Dim row As Long, targetRow As Long, count As Object, dayCount As Object, previousEvents As Boolean
    Dim label As String, called As Boolean, paidAt As Variant, stamp As String
    previousEvents = Application.EnableEvents
    On Error GoTo Done
    Application.EnableEvents = False
    Set ws = TableOrderDemoSession.Storage
    Set Orders = TableOrderDemoData.Orders
    selected = CLng(Val(CStr(ws.Range("W1").Value2)))
    If selected < 1 Or selected > 8 Then selected = 1
    For n = 1 To 8
        sid = TableOrderDemoSession.Token(n)
        state = TableOrderDemoSession.state(n)
        amount = TableOrderDemoData.Total(n, sid) + CCur(Val(CStr(ws.Cells(n + 1, 27).Value2)))
        called = (Val(CStr(ws.Cells(n + 1, 26).Value2)) = 1)
        If Len(sid) > 0 Then active = active + 1
        If called Then calls = calls + 1
        Set card = ws.Shapes("to_table_" & Format$(n, "00"))
        card.TextFrame2.TextRange.text = "Table " & Format$(n, "00") & vbCrLf & StateLabel(state) & vbCrLf & Format$(amount, "#,##0") & "원"
        card.TextFrame2.MarginTop = 1
        card.TextFrame2.MarginBottom = 0
        card.TextFrame2.MarginRight = 65
        card.TextFrame2.TextRange.Font.size = 9
        card.TextFrame2.TextRange.Font.Fill.foreColor.RGB = RGB(234, 243, 255)
        card.OnAction = MacroName("TableOrderDemoPanel.SelectTable")
        card.line.Weight = 1
        card.line.foreColor.RGB = RGB(51, 65, 85)
        If selected = n Then
            card.line.Weight = 2.5
            card.line.foreColor.RGB = RGB(52, 211, 153)
        End If
        If called Then card.Fill.foreColor.RGB = RGB(127, 29, 29) Else card.Fill.foreColor.RGB = RGB(6, 78, 59)
        actionName = "to_action_" & Format$(n, "00")
        Set action = Nothing
        On Error Resume Next
        Set action = ws.Shapes(actionName)
        On Error GoTo Done
        If action Is Nothing Then
            Set action = ws.Shapes.AddShape(msoShapeRectangle, card.left + card.width - 63, card.top + 10, 58, 20)
            action.Name = actionName
        End If
        If called Then
            label = "호출해제"
        Else
            Select Case state
                Case 1: label = "자리해제"
                Case 2: label = "서빙완료"
                Case 3: label = "계산완료"
                Case Else: label = "대기"
            End Select
        End If
        action.TextFrame2.TextRange.text = label
        action.TextFrame2.TextRange.Font.size = 9
        action.TextFrame2.TextRange.Font.Fill.foreColor.RGB = RGB(234, 243, 255)
        action.TextFrame2.VerticalAnchor = msoAnchorMiddle
        action.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        action.TextFrame2.MarginLeft = 1: action.TextFrame2.MarginRight = 1
        action.line.Visible = msoFalse
        action.Fill.foreColor.RGB = RGB(31, 41, 55)
        action.Placement = xlMoveAndSize
        action.OnAction = vbNullString
        If state > 0 Then action.OnAction = MacroName("TableOrderDemoPanel.ActionClick")
    Next n
    Set dayCount = CreateObject("Scripting.Dictionary")
    For row = 2 To Orders.Cells(Orders.Rows.count, 2).End(xlUp).row
        stamp = CStr(Orders.Cells(row, 1).value)
        If IsDate(stamp) Then
            If DateValue(stamp) = Date Then dayCount(CStr(Orders.Cells(row, 3).Value2) & ":" & CStr(Orders.Cells(row, 2).Value2)) = True
        ElseIf left$(stamp, 10) = Format$(Date, "yyyy-mm-dd") Then
            dayCount(CStr(Orders.Cells(row, 3).Value2) & ":" & CStr(Orders.Cells(row, 2).Value2)) = True
        End If
        If CStr(Orders.Cells(row, 12).Value2) = "PAID" Then
            paidAt = Orders.Cells(row, 13).value
            If IsDate(paidAt) Then
                If DateValue(paidAt) = Date Then totalSales = totalSales + CCur(Orders.Cells(row, 11).Value2)
            End If
        End If
    Next row
    ws.Range("H7").Value2 = active: ws.Range("K7").Value2 = calls
    ws.Range("H9").Value2 = dayCount.count: ws.Range("K9").Value2 = totalSales
    ws.Range("K9").NumberFormat = "#,##0""원"""
    ws.Range("O8:U28").ClearContents
    ws.Range("O6").Value2 = "T" & Format$(selected, "00")
    Set count = CreateObject("Scripting.Dictionary")
    targetRow = 8
    sid = TableOrderDemoSession.Token(selected)
    amount = 0
    For row = 2 To Orders.Cells(Orders.Rows.count, 2).End(xlUp).row
        If Val(CStr(Orders.Cells(row, 4).Value2)) = selected And CStr(Orders.Cells(row, 3).Value2) = sid And Len(sid) > 0 Then
            If CStr(Orders.Cells(row, 12).Value2) <> "PAID" Then
                count(CStr(Orders.Cells(row, 2).Value2)) = True
                amount = amount + CCur(Orders.Cells(row, 11).Value2)
                If targetRow <= 28 Then
                    ws.Cells(targetRow, 15).Value2 = CStr(Orders.Cells(row, 1).text)
                    ws.Cells(targetRow, 16).Value2 = Orders.Cells(row, 7).Value2
                    ws.Cells(targetRow, 17).Value2 = Orders.Cells(row, 8).Value2
                    ws.Cells(targetRow, 18).Value2 = Orders.Cells(row, 9).Value2
                    ws.Cells(targetRow, 19).Value2 = Orders.Cells(row, 10).Value2
                    ws.Cells(targetRow, 20).Value2 = Orders.Cells(row, 11).Value2
                    ws.Cells(targetRow, 21).Value2 = StatusLabel(CStr(Orders.Cells(row, 12).Value2))
                End If
                targetRow = targetRow + 1
            End If
        End If
    Next row
    ws.Range("P6").Value2 = CStr(count.count) & "건"
    ws.Range("R6").Value2 = amount
    ws.Range("O30:U30").ClearContents
    If targetRow = 8 Then ws.Range("O8").Value2 = "현재 이용 중인 주문이 없습니다."
    If targetRow > 29 Then ws.Range("O30").Value2 = "첫 21행 표시 · 전체 목록은 Order 시트에서 확인"
Done:
    If Err.number <> 0 Then Debug.Print "[TableOrderDemoPanel.Refresh] " & Err.description
    Application.EnableEvents = previousEvents
    mRefreshing = False
End Sub
Public Function StatusLabel(ByVal value As String) As String
    Select Case value
        Case "PREPARING": StatusLabel = "조리중"
        Case "DINING": StatusLabel = "식사중"
        Case "PAID": StatusLabel = "계산완료"
        Case Else: StatusLabel = value
    End Select
End Function
Public Function MacroName(ByVal procedure As String) As String
    MacroName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procedure
End Function
Public Sub ActionClick()
    Dim n As Long
    n = CLng(Right$(CStr(Application.caller), 2))
    If TableOrderDemoSession.state(n) = 3 And Val(CStr(TableOrderDemoSession.Storage.Cells(n + 1, 26).Value2)) = 0 Then
        If MsgBox("T" & Format$(n, "00") & " 계산을 완료할까요?", vbQuestion + vbYesNo, "VTWas") <> vbYes Then Exit Sub
    End If
    Advance n
End Sub
Public Sub Advance(ByVal n As Long)
    Dim sid As String, ws As Worksheet
    Dim audioType As TableOrderAudioType
    If n < 1 Or n > 8 Then Exit Sub
    audioType = AUDIO_CLICK_CONFIRM
    Set ws = TableOrderDemoSession.Storage
    sid = TableOrderDemoSession.Token(n)
    If Val(CStr(ws.Cells(n + 1, 26).Value2)) = 1 Then
        ws.Cells(n + 1, 26).ClearContents
    Else
        Select Case TableOrderDemoSession.state(n)
            Case 1
                If TableOrderDemoData.Total(n, sid) = 0 Then TableOrderDemoSession.Release n
            Case 2
                TableOrderDemoData.SetStatus n, sid, "DINING"
                TableOrderDemoSession.SetState n, 3
            Case 3
                TableOrderDemoData.SetStatus n, sid, "PAID"
                TableOrderDemoSession.Release n
                audioType = AUDIO_PAYMENT_COMPLETE
        End Select
    End If
    TableOrderDemo_Show n
    WasAudio.Play audioType
End Sub

