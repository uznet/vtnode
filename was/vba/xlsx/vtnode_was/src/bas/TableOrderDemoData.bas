Attribute VB_Name = "TableOrderDemoData"
Option Explicit
Public Function Orders() As Worksheet
    Set Orders = ThisWorkbook.Worksheets("Order")
End Function
Public Function Total(ByVal tableNo As Long, ByVal sessionId As String) As Currency
    Dim r As Long, ws As Worksheet
    Set ws = Orders()
    For r = 2 To ws.Cells(ws.Rows.count, 2).End(xlUp).row
        If Val(CStr(ws.Cells(r, 4).Value2)) = tableNo And CStr(ws.Cells(r, 3).Value2) = sessionId Then
            If CStr(ws.Cells(r, 12).Value2) <> "PAID" Then Total = Total + CCur(ws.Cells(r, 11).Value2)
        End If
    Next r
End Function
Public Function ValidateItems(ByVal payload As Object, ByVal allowEmpty As Boolean) As Currency
    Dim names As Variant, prices As Variant, options As Variant, item As Variant
    Dim index As Long, id As String, qty As Double, amount As Currency, expected As Currency
    names = Array("ºÒ°í±â Á¤½Ä", "Á¦À°ººÀ½", "±èÄ¡ººÀ½¹ä", "µ¹¼Üºñºö¹ä", "ÀÜÄ¡±¹¼ö", "±èÄ¡Âî°³", "¿Õ¸¸µÎ", "°¨ÀÚÆ¢±è", "ÄÝ¶ó", "¾ÆÀÌ½º ¾Æ¸Þ¸®Ä«³ë")
    prices = Array(12000, 11000, 9000, 10000, 8000, 9500, 6000, 5000, 2500, 3500)
    options = Array("|±âº»|°ø±â¹ä Ãß°¡|´ú ¸Ê°Ô|", "|±âº»|´ú ¸Ê°Ô|¸¹ÀÌ ¸Ê°Ô|", "|±âº»|Ä¡Áî Ãß°¡|°è¶õ Ãß°¡|", "|±âº»|°íÃßÀå µû·Î|°è¶õ Ãß°¡|", "|±âº»|°ö»©±â|±èÄ¡ Ãß°¡|", "|±âº»|´ú ¸Ê°Ô|¶ó¸é»ç¸® Ãß°¡|", "|±âº»|°£Àå Ãß°¡|", "|±âº»|ÄÉÃ¸ Ãß°¡|Ä¡Áî¼Ò½º|", "|±âº»|¾óÀ½ÄÅ|", "|±âº»|¿¬ÇÏ°Ô|¼¦ Ãß°¡|")
    If Not payload.Exists("items") Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Missing items."
    If TypeName(payload("items")) <> "Collection" Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Items must be an array."
    If payload("items").count > 100 Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Too many items."
    If Not allowEmpty And payload("items").count = 0 Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Order is empty."
    For Each item In payload("items")
        id = CStr(item("menuId"))
        index = Val(Mid$(id, 2)) - 1
        If index < 0 Or index > 9 Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Unknown menu."
        If id <> "m" & Format$(index + 1, "00") Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Unknown menu."
        qty = CDbl(item("qty"))
        If qty < 1 Or qty > 99 Or qty <> Fix(qty) Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Invalid quantity."
        If CCur(item("unitPrice")) <> CCur(prices(index)) Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Menu price mismatch."
        If InStr(1, options(index), "|" & CStr(item("option")) & "|", vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Invalid menu option."
        If CStr(item("menuName")) <> names(index) Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Menu name mismatch."
        amount = CCur(prices(index)) * qty
        If CCur(item("amount")) <> amount Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Item amount mismatch."
        expected = expected + amount
    Next item
    If CCur(payload("totalAmount")) <> expected Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Order total mismatch."
    ValidateItems = expected
End Function
Public Function AppendOrder(ByVal payload As Object, ByVal tableNo As Long, ByVal sessionId As String) As Boolean
    Dim ws As Worksheet, row As Long, first As Long, count As Long, i As Long
    Dim Values() As Variant, item As Variant, id As String, stamp As String, people As Long
    Dim existing As Object, newFingerprint As String, oldFingerprint As String
    Dim previousEvents As Boolean
    ValidateItems payload, False
    id = CStr(payload("orderId"))
    If Len(id) = 0 Or Len(id) > 100 Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Invalid order ID."
    people = CLng(payload("peopleCount"))
    If people < 1 Or people > 30 Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Invalid people count."
    stamp = CStr(payload("orderTime"))
    If Len(stamp) = 0 Or Len(stamp) > 50 Then Err.Raise vbObjectError + 2710, "TableOrderDemo", "Invalid order time."
    Set ws = Orders()
    first = ws.Cells(ws.Rows.count, 2).End(xlUp).row + 1
    For Each item In payload("items")
        newFingerprint = newFingerprint & CStr(item("menuId")) & "|" & CStr(item("option")) & "|" & CStr(item("qty")) & ";"
    Next item
    For row = 2 To first - 1
        If CStr(ws.Cells(row, 2).Value2) = id And CStr(ws.Cells(row, 3).Value2) = sessionId Then
            oldFingerprint = oldFingerprint & CStr(ws.Cells(row, 6).Value2) & "|" & CStr(ws.Cells(row, 8).Value2) & "|" & CStr(ws.Cells(row, 9).Value2) & ";"
        End If
    Next row
    If Len(oldFingerprint) > 0 Then
        If oldFingerprint <> newFingerprint Then Err.Raise vbObjectError + 2711, "TableOrderDemo", "Order ID already used for different items."
        Exit Function
    End If
    count = payload("items").count
    ReDim Values(1 To count, 1 To 13)
    For Each item In payload("items")
        i = i + 1
        Values(i, 1) = Format$(Now, "yyyy-mm-dd hh:nn:ss"): Values(i, 2) = id: Values(i, 3) = sessionId
        Values(i, 4) = tableNo: Values(i, 5) = people
        Values(i, 6) = CStr(item("menuId")): Values(i, 7) = CStr(item("menuName"))
        Values(i, 8) = CStr(item("option")): Values(i, 9) = CLng(item("qty"))
        Values(i, 10) = CCur(item("unitPrice")): Values(i, 11) = CCur(item("amount"))
        Values(i, 12) = "PREPARING": Values(i, 13) = Empty
    Next item
    previousEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    ws.Range(ws.Cells(first, 1), ws.Cells(first + count - 1, 3)).NumberFormat = "@"
    ws.Range(ws.Cells(first, 6), ws.Cells(first + count - 1, 8)).NumberFormat = "@"
    ws.Range(ws.Cells(first, 1), ws.Cells(first + count - 1, 13)).Value2 = Values
    Application.EnableEvents = previousEvents
    AppendOrder = True
    Exit Function
Failed:
    Dim errorNumber As Long, errorText As String
    errorNumber = Err.number: errorText = Err.description
    On Error Resume Next
    ws.Range(ws.Cells(first, 1), ws.Cells(first + count - 1, 13)).ClearContents
    Application.EnableEvents = previousEvents
    On Error GoTo 0
    Err.Raise errorNumber, "TableOrderDemoData", errorText
End Function
Public Sub SetStatus(ByVal tableNo As Long, ByVal sessionId As String, ByVal value As String)
    Dim ws As Worksheet, row As Long, previousEvents As Boolean
    Set ws = Orders()
    previousEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    For row = 2 To ws.Cells(ws.Rows.count, 2).End(xlUp).row
        If Val(CStr(ws.Cells(row, 4).Value2)) = tableNo And CStr(ws.Cells(row, 3).Value2) = sessionId Then
            If CStr(ws.Cells(row, 12).Value2) <> "PAID" Then
                ws.Cells(row, 12).Value2 = value
                If value = "PAID" Then ws.Cells(row, 13).value = Now
            End If
        End If
    Next row
    Application.EnableEvents = previousEvents
    Exit Sub
Failed:
    Dim n As Long, description As String
    n = Err.number: description = Err.description
    Application.EnableEvents = previousEvents
    Err.Raise n, "TableOrderDemoData", description
End Sub

