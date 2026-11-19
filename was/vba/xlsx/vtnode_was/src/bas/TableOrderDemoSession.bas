Attribute VB_Name = "TableOrderDemoSession"
Option Explicit
Private Type DemoGuid
    A As Long
    b As Integer
    C As Integer
    D(0 To 7) As Byte
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef value As DemoGuid) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef value As DemoGuid, ByVal buffer As LongPtr, ByVal size As Long) As Long

Public Function Storage() As Worksheet
    Set Storage = ThisWorkbook.Worksheets("TableOrder")
End Function
Public Function Token(ByVal tableNo As Long) As String
    Token = CStr(Storage.Cells(tableNo + 1, 24).Value2)
End Function
Public Function state(ByVal tableNo As Long) As Long
    state = CLng(Val(CStr(Storage.Cells(tableNo + 1, 25).Value2)))
End Function
Public Sub SetState(ByVal tableNo As Long, ByVal value As Long)
    Storage.Cells(tableNo + 1, 25).Value2 = value
End Sub
Public Sub Release(ByVal tableNo As Long)
    Storage.Range(Storage.Cells(tableNo + 1, 24), Storage.Cells(tableNo + 1, 29)).ClearContents
End Sub
Public Sub ExpireEmptySessions()
    Dim n As Long, lastSeen As Variant
    For n = 1 To 8
        If Len(Token(n)) > 0 Then
            lastSeen = Storage.Cells(n + 1, 29).value
            If IsDate(lastSeen) Then
                If DateDiff("n", CDate(lastSeen), Now) >= 30 Then
                    If TableOrderDemoData.Total(n, Token(n)) = 0 Then Release n
                End If
            End If
        End If
    Next n
End Sub
Public Function Assign(ByVal previousToken As String) As Long
    Dim n As Long, guid As DemoGuid, buffer As String
    ExpireEmptySessions
    If Len(previousToken) > 0 Then
        For n = 1 To 8
            If Token(n) = previousToken Then
                Storage.Cells(n + 1, 29).value = Now
                Assign = n
                Exit Function
            End If
        Next n
    End If
    For n = 1 To 8
        If Len(Token(n)) = 0 Then
            buffer = String$(39, vbNullChar)
            If CoCreateGuid(guid) <> 0 Then Err.Raise vbObjectError + 2700, "TableOrderDemo", "Session ID generation failed."
            If StringFromGUID2(guid, StrPtr(buffer), 39) = 0 Then Err.Raise vbObjectError + 2700, "TableOrderDemo", "Session ID conversion failed."
            Storage.Cells(n + 1, 24).Value2 = Mid$(buffer, 2, 36)
            SetState n, 1
            Storage.Cells(n + 1, 28).Value2 = 2
            Storage.Cells(n + 1, 29).value = Now
            Assign = n
            Exit Function
        End If
    Next n
    Err.Raise vbObjectError + 2701, "TableOrderDemo", "No available table."
End Function
Public Sub Require(ByVal tableNo As Long, ByVal sessionId As String)
    If tableNo < 1 Or tableNo > 8 Then Err.Raise vbObjectError + 2702, "TableOrderDemo", "Invalid table number."
    If Len(sessionId) = 0 Then Err.Raise vbObjectError + 2703, "TableOrderDemoSession", "Order session expired. Reload the page."
    If Token(tableNo) <> sessionId Then Err.Raise vbObjectError + 2703, "TableOrderDemoSession", "Order session expired. Reload the page."
    Dim seen As Variant
    seen = Storage.Cells(tableNo + 1, 29).value
    If IsDate(seen) Then
        If DateDiff("n", CDate(seen), Now) >= 30 Then
            If TableOrderDemoData.Total(tableNo, sessionId) = 0 Then Release tableNo
            Err.Raise vbObjectError + 2703, "TableOrderDemoSession", "Order session expired. Reload the page."
        End If
    End If
    Storage.Cells(tableNo + 1, 29).value = Now
End Sub
Public Sub Move(ByVal oldNo As Long, ByVal newNo As Long, ByVal sessionId As String)
    Require oldNo, sessionId
    If newNo < 1 Or newNo > 8 Then Err.Raise vbObjectError + 2702, "TableOrderDemo", "Invalid table number."
    If oldNo = newNo Then Exit Sub
    If Len(Token(newNo)) > 0 Then Err.Raise vbObjectError + 2704, "TableOrderDemo", "Selected table is occupied."
    If TableOrderDemoData.Total(oldNo, sessionId) > 0 Then Err.Raise vbObjectError + 2705, "TableOrderDemo", "Table changes are unavailable after submitting an order."
    Storage.Range(Storage.Cells(newNo + 1, 24), Storage.Cells(newNo + 1, 29)).Value2 = _
        Storage.Range(Storage.Cells(oldNo + 1, 24), Storage.Cells(oldNo + 1, 29)).Value2
    Release oldNo
End Sub


