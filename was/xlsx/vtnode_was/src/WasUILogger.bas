Attribute VB_Name = "WasUILogger"
'Attribute VB_Name = "WasSheetLog"
Option Explicit

Private Const LOG_SHEET_NAME As String = "VTNode WAS"
Private Const LOG_FIRST_ROW As Long = 19
Private Const LOG_LAST_ROW As Long = 25

Private Const LOG_TIME_COLUMN As String = "B"
Private Const LOG_METHOD_COLUMN As String = "C"
Private Const LOG_PATH_COLUMN As String = "D"
Private Const LOG_STATUS_COLUMN As String = "G"
Private Const LOG_DURATION_COLUMN As String = "I"
Private Const LOG_DETAIL_COLUMN As String = "K"

Private mNextLogRow As Long
Private mLogRows As Object
Private mRowKeys As Object

Public Sub InitLog()

    Dim ws As Worksheet
    Dim logRow As Long

    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)

    For logRow = LOG_FIRST_ROW To LOG_LAST_ROW
        ClearLogRow ws, logRow
    Next logRow

    Set mLogRows = Nothing
    Set mRowKeys = Nothing
    mNextLogRow = LOG_FIRST_ROW

End Sub

Public Sub OnBeginLog(ByVal timestamp As String)

    Dim ws As Worksheet
    Dim logRow As Long

    timestamp = Trim$(timestamp)
    If Len(timestamp) = 0 Then Exit Sub

    EnsureLogIndexes

    If mLogRows.Exists(timestamp) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    logRow = TakeNextLogRow()

    RemoveRowIndex logRow
    ClearLogRow ws, logRow

    mLogRows.Add timestamp, logRow
    mRowKeys.Add CStr(logRow), timestamp

    ws.Range(LOG_TIME_COLUMN & logRow).Value2 = timestamp

End Sub

Public Sub OnRequestLog( _
    ByVal timestamp As String, _
    ByVal socket As CVtHttpaSocket, _
    ByVal request As CVtHttpaRequest)

    Dim ws As Worksheet
    Dim logRow As Long

    If request Is Nothing Then Exit Sub

    
    logRow = FindLogRow(timestamp)
    If logRow = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)

    
    ws.Range(LOG_METHOD_COLUMN & logRow).Value2 = VTHttpaMethodToString(request.method)
    
'        ReadFirstMember( _
'            request, _
'            Array("MethodString", "Method", "RequestMethod", "HttpMethod"))

    
    ws.Range(LOG_PATH_COLUMN & logRow).Value2 = socket.RemoteAddr & "->" & request.Uri
    
'        ReadFirstMember( _
'            request, _
'            Array("Path", "RequestPath", "RequestTarget", "Target", "Url", "Uri"))

End Sub

Public Sub OnResponseLog( _
    ByVal timestamp As String, _
    ByVal response As CVtHttpaResponse, _
    ByVal detail As String)

    Dim ws As Worksheet
    Dim logRow As Long
    Dim status As String
    Dim reason As String

    logRow = FindLogRow(timestamp)
    If logRow = 0 Then Exit Sub

    If Not response Is Nothing Then
    
       status = CStr(response.StatusCode)
       
    
'        status = ReadFirstMember( _
'                     response, _
'                     Array("StatusCode", "Status", "ResponseStatus"))
'
'        reason = ReadFirstMember( _
'                     response, _
'                     Array("ReasonPhrase", "Reason", "StatusText"))
'
'        If Len(reason) > 0 Then
'            If InStr(1, status, reason, vbTextCompare) = 0 Then
'                status = Trim$(status & " " & reason)
'            End If
'        End If
    End If

    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    ws.Range(LOG_STATUS_COLUMN & logRow).Value2 = status
    ws.Range(LOG_DETAIL_COLUMN & logRow).Value2 = detail

End Sub

Public Sub OnEndLog( _
    ByVal timestamp As String, _
    ByVal duration As Long)
    

    Dim ws As Worksheet
    Dim logRow As Long

    logRow = FindLogRow(timestamp)
    If logRow = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    ws.Range(LOG_DURATION_COLUMN & logRow).Value2 = duration & " ms"

End Sub

Private Function FindLogRow( _
    ByVal timestamp As String) As Long

    timestamp = Trim$(timestamp)
    If Len(timestamp) = 0 Then Exit Function

    EnsureLogIndexes

    If mLogRows.Exists(timestamp) Then
        FindLogRow = CLng(mLogRows(timestamp))
    End If

End Function

Private Sub EnsureLogIndexes()

    If mLogRows Is Nothing Then
        Set mLogRows = CreateObject("Scripting.Dictionary")
        mLogRows.CompareMode = vbBinaryCompare
    End If

    If mRowKeys Is Nothing Then
        Set mRowKeys = CreateObject("Scripting.Dictionary")
        mRowKeys.CompareMode = vbBinaryCompare
    End If

    If mNextLogRow < LOG_FIRST_ROW _
       Or mNextLogRow > LOG_LAST_ROW Then

        mNextLogRow = LOG_FIRST_ROW
    End If

End Sub

Private Function TakeNextLogRow() As Long

    EnsureLogIndexes

    TakeNextLogRow = mNextLogRow
    mNextLogRow = mNextLogRow + 1

    If mNextLogRow > LOG_LAST_ROW Then
        mNextLogRow = LOG_FIRST_ROW
    End If

End Function

Private Sub RemoveRowIndex(ByVal logRow As Long)

    Dim rowKey As String
    Dim timestamp As String

    rowKey = CStr(logRow)

    If mRowKeys.Exists(rowKey) Then
        timestamp = CStr(mRowKeys(rowKey))
        mRowKeys.Remove rowKey

        If mLogRows.Exists(timestamp) Then
            mLogRows.Remove timestamp
        End If
    End If

End Sub

Private Sub ClearLogRow( _
    ByVal ws As Worksheet, _
    ByVal logRow As Long)

    ClearLogCell ws.Range(LOG_TIME_COLUMN & logRow)
    ClearLogCell ws.Range(LOG_METHOD_COLUMN & logRow)
    ClearLogCell ws.Range(LOG_PATH_COLUMN & logRow)
    ClearLogCell ws.Range(LOG_STATUS_COLUMN & logRow)
    ClearLogCell ws.Range(LOG_DURATION_COLUMN & logRow)
    ClearLogCell ws.Range(LOG_DETAIL_COLUMN & logRow)

End Sub

Private Sub ClearLogCell(ByVal targetCell As Range)

    If targetCell.MergeCells Then
        targetCell.MergeArea.ClearContents
    Else
        targetCell.ClearContents
    End If

End Sub

Private Function ReadFirstMember( _
    ByVal source As Object, _
    ByVal memberNames As Variant) As String

    Dim memberName As Variant
    Dim memberValue As Variant

    For Each memberName In memberNames
        If TryReadMember(source, CStr(memberName), memberValue) Then
            ReadFirstMember = CStr(memberValue)
            Exit Function
        End If
    Next memberName

End Function

Private Function TryReadMember( _
    ByVal source As Object, _
    ByVal memberName As String, _
    ByRef memberValue As Variant) As Boolean

    On Error Resume Next

    Err.Clear
    memberValue = CallByName(source, memberName, VbGet)

    If Err.Number = 0 Then
        TryReadMember = True
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    memberValue = CallByName(source, memberName, VbMethod)
    TryReadMember = (Err.Number = 0)

    On Error GoTo 0

End Function


