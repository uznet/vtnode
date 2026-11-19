Attribute VB_Name = "TableOrderDemo"
Option Explicit
Private mAccessUrl As String
Public Sub TableOrderDemo_Activate()
    ThisWorkbook.Worksheets("TableOrder").Activate
End Sub
Public Sub Init()
    Dim n As Long, ws As Worksheet
    Set ws = TableOrderDemoSession.Storage
    For n = 1 To 8
        ws.Cells(n + 1, 23).Value2 = n
    Next n
    ws.Columns("W:AC").Hidden = True
    TableOrderDemoSession.ExpireEmptySessions
    TableOrderDemoPanel.Refresh
    RefreshAccess
End Sub
Public Sub TableOrderDemo_SelectTable()
    TableOrderDemoPanel.SelectTable
End Sub
Public Sub TableOrderDemo_ShowTable(ByVal tableNo As Long)
    TableOrderDemoPanel.TableOrderDemo_Show tableNo
End Sub
Public Sub Uninit()
    ' Preserve unpaid orders and table sessions across service stop/start.
    RefreshAccess
End Sub
Public Sub RefreshAccess()
    Dim ws As Worksheet, target As Range, url As String, picture As Shape
    Dim path As String, side As Double, n As Long, previousEvents As Boolean
    Dim running As Boolean, havePicture As Boolean
    previousEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set ws = TableOrderDemoSession.Storage
    running = VtHttpaEngine.VtHttpaIsRun
    ws.Range("S2").Value2 = IIf(running, "RUNNING", "STOPPED")
    ws.Range("S2").Font.Color = IIf(running, RGB(19, 115, 72), RGB(100, 116, 139))
    If running Then
        url = WasHttpaRuntimeStates.WebsocketDemoAccessUrl()
        url = Replace(url, "/websocket.html", "/order")
    End If
    On Error Resume Next
    Set picture = ws.Shapes("TableOrderDemoQRCode")
    havePicture = Not picture Is Nothing
    On Error GoTo Failed
    If url = mAccessUrl And havePicture And Len(url) > 0 Then GoTo Done
    Set target = ws.Range("B11").MergeArea
    For n = ws.Shapes.count To 1 Step -1
        If ws.Shapes(n).Name = "TableOrderDemoQRCode" Then ws.Shapes(n).Delete
    Next n
    ws.Range("B7").MergeArea.Hyperlinks.Delete
    ws.Range("B7").Value2 = "-"
    ws.Range("B11").Value2 = "서버를 시작하면 QR 코드가 표시됩니다."
    mAccessUrl = vbNullString
    If Len(url) = 0 Then GoTo Done
    ws.Hyperlinks.Add ws.Range("B7"), url, , , url
    target.ClearContents
    path = Environ$("TEMP") & "\vtwas-tableorder-" & Format$(Now, "yyyymmddhhnnss") & ".bmp"
    If Not VtGenerateQRCode(url, 6, path) Then Err.Raise vbObjectError + 2720, "TableOrderDemo", "QR generation failed."
    side = target.width - 8
    If side > target.height - 8 Then side = target.height - 8
    Set picture = ws.Shapes.AddPicture(path, msoFalse, msoTrue, target.left + (target.width - side) / 2, target.top + (target.height - side) / 2, side, side)
    picture.Name = "TableOrderDemoQRCode"
    picture.AlternativeText = url
    picture.Placement = xlMoveAndSize
    mAccessUrl = url
    GoTo Done
Failed:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Range("B11").Value2 = "QR 갱신 실패 · Access URL을 사용하세요."
    Debug.Print "[TableOrderDemo.RefreshAccess] " & Err.description
Done:
    On Error Resume Next
    If Len(path) > 0 Then
        If Len(dir$(path)) > 0 Then Kill path
    End If
    Application.EnableEvents = previousEvents
End Sub

