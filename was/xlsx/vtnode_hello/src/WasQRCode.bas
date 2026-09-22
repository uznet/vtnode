Attribute VB_Name = "WasQRCode"
'Attribute VB_Name = "WasQRCode"
Option Explicit

Private Const RUNTIME_SHEET_NAME As String = "VTNode Hello"
Private Const HTTP_URL_CELL As String = "D27"
Private Const QRCODE_CELL As String = "H15"
Private Const QRCODE_SHAPE_NAME As String = "VTNodeQRCode"
Private Const QRCODE_PADDING As Double = 4
Private mQRCodeText As String


Private Function GenerateQRCode( _
    ByVal text As String) As String

    Dim result As Boolean
    Dim savePath As String

    text = Trim$(text)

    If Len(text) = 0 Then Exit Function
    If text = "-" Then Exit Function

    savePath = VtPathUtils.VtPathCombine(WasMain.WasHomeDirectory(), "qrcode.bmp")

    result = VtGenerateQRCode( _
                 text, _
                 6, _
                 savePath)

    If result Then
        GenerateQRCode = savePath
    Else
        GenerateQRCode = vbNullString
    End If

End Function

Private Sub ClearPictureArea( _
    ByVal targetCellAddress As String)

    Const SHEET_NAME As String = RUNTIME_SHEET_NAME

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim picture As Shape
    Dim index As Long

    Dim targetLeft As Double
    Dim targetTop As Double
    Dim targetRight As Double
    Dim targetBottom As Double

    On Error GoTo EH

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set targetRange = ws.Range(targetCellAddress)

    '
    ' Expand L19 to the complete merged area L19:O23.
    '
    If targetRange.MergeCells Then
        Set targetRange = targetRange.MergeArea
    End If

    targetLeft = targetRange.Left
    targetTop = targetRange.Top
    targetRight = targetRange.Left + targetRange.Width
    targetBottom = targetRange.Top + targetRange.Height

    '
    ' Iterate backwards because shapes can be deleted.
    '
    For index = ws.Shapes.Count To 1 Step -1

        Set picture = ws.Shapes(index)

        '
        ' Do not delete buttons or other UI shapes.
        '
        If picture.Type = msoPicture _
           Or picture.Type = msoLinkedPicture Then

            If ShapeOverlapsArea( _
                   picture, _
                   targetLeft, _
                   targetTop, _
                   targetRight, _
                   targetBottom) Then

                picture.Delete

            End If

        End If

    Next index

    Exit Sub

EH:
    Debug.Print "[ClearPictureArea] " _
              & Err.Number & " : " _
              & Err.description

End Sub
Private Function ShapeOverlapsArea( _
    ByVal item As Shape, _
    ByVal areaLeft As Double, _
    ByVal areaTop As Double, _
    ByVal areaRight As Double, _
    ByVal areaBottom As Double) As Boolean

    Dim shapeRight As Double
    Dim shapeBottom As Double

    shapeRight = item.Left + item.Width
    shapeBottom = item.Top + item.Height

    ShapeOverlapsArea = _
        item.Left < areaRight _
        And shapeRight > areaLeft _
        And item.Top < areaBottom _
        And shapeBottom > areaTop

End Function
Private Sub SetPicture( _
    ByVal targetCellAddress As String, _
    ByVal imagePath As String)

    Const SHEET_NAME As String = RUNTIME_SHEET_NAME
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim picture As Shape

    On Error GoTo EH

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set targetRange = ws.Range(targetCellAddress)

    If targetRange.MergeCells Then
        Set targetRange = targetRange.MergeArea
    End If

    Set picture = ws.Shapes.AddPicture( _
                      fileName:=imagePath, _
                      LinkToFile:=msoFalse, _
                      SaveWithDocument:=msoTrue, _
                      Left:=targetRange.Left + QRCODE_PADDING, _
                      Top:=targetRange.Top + QRCODE_PADDING, _
                      Width:=-1, _
                      Height:=-1)

    picture.Name = QRCODE_SHAPE_NAME
    FitPictureToRange picture, targetRange, QRCODE_PADDING

    Exit Sub

EH:
    Err.Raise Err.Number, _
              "SetPicture", _
              Err.description

End Sub
Private Sub FitPictureToRange( _
    ByVal picture As Shape, _
    ByVal targetRange As Range, _
    ByVal padding As Double)

    Dim pictureSize As Double

    pictureSize = targetRange.Width - padding * 2

    If pictureSize > targetRange.Height - padding * 2 Then
        pictureSize = targetRange.Height - padding * 2
    End If

    If pictureSize <= 0 Then Exit Sub

    '
    ' A QR code must remain square. Fill the largest square that fits in the
    ' complete L19 merged area, then center it horizontally and vertically.
    '
    picture.LockAspectRatio = msoFalse
    picture.Width = pictureSize
    picture.Height = pictureSize

    picture.Left = targetRange.Left _
                 + (targetRange.Width - pictureSize) / 2

    picture.Top = targetRange.Top _
                + (targetRange.Height - pictureSize) / 2

    picture.Placement = xlMoveAndSize

End Sub
Public Sub SetQRCode( _
    ByVal IsRunning As Boolean, _
    ByVal serviceUrl As String)

    Dim savePath As String
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim picture As Shape

    serviceUrl = Trim$(serviceUrl)

    '
    ' Server stopped or URL is unavailable.
    '
    If Not IsRunning _
       Or Len(serviceUrl) = 0 _
       Or serviceUrl = "-" Then

        ClearPictureArea QRCODE_CELL
        mQRCodeText = vbNullString
        Exit Sub

    End If

    '
    ' Do not regenerate when the URL has not changed.
    '
    If StrComp( _
           mQRCodeText, _
           serviceUrl, _
           vbBinaryCompare) = 0 Then

        On Error Resume Next

        Set ws = ThisWorkbook.Worksheets(RUNTIME_SHEET_NAME)
        Set targetRange = ws.Range(QRCODE_CELL)

        If targetRange.MergeCells Then
            Set targetRange = targetRange.MergeArea
        End If

        Set picture = ws.Shapes(QRCODE_SHAPE_NAME)

        On Error GoTo 0

        If Not picture Is Nothing Then
            FitPictureToRange picture, targetRange, QRCODE_PADDING
            Exit Sub
        End If
    End If

    savePath = GenerateQRCode(serviceUrl)

    If Len(savePath) = 0 Then
        ClearPictureArea QRCODE_CELL
        mQRCodeText = vbNullString
        Exit Sub
    End If

    '
    ' Remove the previous QR image before inserting the new one.
    '
    ClearPictureArea QRCODE_CELL
    SetPicture QRCODE_CELL, savePath

    mQRCodeText = serviceUrl

End Sub
'Public Sub UpdateQRCodeFromSheet()
'
'    Dim ws As Worksheet
'    Dim serviceUrl As String
'
'    On Error GoTo EH
'
'    Set ws = ThisWorkbook.Worksheets(RUNTIME_SHEET_NAME)
'
'    serviceUrl = Trim$(CStr( _
'                     ws.Range(HTTP_URL_CELL).Value2))
'
'    SetQRCode _
'        IsRunning:=(Len(serviceUrl) > 0 And serviceUrl <> "-"), _
'        serviceUrl:=serviceUrl
'
'    Exit Sub
'
'EH:
'    Debug.Print "[UpdateQRCodeFromSheet] " _
'              & Err.Number & " : " _
'              & Err.description
'
'End Sub
'
'
'
