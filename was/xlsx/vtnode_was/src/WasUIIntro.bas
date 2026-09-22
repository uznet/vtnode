Attribute VB_Name = "WasUIIntro"
'Attribute VB_Name = "WasUIIntro"
'Attribute VB_Name = "WasUIIntro"
Option Explicit

Private Const MODULE_NAME As String = "WasUIIntro"
Private Const MAIN_SHEET As String = "VTNode WAS"
Private Const ACCESS_URL_CELL As String = "D14"
Private Const QR_FILE_NAME As String = "vtnode-was-intro-qrcode.bmp"
Private Const INTRO_PAGE As String = "what-is-vtnode-was.html"

Public Sub ShowIntro()
    Dim accessUrl As String

    On Error GoTo ShowError

    WasAudio.Play AUDIO_TICK
    
    accessUrl = GetAccessUrl()
    Load frmWasIntro
    frmWasIntro.LoadIntro accessUrl
    frmWasIntro.Show vbModeless
    Exit Sub

ShowError:
    MsgBox _
        "The VTNode WAS introduction could not be opened." & vbCrLf & vbCrLf & _
        Err.description, _
        vbExclamation, _
        "VTNode"

    Debug.Print "[" & MODULE_NAME & ".ShowIntro] " & _
                Err.Number & " : " & Err.description
End Sub

Public Function GetAccessUrl() As String
    Dim accessUrl As String

    accessUrl = Trim$(CStr( _
        ThisWorkbook.Worksheets(MAIN_SHEET) _
            .Range(ACCESS_URL_CELL).Value2))

    If Len(accessUrl) = 0 Or accessUrl = "-" Then Exit Function

    If LCase$(Left$(accessUrl, 7)) <> "http://" And _
       LCase$(Left$(accessUrl, 8)) <> "https://" Then
        accessUrl = "http://" & accessUrl
    End If

    If Right$(accessUrl, 1) <> "/" Then
        accessUrl = accessUrl & "/"
    End If

    accessUrl = accessUrl & INTRO_PAGE
    GetAccessUrl = accessUrl
End Function

Public Sub OpenInBrowser(ByVal accessUrl As String)
    On Error GoTo BrowserError

    accessUrl = Trim$(accessUrl)
    If Len(accessUrl) = 0 Or accessUrl = "-" Then
        MsgBox _
            "Start the server before opening the introduction page.", _
            vbInformation, _
            "VTNode"
        Exit Sub
    End If

    ThisWorkbook.FollowHyperlink _
        Address:=accessUrl, _
        NewWindow:=True
    Exit Sub

BrowserError:
    MsgBox _
        "The browser could not be opened." & vbCrLf & _
        accessUrl & " -> " & Err.description, _
        vbExclamation, _
        "VTNode"
End Sub

Public Function CreateQRCodeImage(ByVal qrText As String) As String
    Dim savePath As String
    Dim result As Boolean

    On Error GoTo QRCodeError

    qrText = Trim$(qrText)
    If Len(qrText) = 0 Or qrText = "-" Then Exit Function

    savePath = Environ$("TEMP") & Application.PathSeparator & QR_FILE_NAME

    On Error Resume Next
    Kill savePath
    On Error GoTo QRCodeError

    result = VtGenerateQRCode(qrText, 6, savePath)
    If result Then CreateQRCodeImage = savePath
    Exit Function

QRCodeError:
    Debug.Print "[" & MODULE_NAME & ".CreateQRCodeImage] " & _
                Err.Number & " : " & Err.description
End Function


