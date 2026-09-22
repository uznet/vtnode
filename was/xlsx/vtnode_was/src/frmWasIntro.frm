VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWasIntro 
   Caption         =   "What is VTNode WAS?"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8955.001
   OleObjectBlob   =   "frmWasIntro.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmWasIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mAccessUrl As String

Private Sub UserForm_Initialize()
    lblStatus.caption = vbNullString
    lblQrPlaceholder.Visible = True
    cmdOpenBrowser.enabled = False
    cmdCopyUrl.enabled = False
End Sub

Public Sub LoadIntro(ByVal accessUrl As String)
    Dim qrImagePath As String

    On Error GoTo LoadError

    mAccessUrl = Trim$(accessUrl)
    txtAccessUrl.text = mAccessUrl
    lblStatus.caption = vbNullString

    If Len(mAccessUrl) = 0 Or mAccessUrl = "-" Then
        txtAccessUrl.text = "Start the server to create an access URL."
        lblStatus.caption = "The VTNode WAS server is not running."
        lblQrPlaceholder.caption = "NO URL"
        lblQrPlaceholder.Visible = True
        cmdOpenBrowser.enabled = False
        cmdCopyUrl.enabled = False
        Exit Sub
    End If

    cmdOpenBrowser.enabled = True
    cmdCopyUrl.enabled = True

    qrImagePath = WasUIIntro.CreateQRCodeImage(mAccessUrl)
    If Len(qrImagePath) > 0 Then
        Set imgQRCode.picture = LoadPicture(qrImagePath)
        lblQrPlaceholder.Visible = False
    Else
        lblQrPlaceholder.caption = "QR UNAVAILABLE"
        lblQrPlaceholder.Visible = True
        lblStatus.caption = "The QR code could not be generated."
    End If

    Exit Sub

LoadError:
    lblStatus.caption = "The introduction screen could not be loaded."
    Debug.Print "[frmWasIntro.LoadIntro] " & _
                Err.Number & " : " & Err.description
End Sub

Private Sub cmdOpenBrowser_Click()
    WasUIIntro.OpenInBrowser mAccessUrl
End Sub

Private Sub cmdCopyUrl_Click()
    Dim clipboard As MSForms.DataObject

    On Error GoTo CopyError

    Set clipboard = New MSForms.DataObject
    clipboard.SetText mAccessUrl
    clipboard.PutInClipboard
    lblStatus.caption = "Access URL copied."
    Exit Sub

CopyError:
    lblStatus.caption = "The access URL could not be copied."
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

