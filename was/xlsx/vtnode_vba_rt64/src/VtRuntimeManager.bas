Attribute VB_Name = "VtRuntimeManager"
Option Explicit

'===============================================================================
' Module Name : VtRuntimeManager
'
' Purpose
'   - Manage relationship between VTNode WAS application workbooks
'     and vtnode_rt64.xlsm
'   - Register / unregister runtime clients
'   - Hide / show VTNode Runtime worksheet
'   - Show / hide VTNode Runtime UserForm
'
' Policy
'   - vtnode_rt64.xlsm is NOT automatically closed.
'   - While at least one client is attached, Runtime worksheet is hidden.
'   - When the last client is detached, Runtime worksheet is shown.
'   - HTTP Listener ownership / channel allocation are NOT managed here.
'===============================================================================


'-------------------------------------------------------------------------------
' Configuration
'-------------------------------------------------------------------------------

Private Const VTNODE_RUNTIME_WORKBOOK As String = "vtnode_rt64.xlsm"
Private Const VTNODE_RUNTIME_SHEET As String = "Runtime"

' UserForm name inside vtnode_rt64.xlsm
Private Const VTNODE_RUNTIME_FORM As String = "VtRuntimeConsole"


'-------------------------------------------------------------------------------
' Module State
'-------------------------------------------------------------------------------

Private mClients As Object       ' Scripting.Dictionary
Private mInitialized As Boolean


'===============================================================================
' Initialization
'===============================================================================

Private Sub EnsureInitialized()

    If mInitialized Then Exit Sub

    Set mClients = CreateObject("Scripting.Dictionary")

    mClients.CompareMode = vbTextCompare

    mInitialized = True

End Sub


'===============================================================================
' Public API
'===============================================================================

'
' Attach a VTNode WAS application to Runtime Manager.
'
' Example:
'
'   Call VtRuntimeAttach("get-started")
'
Public Function VtRuntimeAttach(ByVal clientId As String) As Boolean

    On Error GoTo EH

    EnsureInitialized

    clientId = Trim$(clientId)

    Debug.Print "VtRuntimeAttach " & clientId
    
    If Len(clientId) = 0 Then
        Err.Raise vbObjectError + 1000, _
                  "VtRuntimeAttach", _
                  "clientId is empty."
    End If


    ' Do not register the same client twice.
    If Not mClients.Exists(clientId) Then

        mClients.Add clientId, True

    End If


    '
    ' While one or more clients are using VTNode Runtime,
    ' hide the Runtime worksheet.
    '
    If mClients.count > 0 Then

        Call VtRuntimeHideWorkbookWindow

    End If


    VtRuntimeAttach = True

    Exit Function


EH:

    Debug.Print "[VtRuntimeAttach] " & Err.Number & " : " & Err.Description

    VtRuntimeAttach = False

End Function


'
' Detach a VTNode WAS application.
'
' When the final client is detached,
' show the Runtime worksheet again.
'
Public Sub VtRuntimeDetach(ByVal clientId As String)

    On Error GoTo EH

    EnsureInitialized

    clientId = Trim$(clientId)

    If Len(clientId) = 0 Then Exit Sub


    If mClients.Exists(clientId) Then

        mClients.Remove clientId

    End If


    '
    ' No applications are currently attached.
    '
    If mClients.count = 0 Then

        'Call VtRuntimeShowWorksheet
        Call VtRuntimeScheduleShowWorkbookWindow

    End If

    Exit Sub


EH:

    Debug.Print "[VtRuntimeDetach] " & Err.Number & " : " & Err.Description

End Sub

Public Sub VtRuntimeScheduleShowWorkbookWindow()

    On Error GoTo EH

    Application.OnTime _
        EarliestTime:=Now + TimeSerial(0, 0, 1), _
        Procedure:="'" & VTNODE_RUNTIME_WORKBOOK & _
                   "'!VtRuntimeShowWorkbookWindow"

    Exit Sub

EH:
    Debug.Print "[VtRuntimeScheduleShowWorkbookWindow] " & _
                Err.Number & " : " & Err.Description

End Sub
'
' Returns number of currently attached clients.
'
Public Function VtRuntimeClientCount() As Long

    EnsureInitialized

    VtRuntimeClientCount = mClients.count

End Function


'
' Returns True when client is already attached.
'
Public Function VtRuntimeIsAttached(ByVal clientId As String) As Boolean

    EnsureInitialized

    clientId = Trim$(clientId)

    If Len(clientId) = 0 Then Exit Function

    VtRuntimeIsAttached = mClients.Exists(clientId)

End Function


'===============================================================================
' Runtime Worksheet
'===============================================================================

Public Sub VtRuntimeHideWorkbookWindow()

    Dim wb As Workbook

    On Error GoTo EH

    Set wb = GetRuntimeWorkbook()
    If wb Is Nothing Then Exit Sub

    If wb.Windows.count > 0 Then
        wb.Windows(1).Visible = False
    End If

    Exit Sub

EH:
    Debug.Print "[VtRuntimeHideWorksheet] " & _
                Err.Number & " : " & Err.Description

End Sub


Public Sub VtRuntimeShowWorkbookWindow()

    Dim wb As Workbook

    On Error GoTo EH

    Set wb = GetRuntimeWorkbook()
    If wb Is Nothing Then Exit Sub

    If wb.Windows.count > 0 Then
        wb.Windows(1).Visible = True
        wb.Activate
    End If

    Exit Sub

EH:
    Debug.Print "[VtRuntimeShowWorksheet] " & _
                Err.Number & " : " & Err.Description

End Sub

'===============================================================================
' Runtime UserForm
'===============================================================================

Public Sub VtRuntimeShow()

    Dim wb As Workbook

    On Error GoTo EH


    Set wb = GetRuntimeWorkbook()

    If wb Is Nothing Then Exit Sub


    '
    ' Run public procedure inside vtnode_rt64.xlsm.
    '
    Application.Run _
        "'" & wb.Name & "'!VtRuntimeUIShow"

    Exit Sub


EH:

    Debug.Print "[VtRuntimeShow] " & Err.Number & " : " & Err.Description

End Sub


Public Sub VtRuntimeHide()

    Dim wb As Workbook

    On Error GoTo EH


    Set wb = GetRuntimeWorkbook()

    If wb Is Nothing Then Exit Sub


    Application.Run _
        "'" & wb.Name & "'!VtRuntimeUIHide"

    Exit Sub


EH:

    Debug.Print "[VtRuntimeHide] " & Err.Number & " : " & Err.Description

End Sub


'===============================================================================
' Runtime Workbook
'===============================================================================

Public Function VtRuntimeIsLoaded() As Boolean

    Dim wb As Workbook

    Set wb = GetRuntimeWorkbook()

    VtRuntimeIsLoaded = Not (wb Is Nothing)

End Function


Public Function VtRuntimeWorkbook() As Workbook

    Set VtRuntimeWorkbook = GetRuntimeWorkbook()

End Function


Private Function GetRuntimeWorkbook() As Workbook

    Dim wb As Workbook

    On Error Resume Next

    Set wb = Application.Workbooks(VTNODE_RUNTIME_WORKBOOK)

    On Error GoTo 0


    Set GetRuntimeWorkbook = wb

End Function


Private Function GetRuntimeWorksheet( _
    ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = wb.Worksheets(VTNODE_RUNTIME_SHEET)

    On Error GoTo 0


    Set GetRuntimeWorksheet = ws

End Function


'===============================================================================
' Utilities
'===============================================================================

Private Function VisibleWorksheetCount(ByVal wb As Workbook) As Long

    Dim ws As Worksheet
    Dim count As Long


    For Each ws In wb.Worksheets

        If ws.Visible = xlSheetVisible Then

            count = count + 1

        End If

    Next


    VisibleWorksheetCount = count

End Function


'===============================================================================
' Debug / Diagnostic
'===============================================================================

Public Sub VtRuntimeDumpClients()

    Dim key As Variant

    EnsureInitialized


    Debug.Print "----------------------------------------"
    Debug.Print "VTNode Runtime Clients"
    Debug.Print "Count = " & mClients.count
    Debug.Print "----------------------------------------"


    For Each key In mClients.keys

        Debug.Print CStr(key)

    Next


    Debug.Print "----------------------------------------"

End Sub

