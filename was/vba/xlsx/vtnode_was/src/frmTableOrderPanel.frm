VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableOrderPanel 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmTableOrderPanel.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmTableOrderPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Version 5#
'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTableOrderPanel
'   caption = "VTNode Table Order · 테이블 현황"
'   ClientHeight = 11220
'   ClientLeft = 120
'   ClientTop = 465
'   ClientWidth = 12165
'   OleObjectBlob   =   "frmTableOrderPanel.frx":0000
'   StartUpPosition = 1    '소유자 가운데
'End
'Attribute VB_Name = "frmTableOrderPanel"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
Option Explicit

Private Const DASHBOARD_SHEET As String = "Sheet1"

Private Sub UserForm_Initialize()
    RefreshFromSheet
End Sub

Public Sub RefreshFromSheet()
'    Dim ws As Worksheet
'    Dim i As Long, sourceColumn As String, sourceRow As Long
'    Dim suffix As String, stateText As String, amountValue As Variant
'
'    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
'    lblActiveValue.caption = CStr(ws.Range("N5").Value2)
'    lblStaffValue.caption = CStr(ws.Range("O5").Value2)
'    lblOrderValue.caption = CStr(ws.Range("P5").Value2)
'    lblSalesValue.caption = Format$(Val(ws.Range("Q5").Value2), "#,##0") & "원"
''    lblSelection.caption = CStr(ws.Range("N30").Value2)
'
'    For i = 1 To 8
'        suffix = Format$(i, "00")
'        If i Mod 2 = 1 Then sourceColumn = "N" Else sourceColumn = "P"
'        sourceRow = 10 + ((i - 1) \ 2) * 5
'        stateText = CStr(ws.Range(sourceColumn & CStr(sourceRow + 1)).Value2)
'        amountValue = ws.Range(sourceColumn & CStr(sourceRow + 2)).Value2
'
'        Me.Controls("fraTable" & suffix).Controls("lblTable" & suffix & "Title").caption = _
'            CStr(ws.Range(sourceColumn & CStr(sourceRow)).Value2)
'        Me.Controls("fraTable" & suffix).Controls("lblTable" & suffix & "State").caption = stateText
'        If IsNumeric(amountValue) And Val(amountValue) > 0 Then
'            Me.Controls("fraTable" & suffix).Controls("lblTable" & suffix & "Amount").caption = _
'                Format$(amountValue, "#,##0") & "원"
'        Else
'            Me.Controls("fraTable" & suffix).Controls("lblTable" & suffix & "Amount").caption = "-"
'        End If
'        Me.Controls("fraTable" & suffix).Controls("cmdTable" & suffix).caption = ButtonCaption(stateText)
'        Me.Controls("fraTable" & suffix).BackColor = StateColor(stateText)
'    Next i
End Sub

Private Function ButtonCaption(ByVal stateText As String) As String
    Select Case stateText
        Case "주문중": ButtonCaption = "주문확인"
        Case "조리중": ButtonCaption = "서빙완료"
        Case "식사중": ButtonCaption = "계 산"
        Case "직원호출": ButtonCaption = "호출해제"
        Case Else: ButtonCaption = "확 인"
    End Select
End Function

Private Function StateColor(ByVal stateText As String) As Long
    Select Case stateText
        Case "주문중": StateColor = RGB(113, 63, 18)
        Case "조리중": StateColor = RGB(140, 122, 186)
        Case "식사중": StateColor = RGB(30, 58, 138)
        Case "직원호출": StateColor = RGB(127, 29, 29)
        Case Else: StateColor = RGB(6, 78, 59)
    End Select
End Function

Private Sub HandleTableClick(ByVal tableNo As Long)
    AppTablePanel.UserFormTableButtonClick tableNo
    RefreshFromSheet
End Sub

Private Sub cmdTable01_Click(): HandleTableClick 1: End Sub
Private Sub cmdTable02_Click(): HandleTableClick 2: End Sub
Private Sub cmdTable03_Click(): HandleTableClick 3: End Sub
Private Sub cmdTable04_Click(): HandleTableClick 4: End Sub
Private Sub cmdTable05_Click(): HandleTableClick 5: End Sub
Private Sub cmdTable06_Click(): HandleTableClick 6: End Sub
Private Sub cmdTable07_Click(): HandleTableClick 7: End Sub
Private Sub cmdTable08_Click(): HandleTableClick 8: End Sub
Private Sub cmdRefresh_Click(): RefreshFromSheet: End Sub
Private Sub cmdClose_Click(): Unload Me: End Sub



