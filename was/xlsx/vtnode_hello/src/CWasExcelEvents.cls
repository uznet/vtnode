VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWasExcelEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Excel.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WorkbookBeforeClose( _
    ByVal wb As Workbook, _
    Cancel As Boolean)

    If wb Is ThisWorkbook Then
        Debug.Print "VTNode Workbook closing"

        ' 이 위치에서 HTTPA 중지 요청
        ' Event Loop 종료 확인 후 WasUninit 호출
    End If

End Sub

