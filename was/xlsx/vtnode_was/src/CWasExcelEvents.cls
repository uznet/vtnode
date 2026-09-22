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
    ByVal Wb As Excel.Workbook, Cancel As Boolean)

    Debug.Print "App_WorkbookBeforeClose.1"
    ' WAS와 런타임 통합 문서를 보호
    If Not (Wb Is ThisWorkbook) Then
        Debug.Print "App_WorkbookBeforeClose.2->" & Wb.Name
        
        If LCase$(Wb.Name) <> "vtnode_vba_rt64.xlsm" Then
           Exit Sub
        End If
        
    End If

    Debug.Print "App_WorkbookBeforeClose.3:status=" & WasHttpaService.GetStatusDescription
    

    Select Case WasHttpaService.HttpaStatus
        Case WAS_HTTPA_SERVICE_UNINITIALIZED, _
             WAS_HTTPA_SERVICE_STOPPED
            ' 종료된 상태: 닫기 허용

        Case Else
            Cancel = True

            MsgBox "먼저 STOP SERVER를 실행하세요." & vbCrLf & _
                   "서버 종료가 완료된 후 Excel을 닫아주세요.", _
                   vbExclamation, "VTNode WAS"
    End Select
End Sub

