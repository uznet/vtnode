Attribute VB_Name = "ExportAll"

Option Explicit

Public Sub ExportAllComponents()
    Dim comp As VBComponent
    Dim exportPath As String
    Dim extension As String
    
    ' 내보낼 경로 설정 (현재 엑셀 파일과 같은 경로의 'src' 폴더)
    'exportPath = "D:\job\xlibw\xhttp\HTPackages\source\httpa\vba\" 'ThisWorkbook.Path & "\src\"
    exportPath = ThisWorkbook.path & "\src\"
    
    ' 폴더가 없으면 생성
    If dir(exportPath, vbDirectory) = "" Then MkDir exportPath
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        ' 파일 형식에 따른 확장자 결정
        Select Case comp.Type
            Case vbext_ct_ClassModule: extension = ".cls"
            Case vbext_ct_StdModule:   extension = ".bas"
            Case vbext_ct_MSForm:      extension = ".frm"
            Case Else: extension = ""
        End Select
        
        ' 실제 내보내기 수행
        If extension <> "" Then
            comp.Export exportPath & comp.Name & extension
            Debug.Print "Exported: " & comp.Name & extension
        End If
    Next comp
    
    MsgBox "모든 모듈/클래스가 " & exportPath & " 폴더로 내보내졌습니다.", vbInformation
End Sub
