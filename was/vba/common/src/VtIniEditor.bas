Attribute VB_Name = "VtIniEditor"
Option Explicit

#If VBA7 Then
    ' Windows API: 외부 프로세스 실행 및 파일 오픈 함수
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long _
    ) As LongPtr
#End If

' 윈도우 표시 상수 (정상 크기로 활성화)
Private Const SW_SHOWNORMAL As Long = 1

''' <summary>
''' 지정된 .ini 파일을 시스템 기본 편집기(메모장 등)로 오픈합니다.
''' </summary>
Public Sub OpenIniEditor(ByVal iniFilePath As String)
    ' 1. 파일 경로 유효성 검증
    If iniFilePath = "" Or dir(iniFilePath) = "" Then
        MsgBox "지정된 설정 파일이 존재하지 않거나 경로가 올바르지 않습니다.", vbCritical, "VTNode WAS"
        Exit Sub
    End If
    
    ' 2. "open" 커맨드를 통해 시스템에 등록된 텍스트 편집기로 .ini 실행
    ' (메모장을 하드코딩하지 않아도 OS가 .ini 확장자 연결 프로그램을 찾아 실행합니다)
    Call ShellExecute(0, "open", iniFilePath, vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub

