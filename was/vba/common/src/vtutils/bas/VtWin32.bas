Attribute VB_Name = "VtWin32"
Option Explicit

#If VBA7 Then

'Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'    (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As Long)


Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As LongPtr, _
        ByVal Source As LongPtr, _
        ByVal Length As LongPtr) ' 64비트에서는 길이를 나타내는 size_t가 8바이트일 수 있어 LongPtr 권장
        
        
' 현재 디렉토리를 변경하는 Windows API 선언
Public Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long


Public Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type



' 윈도우 셸 인터페이스 익스포트 함수 바인딩
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr

' 윈도우 출력 상태 상수
Public Const SW_SHOWNORMAL As Long = 1


Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
Public Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
    
#End If

