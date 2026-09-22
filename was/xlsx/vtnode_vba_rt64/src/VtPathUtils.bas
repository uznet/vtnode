Attribute VB_Name = "VtPathUtils"
Option Explicit

' =========================
' Win32 API
' =========================
#If VBA7 Then
    ' 64bit + 32bit VBA7 (Office 2010 이상)
    Private Declare PtrSafe Function PathCombine Lib "shlwapi.dll" Alias "PathCombineW" ( _
        ByVal pszDest As LongPtr, _
        ByVal pszDir As LongPtr, _
        ByVal pszFile As LongPtr _
    ) As LongPtr
#End If

' =========================
' Path Combine
' =========================
Public Function VtPathCombine(ByVal dir As String, ByVal file As String) As String
    Dim buffer As String
    buffer = String$(260, vbNullChar)

    Call PathCombine(StrPtr(buffer), StrPtr(dir), StrPtr(file))

    VtPathCombine = left$(buffer, InStr(buffer, vbNullChar) - 1)
End Function

' =========================
' 절대경로 여부
' =========================
Public Function VtPathIsAbsolute(ByVal Path As String) As Boolean
    If Len(Path) < 2 Then Exit Function

    If Mid$(Path, 2, 1) = ":" Then
        VtPathIsAbsolute = True
        Exit Function
    End If

    If left$(Path, 2) = "\\" Then
        VtPathIsAbsolute = True
    End If
End Function

' =========================
' 파일명 추출
' =========================
Public Function VtPathGetFileName(ByVal Path As String) As String
    Dim pos As Long
    pos = InStrRev(Path, "\")

    If pos > 0 Then
        VtPathGetFileName = Mid$(Path, pos + 1)
    Else
        VtPathGetFileName = Path
    End If
End Function

' =========================
' 확장자 추출
' =========================
Public Function VtPathGetExtension(ByVal Path As String) As String
    Dim fname As String
    Dim pos As Long

    fname = VtPathGetFileName(Path)
    pos = InStrRev(fname, ".")

    If pos > 0 Then
        VtPathGetExtension = Mid$(fname, pos + 1)
    End If
End Function

' =========================
' 확장자 제거
' =========================
Public Function VtPathRemoveExtension(ByVal Path As String) As String
    Dim pos As Long

    pos = InStrRev(Path, ".")

    If pos > 0 Then
        VtPathRemoveExtension = left$(Path, pos - 1)
    Else
        VtPathRemoveExtension = Path
    End If
End Function

' =========================
' 디렉토리 추출
' =========================
Public Function VtPathGetDirectory(ByVal Path As String) As String
    Dim pos As Long

    pos = InStrRev(Path, "\")

    If pos > 0 Then
        VtPathGetDirectory = left$(Path, pos - 1)
    End If
End Function

' =========================
' 경로 정규화
' =========================
Public Function VtPathNormalize(ByVal Path As String) As String
    Dim s As String
    s = Replace(Path, "/", "\")

    Do While InStr(s, "\\") > 0
        s = Replace(s, "\\", "\")
    Loop

    VtPathNormalize = s
End Function

' =========================
' 확장자 비교
' =========================
Public Function VtPathHasExtension(ByVal Path As String, ByVal ext As String) As Boolean
    Dim e As String
    e = LCase$(VtPathGetExtension(Path))

    VtPathHasExtension = (e = LCase$(ext))
End Function

' =========================
' 파일 존재 여부
' =========================
Public Function VtFileExists(ByVal Path As String) As Boolean
    VtFileExists = (dir(Path) <> "")
End Function

' =========================
' 디렉토리 존재 여부
' =========================
Public Function VtDirectoryExists(ByVal Path As String) As Boolean
    On Error Resume Next
    VtDirectoryExists = (GetAttr(Path) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function
