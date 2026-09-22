Attribute VB_Name = "VtCharUtils"
Option Explicit

Private Const CP_UTF8 As Long = 65001
#If VBA7 And Win64 Then


Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As Long)

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long, _
        ByRef lpMultiByteStr As Any, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As LongPtr, _
        ByVal lpUsedDefaultChar As LongPtr) As Long

' 64비트 및 VBA7 환경용 교정 선언
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As LongPtr, _
        ByVal cbMultiByte As Long, _
        ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long) As Long
        
    Private Declare PtrSafe Function PathCombine Lib "shlwapi.dll" Alias "PathCombineW" ( _
    ByVal pszDest As LongPtr, _
    ByVal pszDir As LongPtr, _
    ByVal pszFile As LongPtr _
) As LongPtr
#Else
' 32비트 환경용 교정 선언

#End If


Public Function VtUtf8PtrToStringWithLen(ByVal pUtf8 As LongPtr, ByVal nLen As Long) As String
    Dim nWideLen As Long
    Dim result As String

    If pUtf8 = 0 Or nLen <= 0 Then
        VtUtf8PtrToStringWithLen = ""
        Exit Function
    End If

    ' 1. 변환될 유니코드 문자열의 필요한 길이를 먼저 계산
    nWideLen = MultiByteToWideChar(CP_UTF8, 0, pUtf8, nLen, 0, 0)

    If nWideLen > 0 Then
        ' 2. 결과를 담을 공간 확보 (VBA String은 2바이트 단위)
        result = String$(nWideLen, 0)

        ' 3. 실제 변환 수행
        MultiByteToWideChar CP_UTF8, 0, pUtf8, nLen, StrPtr(result), nWideLen

        VtUtf8PtrToStringWithLen = result
    End If
End Function

' UTF-8 포인터 → VBA String
' ptr is null-terminated UTF-8 string pointer
' 포인터가 가리키는 UTF-8 문자열이 0이 나올 때까지 길이를 계산하여 문자열로 변환
Public Function VtUtf8PtrToString(ptr As LongPtr) As String

    If ptr = 0 Then
        VtUtf8PtrToString = vbNullString
        Exit Function
    End If
    Dim nLen As Long
    nLen = 0
    ' 0이 나올 때까지 길이 계산
    Do While True
        Dim b As Byte
        CopyMemory VarPtr(b), ptr + nLen, 1
        If b = 0 Then Exit Do
        nLen = nLen + 1
    Loop
    VtUtf8PtrToString = VtUtf8PtrToStringWithLen(ptr, nLen)

End Function

' VBA String → UTF-8 Byte() (null-terminated)
Public Function VtStringToUtf8Bytes(ByVal strInput As String) As Byte()
    Dim nLen As Long
    Dim nUtf8Len As Long
    Dim buffer() As Byte

    nLen = Len(strInput)
    If nLen <= 0 Then
        ' 빈 문자열의 경우 널 문자 1개만 반환
        ReDim buffer(0 To 0)
        buffer(0) = 0
        VtStringToUtf8Bytes = buffer
        Exit Function
    End If

    ' 필요한 UTF-8 바이트 수 계산 (널 제외)
    nUtf8Len = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strInput), nLen, 0, 0, 0, 0)

    If nUtf8Len > 0 Then
        ' 널 문자를 위한 공간 1바이트 추가
        ReDim buffer(0 To nUtf8Len)

        ' 실제 변환
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(strInput), nLen, buffer(0), nUtf8Len, 0, 0)

        ' 널 종료 문자 추가
        buffer(nUtf8Len) = 0
    Else
        ' 변환 실패 시 널 문자만 반환
        ReDim buffer(0 To 0)
        buffer(0) = 0
    End If

    VtStringToUtf8Bytes = buffer
End Function
' UTF-8 Byte() → VBA String
Public Function VtUtf8BytesToString(Data() As Byte) As String
    Dim nUtf8Len As Long
    Dim nWideLen As Long
    Dim result As String

    On Error Resume Next
    nUtf8Len = UBound(Data) - LBound(Data) + 1
    
    'Debug.Print "VtUtf8BytesToString.utf8Len=" & nUtf8Len
    
    If Err.Number <> 0 Then
        Err.Clear
        VtUtf8BytesToString = ""
        Exit Function
    End If
    On Error GoTo 0

    If nUtf8Len <= 0 Then
        VtUtf8BytesToString = ""
        Exit Function
    End If

    ' 1. 변환될 유니코드 문자열의 필요한 길이를 먼저 계산
    ' ? MultiByteToWideChar로 수정
    nWideLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Data(LBound(Data))), nUtf8Len, 0, 0)
    
    'Debug.Print "VtUtf8BytesToString.wideLen=" & nWideLen

    If nWideLen > 0 Then
        ' 2. 결과를 담을 공간 확보 (VBA String은 2바이트 단위)
        result = String$(nWideLen, 0)
        ' 3. 실제 변환 수행
        ' ? MultiByteToWideChar로 수정
        Call MultiByteToWideChar(CP_UTF8, 0, VarPtr(Data(LBound(Data))), nUtf8Len, StrPtr(result), nWideLen)
        
        'Debug.Print "VtUtf8BytesToString.str=" & result
        VtUtf8BytesToString = result
    Else
        VtUtf8BytesToString = ""
    End If
End Function





'Public Function VtUtf8PtrToBytes(ByVal DataPtr As LongPtr, ByVal DataLen As Long) As Byte()
'    Dim buf() As Byte
'    Dim cb As Long
'
'    If DataPtr = 0 Then
'        ReDim buf(0 To -1)
'        VtUtf8PtrToBytes = buf
'        Exit Function
'    End If
'
'    ' 넘겨받은 확실한 길이를 그대로 사용 (성능 우수)
'    cb = DataLen
'
'    If DataPtr = 0 Or cb <= 0 Then
'        ReDim buf(0 To -1)
'        VtUtf8PtrToBytes = buf
'        Exit Function
'    End If
'
'    ReDim buf(0 To cb - 1)
'
'    ' ? 기존: CopyMemory buf(0), ptr, cb
'    ' ? 수정: 매개변수명인 dataPtr로 일치시킴
'    Call CopyMemory(VarPtr(buf(0)), DataPtr, cb)
'
'    VtUtf8PtrToBytes = buf
'End Function


Public Function VtPtrToBytes( _
    ByVal DataPtr As LongPtr, _
    ByVal DataLen As Long _
) As Byte()

    Dim result() As Byte
    Dim copyLength As Long

    On Error GoTo EH

    ' NULL 포인터 또는 길이가 0이면 초기화되지 않은 빈 배열 반환
    If DataPtr = 0 Or DataLen <= 0 Then
        VtPtrToBytes = result
        Exit Function
    End If

    ' VBA Byte 배열의 크기와 ReDim 인덱스는 Long 범위로 제한
    If DataLen > &H7FFFFFFF Then
        Err.Raise _
            vbObjectError + 1001, _
            "VtPtrToBytes", _
            "Data length exceeds the VBA Byte array limit."
    End If

    copyLength = CLng(DataLen)

    ReDim result(0 To copyLength - 1)

    CopyMemory _
        VarPtr(result(0)), _
        DataPtr, _
        copyLength

    VtPtrToBytes = result
    Exit Function

EH:
    Debug.Print _
        "[VtPtrToBytes] Error " & _
        Err.Number & ": " & Err.Description

    Erase result
    VtPtrToBytes = result

End Function




Public Function VtByteArrayLength(ByRef Bytes() As Byte) As Long

    On Error GoTo EH

    VtByteArrayLength = UBound(Bytes) - LBound(Bytes) + 1
    Exit Function

EH:
    VtByteArrayLength = 0

End Function
