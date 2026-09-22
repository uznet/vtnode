VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private mHeaders As Object    ' Scripting.Dictionary


Private Sub Class_Initialize()

    Set mHeaders = CreateObject("Scripting.Dictionary")
    mHeaders.CompareMode = vbTextCompare

End Sub


Private Sub Class_Terminate()

    Set mHeaders = Nothing

End Sub


' ============================================================
' Default Item Property
'
' Examples:
'   Headers("Content-Type") = "application/json"
'   Debug.Print Headers("Content-Type")
' ============================================================

Public Property Get item(ByVal HdrName As String) As String


    item = GetHeader(HdrName)

End Property


Public Property Let item(ByVal HdrName As String, ByVal HdrValue As String)

    SetHeader HdrName, HdrValue

End Property


' ============================================================
' Build / Apply
' ============================================================

' Content Handle에서 모든 헤더를 읽어 현재 객체를 구성한다.
Public Function BuildFromContentHandle( _
    ByVal VtContentHandle As LongLong) As Boolean

    On Error GoTo ErrorHandler

    Dim index As Long
    Dim HdrName As String
    Dim HdrValue As String
    Dim result As Boolean

    Clear

    If VtContentHandle = 0 Then
        BuildFromContentHandle = False
        Exit Function
    End If

    index = 1

    Do
        HdrName = vbNullString
        HdrValue = vbNullString

        result = VtHttpaContentGetHeaderAt( _
                    VtContentHandle, _
                    index, _
                    HdrName, _
                    HdrValue)

        If Not result Then Exit Do

        If Len(Trim$(HdrName)) > 0 Then
            AddInternal HdrName, HdrValue
        End If

        index = index + 1
    Loop

    BuildFromContentHandle = True
    Exit Function

ErrorHandler:
    Debug.Print "CVtHeaders.BuildFromContentHandle error: " & Err.Description
    BuildFromContentHandle = False

End Function


' 현재 객체의 모든 헤더를 Content Handle에 적용한다.
Public Function ApplyToContent( _
    ByVal ContentHandle As LongLong) As Boolean

    Dim HdrName As String
    Dim HdrValues As Collection
    Dim i As Long
    Dim j As Long

    If ContentHandle = 0 Then
        ApplyToContent = False
        Exit Function
    End If

    If UniqueHeaderCount = 0 Then
        ApplyToContent = True
        Exit Function
    End If

    For i = 1 To UniqueHeaderCount

        HdrName = HeaderNameAt(i)

        If Len(HdrName) > 0 Then

            Set HdrValues = ToLinesByPolicy(HdrName)

            For j = 1 To HdrValues.count

                If Not VtHttpaContent.VtHttpaContentAddHeader( _
                            ContentHandle, _
                            HdrName, _
                            CStr(HdrValues(j))) Then

                    ApplyToContent = False
                    Exit Function

                End If

            Next j

        End If

    Next i

    ApplyToContent = True

End Function


' ============================================================
' Header Name Normalization
' ============================================================

Private Function NormalizeHeaderName( _
    ByVal HdrName As String) As String

    NormalizeHeaderName = LCase$(Trim$(HdrName))

End Function


Private Sub ValidateHeaderName(ByVal HdrName As String)

    If Len(Trim$(HdrName)) = 0 Then
        Err.Raise vbObjectError + 1001, _
                  "CVtHeaders", _
                  "Header name is empty."
    End If

End Sub


' ============================================================
' Internal Storage
' ============================================================

Private Sub AddInternal( _
    ByVal HdrName As String, _
    ByVal HdrValue As String)

    Dim key As String
    Dim Values As Collection

    key = NormalizeHeaderName(HdrName)

    If Not mHeaders.Exists(key) Then
        Set Values = New Collection
        mHeaders.Add key, Values
    Else
        Set Values = mHeaders(key)
    End If

    Values.Add HdrValue

End Sub


' ============================================================
' Header Access
' ============================================================

' 첫 번째 값을 반환한다.
Public Function GetHeader(ByVal HdrName As String) As String

    Dim key As String
    Dim Values As Collection

    key = NormalizeHeaderName(HdrName)

    If Not mHeaders.Exists(key) Then
        GetHeader = vbNullString
        Exit Function
    End If

    Set Values = mHeaders(key)

    If Values.count = 0 Then
        GetHeader = vbNullString
        Exit Function
    End If

    GetHeader = CStr(Values(1))

End Function


' 같은 이름을 가진 모든 값을 반환한다.
' 존재하지 않으면 Nothing을 반환한다.
Public Function GetHeaders(ByVal HdrName As String) As Collection

    Dim key As String

    key = NormalizeHeaderName(HdrName)

    If Not mHeaders.Exists(key) Then
        Set GetHeaders = Nothing
        Exit Function
    End If

    Set GetHeaders = mHeaders(key)

End Function


' 기존 값을 유지하면서 값을 추가한다.
Public Sub AddHeader( _
    ByVal HdrName As String, _
    ByVal HdrValue As String)

    ValidateHeaderName HdrName
    AddInternal HdrName, HdrValue

End Sub


' 기존 값을 모두 제거하고 하나의 값으로 교체한다.
Public Sub SetHeader( _
    ByVal HdrName As String, _
    ByVal HdrValue As String)

    ValidateHeaderName HdrName

    RemoveHeader HdrName
    AddInternal HdrName, HdrValue

End Sub


Public Function RemoveHeader(ByVal HdrName As String) As Boolean

    Dim key As String

    key = NormalizeHeaderName(HdrName)

    If mHeaders.Exists(key) Then
        mHeaders.Remove key
        RemoveHeader = True
    Else
        RemoveHeader = False
    End If

End Function


Public Function Exists(ByVal HdrName As String) As Boolean

    Dim key As String

    key = NormalizeHeaderName(HdrName)
    Exists = mHeaders.Exists(key)

End Function


Public Sub Clear()

    If Not mHeaders Is Nothing Then
        mHeaders.RemoveAll
    End If

End Sub


' ============================================================
' Date Header
' ============================================================

Public Sub SetDateToGMT( _
    ByVal HdrName As String, _
    ByVal dtValue As Date)

    Dim httpDate As String

    httpDate = VtDateToGmtString(dtValue)

    If Len(httpDate) = 0 Then
        Err.Raise vbObjectError + 1002, _
                  "CVtHeaders.SetDateToGMT", _
                  "Invalid date value: " & CStr(dtValue)
    End If

    SetHeader HdrName, httpDate

End Sub


Public Function GetDateFromGMT(ByVal HdrName As String) As Date

    Dim HdrValue As String
    Dim dt As Date

    HdrValue = GetHeader(HdrName)

    If Len(HdrValue) = 0 Then
        Err.Raise vbObjectError + 1003, _
                  "CVtHeaders.GetDateFromGMT", _
                  "Header not found: " & HdrName
    End If

    dt = VtGmtStringToDate(HdrValue)

    If dt = 0 Then
        Err.Raise vbObjectError + 1004, _
                  "CVtHeaders.GetDateFromGMT", _
                  "Invalid GMT date: " & HdrValue
    End If

    GetDateFromGMT = dt

End Function


' ============================================================
' Count / Index
' ============================================================

' 고유한 헤더 이름 개수
Public Property Get UniqueHeaderCount() As Long

    If mHeaders Is Nothing Then
        UniqueHeaderCount = 0
    Else
        UniqueHeaderCount = mHeaders.count
    End If

End Property


' 모든 헤더 값의 총 개수
Public Property Get count() As Long

    Dim key As Variant
    Dim Values As Collection
    Dim totalCount As Long

    If mHeaders Is Nothing Then
        count = 0
        Exit Property
    End If

    For Each key In mHeaders.keys
        Set Values = mHeaders(key)
        totalCount = totalCount + Values.count
    Next key

    count = totalCount

End Property


Public Function GetHeaderValueCount( _
    ByVal HdrName As String) As Long

    Dim key As String
    Dim Values As Collection

    key = NormalizeHeaderName(HdrName)

    If Not mHeaders.Exists(key) Then
        GetHeaderValueCount = 0
        Exit Function
    End If

    Set Values = mHeaders(key)
    GetHeaderValueCount = Values.count

End Function


' 고유 헤더 이름을 1-based index로 반환한다.
Public Function HeaderNameAt(ByVal index As Long) As String

    Dim key As Variant
    Dim currentIndex As Long

    If mHeaders Is Nothing Then
        HeaderNameAt = vbNullString
        Exit Function
    End If

    If index < 1 Or index > mHeaders.count Then
        HeaderNameAt = vbNullString
        Exit Function
    End If

    currentIndex = 1

    For Each key In mHeaders.keys

        If currentIndex = index Then
            HeaderNameAt = CStr(key)
            Exit Function
        End If

        currentIndex = currentIndex + 1

    Next key

    HeaderNameAt = vbNullString

End Function


' ============================================================
' Merge Policy
' ============================================================

Private Function JoinCollection( _
    ByVal Values As Collection, _
    ByVal Separator As String) As String

    Dim value As Variant
    Dim result As String

    For Each value In Values

        If Len(result) > 0 Then
            result = result & Separator
        End If

        result = result & CStr(value)

    Next value

    JoinCollection = result

End Function


Public Function ToLinesByPolicy( _
    ByVal HdrName As String) As Collection

    Dim result As Collection
    Dim key As String
    Dim mergePolicy As VTHttpaHeaderMergePolicy
    Dim value As Variant

    Set result = New Collection
    key = NormalizeHeaderName(HdrName)

    If Not mHeaders.Exists(key) Then
        Set ToLinesByPolicy = result
        Exit Function
    End If

    mergePolicy = VtHttpaContent.VtHttpaHeaderGetMergePolicy(key)

    Select Case mergePolicy

        Case VTHTTPA_MERGE_COMMA
            result.Add JoinCollection(mHeaders(key), ", ")

        Case VTHTTPA_MERGE_SEMICOLON
            result.Add JoinCollection(mHeaders(key), "; ")

        Case VTHTTPA_MERGE_MULTI_KEEP
            For Each value In mHeaders(key)
                result.Add CStr(value)
            Next value

        Case Else
            If mHeaders(key).count > 0 Then
                result.Add CStr(mHeaders(key)(1))
            End If

    End Select

    Set ToLinesByPolicy = result

End Function

Public Function CopyToContentHandle( _
    ByVal VtContentHandle As LongPtr _
) As Boolean

    Dim HdrName As String
    Dim HdrValues As Collection
    Dim i As Long
    Dim j As Long

    CopyToContentHandle = False

    If VtContentHandle = 0 Then
        Exit Function
    End If

    ' Header가 없는 것은 정상
    If Me.count <= 0 Then
        CopyToContentHandle = True
        Exit Function
    End If

    On Error GoTo EH

    For i = 1 To Me.count

        HdrName = Me.HeaderNameAt(i)

        If HdrName <> vbNullString Then

            Set HdrValues = ToLinesByPolicy(HdrName)

            If Not HdrValues Is Nothing Then

                For j = 1 To HdrValues.count

                    If Not VtHttpaContent.VtHttpaContentAddHeader( _
                                VtContentHandle, _
                                HdrName, _
                                CStr(HdrValues(j))) Then

                        Exit Function

                    End If

                Next j

            End If

            Set HdrValues = Nothing

        End If

    Next i

    CopyToContentHandle = True
    Exit Function

EH:

    Debug.Print _
        "[CVtHttpaHeaders.CopyToContentHandle] Error " & _
        Err.Number & ": " & Err.Description

End Function

Public Function CopyToResponseHandle( _
    ByVal VtResponseHandle As LongPtr _
) As Boolean

    Dim HdrName As String
    Dim HdrValues As Collection
    Dim i As Long
    Dim j As Long

    CopyToResponseHandle = False

    If VtResponseHandle = 0 Then
        Exit Function
    End If

    ' Header가 없는 것은 오류가 아님
    If Me.count <= 0 Then
        CopyToResponseHandle = True
        Exit Function
    End If

    On Error GoTo EH

    For i = 1 To Me.count

        HdrName = Me.HeaderNameAt(i)

        If HdrName <> vbNullString Then

            Set HdrValues = Me.ToLinesByPolicy(HdrName)

            If Not HdrValues Is Nothing Then

                For j = 1 To HdrValues.count
                    If Not VtHttpaResponse.VtHttpaResponseAddHeader( _
                                VtResponseHandle, _
                                HdrName, _
                                CStr(HdrValues(j))) Then

                        Debug.Print "CopyToResponseHandle(" & HdrName & "->failed)" & CStr(HdrValues(j))
                        Exit Function

                    End If

                Next j

            End If

            Set HdrValues = Nothing

        End If

    Next i

    CopyToResponseHandle = True
    Exit Function

EH:

    Debug.Print _
        "[CVtHttpaHeaders.CopyToResponseHandle] Error " & _
        Err.Number & ": " & Err.Description

End Function
' ============================================================
' Summary
' ============================================================

Public Function ToSummary() As String

    Dim key As Variant
    Dim lines As Collection
    Dim lineValue As Variant
    Dim result As String

    If mHeaders Is Nothing Then
        ToSummary = vbNullString
        Exit Function
    End If

    For Each key In mHeaders.keys

        Set lines = ToLinesByPolicy(CStr(key))

        If Not lines Is Nothing Then

            For Each lineValue In lines

                result = result _
                       & CStr(key) _
                       & ": " _
                       & CStr(lineValue) _
                       & vbCrLf

            Next lineValue

        End If

    Next key

    ToSummary = result

End Function
' example code

'Dim Headers As CVtHeaders
'Set Headers = New CVtHeaders
'
'Headers("Content-Type") = "application/json"
'
'Headers.AddHeader "Set-Cookie", "SessionID=100"
'Headers.AddHeader "Set-Cookie", "Theme=Dark"
'
'Debug.Print Headers("content-type")
'Debug.Print Headers.ToSummary
