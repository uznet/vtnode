VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mQueryString As String
Private mQueryParams As Object     ' Scripting.Dictionary


Private Sub Class_Initialize()

    mQueryString = vbNullString

    Set mQueryParams = CreateObject("Scripting.Dictionary")

    ' 쿼리 파라미터 이름을 대소문자 구분 없이 처리하려면 활성화
    ' 반드시 데이터 추가 전에 설정해야 함
    ' mQueryParams.CompareMode = vbTextCompare

End Sub


Private Sub Class_Terminate()

    mQueryString = vbNullString
    Set mQueryParams = Nothing

End Sub


' ============================================================
' 기본 정보
' ============================================================

Public Property Get QueryString() As String

    QueryString = mQueryString

End Property


Public Property Get count() As Long

    If mQueryParams Is Nothing Then
        count = 0
    Else
        count = mQueryParams.count
    End If

End Property


Public Sub Clear()

    mQueryString = vbNullString

    If Not mQueryParams Is Nothing Then
        mQueryParams.RemoveAll
    End If

End Sub


Public Function Exists(ByVal key As String) As Boolean

    If mQueryParams Is Nothing Then
        Exists = False
        Exit Function
    End If

    Exists = mQueryParams.Exists(key)

End Function


' ============================================================
' Default Property
'
' 사용 예:
'   value = Query("table")
'   value = Query.Item("table")
' ============================================================

Public Property Get item(ByVal key As String) As String

    item = GetParam(key)

End Property


' Query는 요청 데이터이므로 일반적으로 읽기 전용이 적합함.
' 직접 설정도 허용하려면 아래 Property Let을 활성화할 수 있음.
'
' Public Default Property Let Item( _
'     ByVal key As String, _
'     ByVal value As String)
'
'     SetParam key, value
'
' End Property


' ============================================================
' Parsing
' ============================================================

Public Sub Parse(ByVal queryText As String)

    Dim pairs() As String
    Dim kv() As String

    Dim PairText As String
    Dim key As String
    Dim value As String

    Dim i As Long

    Clear

    queryText = NormalizeQueryString(queryText)

    If Len(queryText) = 0 Then
        Exit Sub
    End If

    mQueryString = queryText
    pairs = Split(queryText, "&")

    For i = LBound(pairs) To UBound(pairs)

        PairText = pairs(i)

        If Len(PairText) > 0 Then

            kv = Split(PairText, "=", 2)

            key = VtUrlDecode(kv(0))

            If UBound(kv) >= 1 Then
                value = VtUrlDecode(kv(1))
            Else
                value = vbNullString
            End If

            AddParam key, value

        End If

    Next i

End Sub


Private Function NormalizeQueryString( _
    ByVal queryText As String) As String

    Dim questionPos As Long

    queryText = Trim$(queryText)

    If Len(queryText) = 0 Then
        NormalizeQueryString = vbNullString
        Exit Function
    End If

    ' 전체 URL 또는 "/path?key=value"가 전달된 경우
    questionPos = InStr(1, queryText, "?", vbBinaryCompare)

    If questionPos > 0 Then
        queryText = Mid$(queryText, questionPos + 1)
    ElseIf left$(queryText, 1) = "?" Then
        queryText = Mid$(queryText, 2)
    End If

    ' Fragment는 서버로 전송되지 않지만 방어적으로 제거
    questionPos = InStr(1, queryText, "#", vbBinaryCompare)

    If questionPos > 0 Then
        queryText = left$(queryText, questionPos - 1)
    End If

    NormalizeQueryString = queryText

End Function


' ============================================================
' 값 조회
' ============================================================

' 첫 번째 값 반환
Public Function GetParam(ByVal key As String) As String

    Dim Values As Collection

    If Not Exists(key) Then
        GetParam = vbNullString
        Exit Function
    End If

    Set Values = mQueryParams(key)

    If Values Is Nothing Then
        GetParam = vbNullString
        Exit Function
    End If

    If Values.count = 0 Then
        GetParam = vbNullString
        Exit Function
    End If

    GetParam = CStr(Values(1))

End Function


' 같은 이름의 전체 값 반환
'
' 예:
'   ?menu=coffee&menu=tea
Public Function GetParams(ByVal key As String) As Collection

    If Not Exists(key) Then
        Set GetParams = Nothing
        Exit Function
    End If

    Set GetParams = mQueryParams(key)

End Function


' ============================================================
' 값 추가 및 변경
' ============================================================

' 기존 값은 유지하고 새 값을 추가
Public Function AddParam( _
    ByVal key As String, _
    ByVal value As String) As Boolean

    Dim Values As Collection

    On Error GoTo ErrorHandler

    If mQueryParams Is Nothing Then
        Set mQueryParams = CreateObject("Scripting.Dictionary")
    End If

    If Not mQueryParams.Exists(key) Then

        Set Values = New Collection
        mQueryParams.Add key, Values

    Else

        Set Values = mQueryParams(key)

    End If

    Values.Add value

    AddParam = True
    Exit Function

ErrorHandler:

    AddParam = False

End Function


' 기존 값을 모두 제거하고 하나의 값으로 교체
Public Function SetParam( _
    ByVal key As String, _
    ByVal value As String) As Boolean

    Dim Values As Collection

    On Error GoTo ErrorHandler

    If mQueryParams Is Nothing Then
        Set mQueryParams = CreateObject("Scripting.Dictionary")
    End If

    If mQueryParams.Exists(key) Then
        mQueryParams.Remove key
    End If

    Set Values = New Collection
    Values.Add value

    mQueryParams.Add key, Values

    SetParam = True
    Exit Function

ErrorHandler:

    SetParam = False

End Function


Public Function RemoveParam(ByVal key As String) As Boolean

    If Not Exists(key) Then
        RemoveParam = False
        Exit Function
    End If

    mQueryParams.Remove key
    RemoveParam = True

End Function


' ============================================================
' Query String 생성
' ============================================================

Public Function ToSummary() As String

    Dim key As Variant
    Dim value As Variant
    Dim result As String

    If mQueryParams Is Nothing Then
        ToSummary = vbNullString
        Exit Function
    End If

    If mQueryParams.count = 0 Then
        ToSummary = vbNullString
        Exit Function
    End If

    For Each key In mQueryParams.keys

        For Each value In mQueryParams(key)

            If Len(result) > 0 Then
                result = result & "&"
            End If

            result = result _
                   & VtUrlEncode(CStr(key)) _
                   & "=" _
                   & VtUrlEncode(CStr(value))

        Next value

    Next key

    ToSummary = result

End Function

