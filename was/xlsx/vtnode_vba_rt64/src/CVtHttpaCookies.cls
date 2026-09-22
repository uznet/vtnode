VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaCookies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ============================================================
' CVtHttpaCookies
'
' HTTP Cookie 이름/값 컬렉션
'
' 주요 역할:
'   - Request의 Cookie 헤더 파싱
'   - Response의 Set-Cookie 헤더에서 이름/값 추출
'   - Cookie 추가, 수정, 조회, 제거
'   - Request Cookie 헤더 문자열 생성
'
' 주의:
'   Response Set-Cookie의 Path, Domain, Expires, Max-Age,
'   Secure, HttpOnly, SameSite 속성은 저장하지 않습니다.
'
' Cookie 이름은 대소문자를 구분합니다.
' NameAt(), ValueAt()의 Index는 1-based입니다.
' ============================================================

Private mCookies As Object    ' Scripting.Dictionary


' ============================================================
' Class lifecycle
' ============================================================

Private Sub Class_Initialize()

    InitializeDictionary

End Sub


Private Sub Class_Terminate()

    Set mCookies = Nothing

End Sub


' ============================================================
' Internal methods
' ============================================================

''' <summary>
''' 내부 Dictionary를 생성합니다.
''' Cookie 이름은 대소문자를 구분합니다.
''' </summary>
Private Sub InitializeDictionary()

    Set mCookies = CreateObject("Scripting.Dictionary")
    mCookies.CompareMode = vbBinaryCompare

End Sub


''' <summary>
''' 내부 Dictionary가 유효한 상태인지 확인합니다.
''' </summary>
Private Sub EnsureInitialized()

    If mCookies Is Nothing Then
        InitializeDictionary
    End If

End Sub


''' <summary>
''' 하나의 name=value 문자열을 분석하여 Cookie에 추가하거나 교체합니다.
'''
''' 값에 "=" 문자가 포함될 수 있으므로 첫 번째 "="만 구분자로 사용합니다.
'''
''' 예:
'''   sessionId=ABC123
'''   token=abc=def==
''' </summary>
Private Function ParsePair( _
    ByVal PairText As String _
) As Boolean

    Dim separatorPosition As Long
    Dim cookieName As String
    Dim CookieValue As String

    On Error GoTo EH

    ParsePair = False

    EnsureInitialized

    PairText = Trim$(PairText)

    If Len(PairText) = 0 Then
        Exit Function
    End If

    separatorPosition = _
        InStr(1, PairText, "=", vbBinaryCompare)

    ' Cookie 이름이 없거나 "="가 없는 경우
    If separatorPosition <= 1 Then
        Exit Function
    End If

    cookieName = _
        Trim$(left$(PairText, separatorPosition - 1))

    CookieValue = _
        Trim$(Mid$(PairText, separatorPosition + 1))

    If Len(cookieName) = 0 Then
        Exit Function
    End If

    If mCookies.Exists(cookieName) Then
        mCookies(cookieName) = CookieValue
    Else
        mCookies.Add cookieName, CookieValue
    End If

    ParsePair = True
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaCookies.ParsePair] Error " & _
        Err.Number & ": " & Err.Description

    ParsePair = False

End Function


' ============================================================
' Parsing
' ============================================================

''' <summary>
''' HTTP Request의 Cookie 헤더 값을 분석합니다.
'''
''' 기존 Cookie는 삭제하지 않고 추가하거나 교체합니다.
'''
''' 예:
'''   sessionId=ABC123; theme=dark; language=ko
''' </summary>
Public Sub Parse( _
    ByVal CookieString As String _
)

    Dim parts() As String
    Dim partValue As Variant

    On Error GoTo EH

    EnsureInitialized

    CookieString = Trim$(CookieString)

    If Len(CookieString) = 0 Then
        Exit Sub
    End If

    parts = Split(CookieString, ";")

    For Each partValue In parts
        Call ParsePair(CStr(partValue))
    Next partValue

    Exit Sub

EH:
    Debug.Print _
        "[CVtHttpaCookies.Parse] Error " & _
        Err.Number & ": " & Err.Description

End Sub


''' <summary>
''' HTTP Header 컬렉션으로부터 Cookie 정보를 생성합니다.
'''
''' IsRequest = True:
'''   Cookie 헤더의 모든 name=value 쌍을 분석합니다.
'''
''' IsRequest = False:
'''   각 Set-Cookie 헤더의 첫 번째 name=value 쌍만 분석합니다.
'''   Path, Domain, Expires 등의 속성은 저장하지 않습니다.
'''
''' 기존 Cookie는 모두 제거한 후 다시 생성합니다.
''' </summary>
Public Sub GenerateFromHeaders( _
    ByVal Headers As CVtHttpaHeaders, _
    ByVal IsRequest As Boolean _
)

    Dim headerList As Collection
    Dim HeaderValue As Variant

    Dim parts() As String
    Dim firstPair As String

    On Error GoTo EH

    Clear

    If Headers Is Nothing Then
        Exit Sub
    End If

    If IsRequest Then

        Set headerList = Headers.GetHeaders("Cookie")

        If headerList Is Nothing Then
            Exit Sub
        End If

        For Each HeaderValue In headerList

            If Len(Trim$(CStr(HeaderValue))) > 0 Then
                Parse CStr(HeaderValue)
            End If

        Next HeaderValue

    Else

        Set headerList = Headers.GetHeaders("Set-Cookie")

        If headerList Is Nothing Then
            Exit Sub
        End If

        For Each HeaderValue In headerList

            If Len(Trim$(CStr(HeaderValue))) > 0 Then

                parts = Split(CStr(HeaderValue), ";")

                If UBound(parts) >= LBound(parts) Then
                    firstPair = Trim$(CStr(parts(LBound(parts))))
                    Call ParsePair(firstPair)
                End If

            End If

        Next HeaderValue

    End If

    Exit Sub

EH:
    Debug.Print _
        "[CVtHttpaCookies.GenerateFromHeaders] Error " & _
        Err.Number & ": " & Err.Description

End Sub


' ============================================================
' Cookie management
' ============================================================

''' <summary>
''' 새로운 Cookie를 추가합니다.
''' 같은 이름의 Cookie가 이미 존재하면 False를 반환합니다.
''' </summary>
Public Function Add( _
    ByVal cookieName As String, _
    ByVal CookieValue As Variant _
) As Boolean

    On Error GoTo EH

    Add = False

    EnsureInitialized

    cookieName = Trim$(cookieName)

    If Len(cookieName) = 0 Then
        Exit Function
    End If

    If mCookies.Exists(cookieName) Then
        Exit Function
    End If

    mCookies.Add cookieName, CStr(CookieValue)

    Add = True
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaCookies.Add] Error " & _
        Err.Number & ": " & Err.Description

    Add = False

End Function


''' <summary>
''' 기존 Cookie 값을 교체합니다.
''' Cookie가 존재하지 않으면 False를 반환합니다.
''' </summary>
Public Function Replace( _
    ByVal cookieName As String, _
    ByVal CookieValue As Variant _
) As Boolean

    On Error GoTo EH

    Replace = False

    EnsureInitialized

    cookieName = Trim$(cookieName)

    If Len(cookieName) = 0 Then
        Exit Function
    End If

    If Not mCookies.Exists(cookieName) Then
        Exit Function
    End If

    mCookies(cookieName) = CStr(CookieValue)

    Replace = True
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaCookies.Replace] Error " & _
        Err.Number & ": " & Err.Description

    Replace = False

End Function


''' <summary>
''' Cookie가 존재하면 값을 교체하고, 존재하지 않으면 추가합니다.
''' </summary>
Public Function SetValue( _
    ByVal cookieName As String, _
    ByVal CookieValue As Variant _
) As Boolean

    On Error GoTo EH

    SetValue = False

    EnsureInitialized

    cookieName = Trim$(cookieName)

    If Len(cookieName) = 0 Then
        Exit Function
    End If

    If mCookies.Exists(cookieName) Then
        mCookies(cookieName) = CStr(CookieValue)
    Else
        mCookies.Add cookieName, CStr(CookieValue)
    End If

    SetValue = True
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaCookies.SetValue] Error " & _
        Err.Number & ": " & Err.Description

    SetValue = False

End Function


''' <summary>
''' 지정한 Cookie를 제거합니다.
''' 제거에 성공하면 True를 반환합니다.
''' </summary>
Public Function Remove( _
    ByVal cookieName As String _
) As Boolean

    On Error GoTo EH

    Remove = False

    EnsureInitialized

    cookieName = Trim$(cookieName)

    If Len(cookieName) = 0 Then
        Exit Function
    End If

    If Not mCookies.Exists(cookieName) Then
        Exit Function
    End If

    mCookies.Remove cookieName

    Remove = True
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaCookies.Remove] Error " & _
        Err.Number & ": " & Err.Description

    Remove = False

End Function


''' <summary>
''' 모든 Cookie를 제거합니다.
''' </summary>
Public Sub Clear()

    On Error GoTo EH

    EnsureInitialized

    mCookies.RemoveAll
    Exit Sub

EH:
    Debug.Print _
        "[CVtHttpaCookies.Clear] Error " & _
        Err.Number & ": " & Err.Description

End Sub


' ============================================================
' Lookup
' ============================================================

''' <summary>
''' 지정한 이름의 Cookie가 존재하는지 반환합니다.
''' </summary>
Public Function Contains( _
    ByVal cookieName As String _
) As Boolean

    On Error GoTo EH

    Contains = False

    EnsureInitialized

    cookieName = Trim$(cookieName)

    If Len(cookieName) = 0 Then
        Exit Function
    End If

    Contains = mCookies.Exists(cookieName)
    Exit Function

EH:
    Contains = False

End Function


''' <summary>
''' 지정한 Cookie 값을 반환합니다.
''' Cookie가 없으면 DefaultValue를 반환합니다.
''' </summary>
Public Function GetValue( _
    ByVal cookieName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String

    On Error GoTo EH

    GetValue = defaultValue

    EnsureInitialized

    cookieName = Trim$(cookieName)

    If Len(cookieName) = 0 Then
        Exit Function
    End If

    If mCookies.Exists(cookieName) Then
        GetValue = CStr(mCookies(cookieName))
    End If

    Exit Function

EH:
    GetValue = defaultValue

End Function


' ============================================================
' Properties
' ============================================================

''' <summary>
''' 등록된 Cookie 개수를 반환합니다.
''' </summary>
Public Property Get count() As Long

    EnsureInitialized

    count = mCookies.count

End Property


''' <summary>
''' Cookie가 하나도 없는지 반환합니다.
''' </summary>
Public Property Get IsEmpty() As Boolean

    IsEmpty = (count = 0)

End Property


''' <summary>
''' Cookie 이름으로 값을 반환합니다.
''' Cookie가 없으면 vbNullString을 반환합니다.
''' </summary>
Public Property Get item( _
    ByVal cookieName As String _
) As String

    item = GetValue(cookieName)

End Property


''' <summary>
''' Cookie 이름으로 값을 추가하거나 교체합니다.
''' </summary>
Public Property Let item( _
    ByVal cookieName As String, _
    ByVal CookieValue As String _
)

    Call SetValue(cookieName, CookieValue)

End Property


' ============================================================
' Indexed access
' ============================================================

''' <summary>
''' 1-based Index로 Cookie 이름을 반환합니다.
''' 범위를 벗어나면 vbNullString을 반환합니다.
''' </summary>
Public Function NameAt( _
    ByVal index As Long _
) As String

    Dim keys As Variant

    On Error GoTo EH

    NameAt = vbNullString

    EnsureInitialized

    If index < 1 Or index > mCookies.count Then
        Exit Function
    End If

    keys = mCookies.keys

    NameAt = CStr(keys(index - 1))
    Exit Function

EH:
    NameAt = vbNullString

End Function


''' <summary>
''' 1-based Index로 Cookie 값을 반환합니다.
''' 범위를 벗어나면 vbNullString을 반환합니다.
''' </summary>
Public Function ValueAt( _
    ByVal index As Long _
) As String

    Dim Values As Variant

    On Error GoTo EH

    ValueAt = vbNullString

    EnsureInitialized

    If index < 1 Or index > mCookies.count Then
        Exit Function
    End If

    Values = mCookies.items

    ValueAt = CStr(Values(index - 1))
    Exit Function

EH:
    ValueAt = vbNullString

End Function


''' <summary>
''' 1-based Index로 Cookie 이름과 값을 함께 반환합니다.
''' 성공하면 True를 반환합니다.
''' </summary>
Public Function GetAt( _
    ByVal index As Long, _
    ByRef CookieNameOut As String, _
    ByRef CookieValueOut As String _
) As Boolean

    On Error GoTo EH

    CookieNameOut = vbNullString
    CookieValueOut = vbNullString
    GetAt = False

    EnsureInitialized

    If index < 1 Or index > mCookies.count Then
        Exit Function
    End If

    CookieNameOut = NameAt(index)
    CookieValueOut = ValueAt(index)

    GetAt = True
    Exit Function

EH:
    CookieNameOut = vbNullString
    CookieValueOut = vbNullString
    GetAt = False

End Function


' ============================================================
' String generation
' ============================================================

''' <summary>
''' HTTP Request의 Cookie 헤더 값으로 사용할 문자열을 생성합니다.
'''
''' 예:
'''   sessionId=ABC123; theme=dark
''' </summary>
Public Function ToString() As String

    Dim cookieName As Variant
    Dim result As String

    On Error GoTo EH

    EnsureInitialized

    result = vbNullString

    For Each cookieName In mCookies.keys

        If Len(result) > 0 Then
            result = result & "; "
        End If

        result = _
            result & _
            CStr(cookieName) & _
            "=" & _
            CStr(mCookies(cookieName))

    Next cookieName

    ToString = result
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaCookies.ToString] Error " & _
        Err.Number & ": " & Err.Description

    ToString = vbNullString

End Function



