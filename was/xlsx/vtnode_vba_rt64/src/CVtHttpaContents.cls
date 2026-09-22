VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ============================================================
' CVtHttpaContents
'
' HTTP Request 또는 Response에 포함된
' CVtHttpaContentPart 컬렉션입니다.
'
' - 인덱스는 1-based
' - Item 기본 속성 지원
' - Form field name으로 Part 검색 가능
' - Native 객체에서 로드되었는지 상태 관리
' ============================================================

Private mContentList As Collection
Private mIsAttached As Boolean


' ============================================================
' Life cycle
' ============================================================

Private Sub Class_Initialize()

    Set mContentList = New Collection
    mIsAttached = False

End Sub


Private Sub Class_Terminate()

    Clear
    Set mContentList = Nothing

End Sub



' ============================================================
' Properties
' ============================================================

''' <summary>
''' 등록된 Content Part 개수를 반환합니다.
''' </summary>
Public Property Get count() As Long

    If mContentList Is Nothing Then
        count = 0
    Else
        count = mContentList.count
    End If

End Property




''' <summary>
''' Native Request 또는 Response에서
''' Content를 로드한 상태인지 반환합니다.
''' </summary>
Public Property Get IsAttached() As Boolean

    IsAttached = mIsAttached

End Property


''' <summary>
''' Content Part가 하나도 없는지 반환합니다.
''' </summary>
Public Property Get IsEmpty() As Boolean

    IsEmpty = (count = 0)

End Property

''' <summary>
''' 현재 컬렉션에 Part가 2개 이상 존재하는지 반환합니다.
'''
''' 주의:
''' 실제 Content-Type이 multipart인지 여부를 의미하지 않습니다.
''' </summary>
Public Property Get HasMultipleParts() As Boolean

    Dim FirstPart As CVtHttpaContentPart
    Dim ContentType As String

    Set FirstPart = GetPartAt(1)

    If FirstPart Is Nothing Then
        HasMultipleParts = False
        Exit Property
    End If

    ContentType = FirstPart.Headers.GetHeader("Content-Type")

    HasMultipleParts = _
        (InStr(1, ContentType, "multipart/", vbTextCompare) > 0)

End Property

''' <summary>
''' 1-based Index로 Content Part를 반환합니다.
'''
''' 사용 예:
'''   Set Part = Contents(1)
''' </summary>
Public Property Get item( _
    ByVal index As Long _
) As CVtHttpaContentPart

    Set item = GetPartAt(index)

End Property


' ============================================================
' Part Access
' ============================================================

''' <summary>
''' 1-based Index로 Content Part를 반환합니다.
''' 범위를 벗어나면 Nothing을 반환합니다.
''' </summary>
Public Function GetPartAt( _
    ByVal index As Long _
) As CVtHttpaContentPart

    On Error GoTo ErrorHandler

    Set GetPartAt = Nothing

    If mContentList Is Nothing Then
        Exit Function
    End If

    If index < 1 Or index > mContentList.count Then
        Exit Function
    End If

    Set GetPartAt = mContentList(index)
    Exit Function

ErrorHandler:

    Set GetPartAt = Nothing

End Function




''' <summary>
''' Form field name으로 Content Part를 검색합니다.
''' 찾지 못하면 Nothing을 반환합니다.
'''
''' Form field name은 대소문자를 구분합니다.
''' </summary>
Public Function GetPartByFormName( _
    ByVal FormName As String _
) As CVtHttpaContentPart

    Dim part As CVtHttpaContentPart
    Dim i As Long

    On Error GoTo ErrorHandler

    Set GetPartByFormName = Nothing

    If Len(FormName) = 0 Then
        Exit Function
    End If

    If mContentList Is Nothing Then
        Exit Function
    End If

    For i = 1 To mContentList.count

        Set part = mContentList(i)

        If Not part Is Nothing Then

            If StrComp( _
                part.FormName, _
                FormName, _
                vbBinaryCompare) = 0 Then

                Set GetPartByFormName = part
                Exit Function

            End If

        End If

    Next i

    Exit Function

ErrorHandler:

    Set GetPartByFormName = Nothing

End Function
''' <summary>
''' Form field name이 존재하는지 반환합니다.
''' </summary>
Public Function ContainsFormName( _
    ByVal FormName As String _
) As Boolean

    Dim part As CVtHttpaContentPart

    Set part = GetPartByFormName(FormName)

    ContainsFormName = Not (part Is Nothing)

End Function


' ============================================================
' Part Management
' ============================================================

''' <summary>
''' Content Part를 컬렉션에 추가합니다.
''' 성공하면 True를 반환합니다.
''' </summary>
Public Function AddPart( _
    ByVal part As CVtHttpaContentPart _
) As Boolean

    On Error GoTo ErrorHandler

    AddPart = False

    If part Is Nothing Then
        Exit Function
    End If

    If mContentList Is Nothing Then
        Set mContentList = New Collection
    End If

    mContentList.Add part

    AddPart = True
    Exit Function

ErrorHandler:

    Debug.Print _
        "[CVtHttpaContents.AddPart] Error " & _
        Err.Number & ": " & Err.Description

    AddPart = False

End Function

''' <summary>
''' 1-based Index의 Content Part를 제거합니다.
''' 성공하면 True를 반환합니다.
''' </summary>
Public Function RemovePartAt( _
    ByVal index As Long _
) As Boolean

    On Error GoTo ErrorHandler

    RemovePartAt = False

    If mContentList Is Nothing Then
        Exit Function
    End If

    If index < 1 Or index > mContentList.count Then
        Exit Function
    End If

    mContentList.Remove index

    RemovePartAt = True
    Exit Function

ErrorHandler:

    Debug.Print _
        "[CVtHttpaContents.RemovePartAt] Error " & _
        Err.Number & ": " & Err.Description

    RemovePartAt = False

End Function



''' <summary>
''' 모든 Content Part를 제거하고
''' Native 로드 상태도 초기화합니다.
''' </summary>
Public Sub Clear()

    On Error GoTo ErrorHandler

    Set mContentList = New Collection
    mIsAttached = False

    Exit Sub

ErrorHandler:

    Debug.Print _
        "[CVtHttpaContents.Clear] Error " & _
        Err.Number & ": " & Err.Description

End Sub


' ============================================================
' Native load state
' ============================================================

''' <summary>
''' Native 객체에서 Content Part 로드를
''' 완료한 상태로 표시합니다.
''' </summary>
Friend Sub MarkAttached()

    mIsAttached = True

End Sub


''' <summary>
''' Native 로드 상태를 해제합니다.
''' Content Part 자체는 유지합니다.
''' </summary>
Friend Sub MarkDetached()

    mIsAttached = False

End Sub

' ============================================================
' Summary
' ============================================================

''' <summary>
''' 현재 Content Part 컬렉션 정보를 문자열로 반환합니다.
''' </summary>
Public Function ToSummary() As String

    Dim result As String
    Dim part As CVtHttpaContentPart
    Dim headerText As String
    Dim i As Long

    result = "Contents" & vbCrLf
    result = result & "  Count: " & CStr(count) & vbCrLf
    result = result & "  IsAttached: " & CStr(mIsAttached) & vbCrLf
    result = result & "  HasMultipleParts: " _
                    & CStr(HasMultipleParts) & vbCrLf

    For i = 1 To count

        Set part = GetPartAt(i)

        result = result & vbCrLf
        result = result & "Part[" & CStr(i) & "]" & vbCrLf

        If part Is Nothing Then

            result = result & "  Nothing" & vbCrLf

        Else

            ' ------------------------------------------------
            ' Headers
            ' ------------------------------------------------
            headerText = part.Headers.ToSummary

            If Len(headerText) > 0 Then
                result = result & "  Headers" & vbCrLf
                result = result & headerText
            End If

            ' ------------------------------------------------
            ' Content Part
            ' ------------------------------------------------
            result = result & "  " _
                    & part.ToSummary _
                    & vbCrLf

        End If

    Next i

    ToSummary = result

End Function

