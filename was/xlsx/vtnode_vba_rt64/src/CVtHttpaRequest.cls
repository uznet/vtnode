VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ============================================================
' CVtHttpaRequest
'
' HTTP Request의 Method, URI, Query, Content, Cookie 등을 관리
' ============================================================

Private mLine As String
Private mMethod As VTHttpaMethod
Private mVersion As String

Private mUri As String
Private mPath As String
Private mQueryString As String

Private mQueryParams As CVtHttpaQuery
Private mContents As CVtHttpaContents
Private mCookies As CVtHttpaCookies

Private mIsAttached As Boolean


' ============================================================
' Class lifecycle
' ============================================================

Private Sub Class_Initialize()

    mLine = vbNullString
    mMethod = VTHttpaMethod.VTHTTPA_METHOD_UNKNOWN
    mVersion = vbNullString

    mUri = vbNullString
    mPath = vbNullString
    mQueryString = vbNullString

    Set mQueryParams = New CVtHttpaQuery
    Set mContents = New CVtHttpaContents
    Set mCookies = New CVtHttpaCookies

    mIsAttached = False
    
    Call AddDefaultContentPart

    Debug.Print "[CVtHttpaRequest] Initialize"

End Sub


Private Sub Class_Terminate()

    On Error Resume Next

    If Not mContents Is Nothing Then
        mContents.Clear
    End If

    Set mQueryParams = Nothing
    Set mContents = Nothing
    Set mCookies = Nothing

    mLine = vbNullString
    mMethod = VTHTTPA_METHOD_UNKNOWN
    mVersion = vbNullString

    mUri = vbNullString
    mPath = vbNullString
    mQueryString = vbNullString

    mIsAttached = False

End Sub



''' <summary>
''' 비어 있는 기본 Content Part를 생성하여 첫 번째 Part로 추가합니다.
''' </summary>
Private Function AddDefaultContentPart() As Boolean

    Dim ContentPart As CVtHttpaContentPart

    On Error GoTo EH

    AddDefaultContentPart = False

    If mContents Is Nothing Then
        Set mContents = New CVtHttpaContents
    End If

    Set ContentPart = New CVtHttpaContentPart

    ' 빈 ContentPart로 초기화합니다.
    If Not ContentPart.CreateEmpty() Then
        Set ContentPart = Nothing
        Exit Function
    End If

    If Not mContents.AddPart(ContentPart) Then
        Set ContentPart = Nothing
        Exit Function
    End If

    Set ContentPart = Nothing

    AddDefaultContentPart = True
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaRequest.AddDefaultContentPart] Error " & _
        Err.Number & ": " & Err.Description

    Set ContentPart = Nothing
    AddDefaultContentPart = False

End Function


' ============================================================
' Internal methods
' ============================================================

''' <summary>
''' Native Request Handle에서 Content Part들을 로드합니다.
'''
''' 기존 기본 Content Part 및 이전 Content Part를 모두 제거한 후
''' Native Request의 Content Part들로 다시 구성합니다.
'''
''' Native Request에 Content가 없으면 기본 Content Part 하나를
''' 다시 추가합니다.
''' </summary>
Private Function BuildContentsFromRequestHandle(ByVal RequestHandle As LongPtr) As Boolean

    Dim ContentCount As Long
    Dim ContentHandle As LongPtr
    Dim ContentPart As CVtHttpaContentPart
    Dim i As Long

    On Error GoTo EH

    BuildContentsFromRequestHandle = False

    If RequestHandle = 0 Then
        Exit Function
    End If

    If mContents Is Nothing Then
        Set mContents = New CVtHttpaContents
    Else
        ' Class_Initialize에서 생성한 기본 Part와
        ' 이전에 로드된 모든 Part를 제거합니다.
        Call mContents.Clear
    End If

    ContentCount = VtHttpaRequest.VtHttpaRequestGetContentCount(RequestHandle)

    'Debug.Print " BuildContentsFromRequestHandle.ContentCount=" & CStr(ContentCount)
    
    ' Native Request에 Content가 없는 경우에도
    ' 기본 Content Part 하나를 유지합니다.
    If ContentCount <= 0 Then

        BuildContentsFromRequestHandle = AddDefaultContentPart()

        Exit Function

    End If

    ' Native Request의 Content Part를 1-based Index로 로드합니다.
    For i = 1 To ContentCount

        ContentHandle = VtHttpaRequest.VtHttpaRequestGetContentAt(RequestHandle, i)

        
        
        If ContentHandle = 0 Then
        Debug.Print "BuildContentsFromRequestHandle:ContentHandle is NULL" & CStr(ContentHandle)
        
            GoTo LoadFailed
        End If

        Set ContentPart = New CVtHttpaContentPart

        If Not ContentPart.InitFromContentHandle(ContentHandle) Then

        Debug.Print "BuildContentsFromRequestHandle:InitFromContentHandle() failed" & CStr(ContentHandle)
            Set ContentPart = Nothing
            GoTo LoadFailed

        End If

        If Not mContents.AddPart(ContentPart) Then
            Set ContentPart = Nothing
            GoTo LoadFailed
        End If

        
        Call ContentPart.ToSummary
        
        Set ContentPart = Nothing

    Next i

    BuildContentsFromRequestHandle = True
    Exit Function


LoadFailed:

    'Debug.Print "BuildContentsFromRequestHandle:LoadFailed"
    Set ContentPart = Nothing

    ' 부분적으로 로드된 Content를 모두 제거합니다.
    mContents.Clear

    ' 실패하더라도 Request 객체는 기본 Part 하나를 유지합니다.
    Call AddDefaultContentPart

    BuildContentsFromRequestHandle = False
    Exit Function


EH:
    Debug.Print _
        "[CVtHttpaRequest.BuildContentsFromRequestHandle] Error " & _
        Err.Number & ": " & Err.Description

    Set ContentPart = Nothing

    If Not mContents Is Nothing Then
        mContents.Clear
        Call AddDefaultContentPart
    End If

    BuildContentsFromRequestHandle = False

End Function



''' <summary>
''' Native Request Handle에서 Request 정보를 로드합니다.
''' </summary>
Private Function InitFromRequestHandle( _
    ByVal RequestHandle As LongPtr _
) As Boolean

    Dim methodText As String

    On Error GoTo EH

    InitFromRequestHandle = False

    If RequestHandle = 0 Then
        Exit Function
    End If

    mLine = _
        VtHttpaRequest.VtHttpaRequestGetRequestLine( _
            RequestHandle)

    methodText = _
        VtHttpaRequest.VtHttpaRequestGetMethod( _
            RequestHandle)

    mMethod = _
        VtHttpaType.VTHttpaMethodFromString( _
            methodText)

    mVersion = _
        VtHttpaRequest.VtHttpaRequestGetVersion( _
            RequestHandle)

    mUri = _
        VtHttpaRequest.VtHttpaRequestGetUri( _
            RequestHandle)

    mPath = _
        VtHttpaRequest.VtHttpaRequestGetPath( _
            RequestHandle)

    mQueryString = _
        VtHttpaRequest.VtHttpaRequestGetQuery( _
            RequestHandle)

    If mQueryParams Is Nothing Then
        Set mQueryParams = New CVtHttpaQuery
    Else
        mQueryParams.Clear
    End If

    If Len(mQueryString) > 0 Then
        Call mQueryParams.Parse(mQueryString)
    End If

    If Not BuildContentsFromRequestHandle(RequestHandle) Then
        Exit Function
    End If

    Call BuildCookies

    
    InitFromRequestHandle = True
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaRequest.BuildFromRequestHandle] Error " & _
        Err.Number & ": " & Err.Description

    InitFromRequestHandle = False

End Function


''' <summary>
''' Cookie Header를 CVtHttpaCookies 객체에 반영합니다.
''' </summary>
Private Sub BuildCookies()

    Dim cookieHeader As String

    On Error GoTo EH

    If mCookies Is Nothing Then
        Set mCookies = New CVtHttpaCookies
    Else
        mCookies.Clear
    End If

    cookieHeader = Headers.GetHeader("Cookie")

    If Len(cookieHeader) > 0 Then
        Call mCookies.Parse(cookieHeader)
    End If

    Exit Sub

EH:
    Debug.Print _
        "[CVtHttpaRequest.BuildCookies] Error " & _
        Err.Number & ": " & Err.Description

End Sub


' ============================================================
' Creation
' ============================================================

''' <summary>
''' HTTP Event로부터 Request 정보를 로드합니다.
''' </summary>
Public Function initFromEvent( _
    ByVal EventId As LongLong _
) As Boolean

    Dim RequestHandle As LongPtr

    On Error GoTo EH

    initFromEvent = False
    mIsAttached = False

    If EventId = 0 Then
        Exit Function
    End If

    RequestHandle = VtHttpaRequest.VtHttpaRequestCreateFromEvent(EventId)

    If RequestHandle = 0 Then
        Exit Function
    End If

    If Not InitFromRequestHandle(RequestHandle) Then
        GoTo CleanUp
    End If

    mIsAttached = True
    initFromEvent = True

CleanUp:
    If RequestHandle <> 0 Then
        VtHttpaRequest.VtHttpaRequestDestroy RequestHandle
    End If

    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaRequest.CreateFromEvent] Error " & _
        Err.Number & ": " & Err.Description

    mIsAttached = False
    initFromEvent = False

    Resume CleanUp

End Function


' ============================================================
' Properties
' ============================================================

Public Property Get Method() As VTHttpaMethod

    Method = mMethod

End Property


Public Property Let Method( _
    ByVal value As VTHttpaMethod _
)

    mMethod = value

End Property


Public Property Get Version() As String

    Version = mVersion

End Property


Public Property Let Version( _
    ByVal value As String _
)

    mVersion = value

End Property


Public Property Get Uri() As String

    Uri = mUri

End Property


Public Property Let Uri( _
    ByVal value As String _
)

    mUri = value

End Property


Public Property Get Path() As String

    Path = mPath

End Property


Public Property Let Path( _
    ByVal value As String _
)

    mPath = value

End Property


Public Property Get QueryString() As String

    QueryString = mQueryString

End Property


Public Property Let QueryString( _
    ByVal value As String _
)

    mQueryString = value

    If mQueryParams Is Nothing Then
        Set mQueryParams = New CVtHttpaQuery
    Else
        mQueryParams.Clear
    End If

    If Len(mQueryString) > 0 Then
        Call mQueryParams.Parse(mQueryString)
    End If

End Property


Public Property Get line() As String

    line = mLine

End Property


Public Property Get QueryParams() As CVtHttpaQuery

    Set QueryParams = mQueryParams

End Property


Public Property Get Contents() As CVtHttpaContents

    Set Contents = mContents

End Property


''' <summary>
''' 첫 번째 Content Part를 Body로 반환합니다.
''' Content가 없으면 Nothing을 반환합니다.
''' </summary>
Public Property Get Body() As CVtHttpaContentPart

    Set Body = Nothing

    If mContents Is Nothing Then
        Exit Property
    End If

    If mContents.count < 1 Then
        Exit Property
    End If

    Set Body = mContents.GetPartAt(1)

End Property
Public Property Get Headers() As CVtHttpaHeaders

    Set Headers = Nothing

    If mContents Is Nothing Then
        Exit Property
    End If

    If mContents.count < 1 Then
        Exit Property
    End If

    Set Headers = Body.Headers


End Property


''' <summary>
''' Request Cookie 컬렉션을 반환합니다.
''' </summary>
Public Property Get Cookies() As CVtHttpaCookies

    Set Cookies = mCookies

End Property


Public Property Get IsAttached() As Boolean

    IsAttached = mIsAttached

End Property


' ============================================================
' Request 정보를 한 줄 문자열로 반환합니다.
' ============================================================
Public Function ToSummary() As String

    Dim result As String

    result = "HTTP Request: "

    If Len(mLine) > 0 Then
        result = result & mLine
    Else
        result = result & _
                 VtHttpaType.VTHttpaMethodToString(mMethod) & _
                 " " & _
                 mUri & _
                 " " & _
                 mVersion
    End If

    result = result & _
             ", ContentCount=" & _
             CStr(mContents.count)

    ToSummary = result

End Function

