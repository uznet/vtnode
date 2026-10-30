VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

' ============================================================
' CVtHttpaResponse
'
' HTTP Response의 Version, Status, Header, Content 등을 관리
'
' Native Response Handle은 CreateFromEvent() 내부에서만 사용하며,
' 필요한 값을 VBA 객체로 복사한 후 즉시 해제합니다.
' ============================================================

Private mLine As String
Private mVersion As String
Private mStatusCode As Long
Private mReasonPhrase As String

Private mContents As CVtHttpaContents
Private mIsAttached As Boolean


' ============================================================
' Class lifecycle
' ============================================================

Private Sub Class_Initialize()

    mLine = vbNullString
    mVersion = vbNullString
    mStatusCode = 0
    mReasonPhrase = vbNullString

    Set mContents = New CVtHttpaContents

    ' Response는 항상 기본 Content Part 하나를 보유합니다.
    Call AddDefaultContentPart

    mIsAttached = False

End Sub


Private Sub Class_Terminate()

    On Error Resume Next

    If Not mContents Is Nothing Then
        mContents.Clear
    End If

    Set mContents = Nothing

    mLine = vbNullString
    mVersion = vbNullString
    mStatusCode = 0
    mReasonPhrase = vbNullString

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

    AddDefaultContentPart = _
        mContents.AddPart(ContentPart)

    Set ContentPart = Nothing
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaResponse.AddDefaultContentPart] Error " & _
        Err.Number & ": " & Err.Description

    Set ContentPart = Nothing
    AddDefaultContentPart = False

End Function
' ============================================================
' Internal methods
' ============================================================

''' <summary>
''' 현재 Response 데이터를 초기화합니다.
''' </summary>
Private Sub ClearResponseData()

    mLine = vbNullString
    mVersion = vbNullString
    mStatusCode = 0
    mReasonPhrase = vbNullString

    If mContents Is Nothing Then
        Set mContents = New CVtHttpaContents
    Else
        mContents.Clear
    End If

    mIsAttached = False

End Sub


''' <summary>
''' Native Response Handle에서 Content Part들을 로드합니다.
'''
''' Class_Initialize에서 생성한 기본 Content Part를 제거하고,
''' Native Response의 Content Part로 교체합니다.
'''
''' Native Response에 Content가 없으면 기본 Content Part를 다시 생성합니다.
''' </summary>
Private Function BuildContentsFromResponseHandle( _
    ByVal ResponseHandle As LongPtr _
) As Boolean

    Dim ContentCount As Long
    Dim ContentHandle As LongPtr
    Dim ContentPart As CVtHttpaContentPart
    Dim i As Long

    On Error GoTo EH

    BuildContentsFromResponseHandle = False

    If ResponseHandle = 0 Then
        Exit Function
    End If

    If mContents Is Nothing Then
        Set mContents = New CVtHttpaContents
    End If

    ' Class_Initialize에서 생성된 기본 Content Part를 제거합니다.
    mContents.Clear

    ContentCount = _
        VtHttpaResponse.VtHttpaResponseGetContentCount( _
            ResponseHandle)

    ' Native Response에 Content가 없으면
    ' 기본 Content Part 하나를 다시 유지합니다.
    If ContentCount <= 0 Then

        BuildContentsFromResponseHandle = _
            AddDefaultContentPart()

        Exit Function

    End If

    For i = 1 To ContentCount

        ContentHandle = _
            VtHttpaResponse.VtHttpaResponseGetContentAt( _
                ResponseHandle, _
                i)

        If ContentHandle = 0 Then
            GoTo LoadFailed
        End If

        Set ContentPart = New CVtHttpaContentPart

        If Not ContentPart.InitFromContentHandle(ContentHandle) Then
            Set ContentPart = Nothing
            GoTo LoadFailed
        End If

        If Not mContents.AddPart(ContentPart) Then
            Set ContentPart = Nothing
            GoTo LoadFailed
        End If

        Set ContentPart = Nothing

    Next i


    
    BuildContentsFromResponseHandle = True
    Exit Function

LoadFailed:
    Set ContentPart = Nothing

    ' 부분적으로 로드된 Content를 제거합니다.
    mContents.Clear

    ' 실패하더라도 Response 객체는 기본 Part 하나를 유지합니다.
    Call AddDefaultContentPart

    BuildContentsFromResponseHandle = False
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaResponse.BuildContentsFromResponseHandle] Error " & _
        Err.Number & ": " & Err.Description

    Set ContentPart = Nothing

    If Not mContents Is Nothing Then
        mContents.Clear
        Call AddDefaultContentPart
    End If

    BuildContentsFromResponseHandle = False

End Function



''' <summary>
''' Native Response Handle에서 모든 Response 정보를 로드합니다.
''' </summary>
Private Function InitFromResponseHandle( _
    ByVal ResponseHandle As LongPtr _
) As Boolean

    On Error GoTo EH

    InitFromResponseHandle = False

    If ResponseHandle = 0 Then
        Exit Function
    End If

    mLine = _
        VtHttpaResponse.VtHttpaResponseGetResponseLine( _
            ResponseHandle)

    mVersion = _
        VtHttpaResponse.VtHttpaResponseGetVersion( _
            ResponseHandle)

    mStatusCode = _
        VtHttpaResponse.VtHttpaResponseGetStatusCode( _
            ResponseHandle)

    mReasonPhrase = _
        VtHttpaResponse.VtHttpaResponseGetReasonPhrase( _
            ResponseHandle)

    If Not BuildContentsFromResponseHandle(ResponseHandle) Then
        Exit Function
    End If

    InitFromResponseHandle = True
    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaResponse.InitFromResponseHandle] Error " & _
        Err.Number & ": " & Err.Description

    InitFromResponseHandle = False

End Function


' ============================================================
' Creation
' ============================================================

''' <summary>
''' HTTP Event로부터 Response 정보를 생성합니다.
''' 성공하면 True를 반환합니다.
''' </summary>
Public Function initFromEvent( _
    ByVal EventId As LongLong _
) As Boolean

    Dim ResponseHandle As LongPtr

    On Error GoTo EH

    initFromEvent = False

    ClearResponseData

    If EventId = 0 Then
        Exit Function
    End If

    ResponseHandle = _
        VtHttpaResponse.VtHttpaResponseCreateFromEvent( _
            EventId)

    If ResponseHandle = 0 Then
        Exit Function
    End If

    If Not InitFromResponseHandle(ResponseHandle) Then
        GoTo CleanUp
    End If

    mIsAttached = True
    initFromEvent = True

CleanUp:
    If ResponseHandle <> 0 Then
        VtHttpaResponse.VtHttpaResponseDestroy ResponseHandle
    End If

    Exit Function

EH:
    Debug.Print _
        "[CVtHttpaResponse.InitFromEvent] Error " & _
        Err.Number & ": " & Err.Description

    mIsAttached = False
    initFromEvent = False

    Resume CleanUp

End Function


' ============================================================
' Properties
' ============================================================

''' <summary>
''' 전체 HTTP Response Line을 반환합니다.
''' 예: "HTTP/1.1 200 OK"
''' </summary>
Public Property Get line() As String

    line = mLine

End Property


''' <summary>
''' HTTP 버전을 반환합니다.
''' 예: "HTTP/1.1"
''' </summary>
Public Property Get Version() As String

    Version = mVersion

End Property


Public Property Let Version( _
    ByVal value As String _
)

    mVersion = value
    RebuildResponseLine

End Property


''' <summary>
''' HTTP 상태 코드를 반환합니다.
''' </summary>
Public Property Get StatusCode() As Long

    StatusCode = mStatusCode

End Property


Public Property Let StatusCode( _
    ByVal value As Long _
)

    mStatusCode = value
    RebuildResponseLine

End Property


''' <summary>
''' HTTP Reason Phrase를 반환합니다.
''' 예: "OK", "Not Found"
''' </summary>
Public Property Get ReasonPhrase() As String

    ReasonPhrase = mReasonPhrase

End Property


Public Property Let ReasonPhrase( _
    ByVal value As String _
)

    mReasonPhrase = value
    RebuildResponseLine

End Property


''' <summary>
''' Response에 포함된 Content 컬렉션을 반환합니다.
''' </summary>
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


    If Not Body Is Nothing Then
    
        Set Headers = Body.Headers
    Else
        Set Headers = Nothing
    End If
    
   
End Property


''' <summary>
''' Event로부터 Response가 정상적으로 생성되었는지 반환합니다.
''' </summary>
Public Property Get IsAttached() As Boolean

    IsAttached = mIsAttached

End Property


' ============================================================
' Status classification
' ============================================================

''' <summary>
''' 1xx Informational 응답인지 반환합니다.
''' </summary>
Public Property Get IsInformational() As Boolean

    IsInformational = _
        (mStatusCode >= 100 And mStatusCode < 200)

End Property


''' <summary>
''' 2xx Success 응답인지 반환합니다.
''' </summary>
Public Property Get IsSuccess() As Boolean

    IsSuccess = _
        (mStatusCode >= 200 And mStatusCode < 300)

End Property


''' <summary>
''' 3xx Redirect 응답인지 반환합니다.
''' </summary>
Public Property Get IsRedirect() As Boolean

    IsRedirect = _
        (mStatusCode >= 300 And mStatusCode < 400)

End Property


''' <summary>
''' 4xx Client Error 응답인지 반환합니다.
''' </summary>
Public Property Get IsClientError() As Boolean

    IsClientError = _
        (mStatusCode >= 400 And mStatusCode < 500)

End Property


''' <summary>
''' 5xx Server Error 응답인지 반환합니다.
''' </summary>
Public Property Get IsServerError() As Boolean

    IsServerError = _
        (mStatusCode >= 500 And mStatusCode < 600)

End Property


''' <summary>
''' 4xx 또는 5xx 응답인지 반환합니다.
''' </summary>
Public Property Get IsError() As Boolean

    IsError = _
        (mStatusCode >= 400 And mStatusCode < 600)

End Property


' ============================================================
' Helper methods
' ============================================================

''' <summary>
''' Version, StatusCode, ReasonPhrase로 Response Line을 갱신합니다.
''' </summary>
Private Sub RebuildResponseLine()

    If Len(mVersion) = 0 And _
       mStatusCode = 0 And _
       Len(mReasonPhrase) = 0 Then

        mLine = vbNullString
        Exit Sub

    End If

    mLine = Trim$( _
        mVersion & _
        " " & _
        CStr(mStatusCode) & _
        " " & _
        mReasonPhrase)

End Sub


''' <summary>
''' 전체 HTTP Status Line을 반환합니다.
''' </summary>
Public Function GetStatusLine() As String

    If Len(mLine) > 0 Then
        GetStatusLine = mLine
    Else
        GetStatusLine = Trim$( _
            mVersion & _
            " " & _
            CStr(mStatusCode) & _
            " " & _
            mReasonPhrase)
    End If

End Function




Public Function SendStatus( _
    ByVal s As CVtHttpaSession, _
    ByVal StatusCode As Long, _
    ByRef errorOut As String _
) As Boolean

    Dim VtResponseHandle As LongPtr
    Dim SendErrorOut As String

    SendStatus = False
    errorOut = vbNullString

    On Error GoTo EH

    Me.StatusCode = StatusCode
    
    ' --------------------------------------------------------
    ' 1. Create Response Handle
    ' --------------------------------------------------------
    VtResponseHandle = _
        VtHttpaResponse.VtHttpaResponseCreateWithStatus(StatusCode)

    If VtResponseHandle = 0 Then
        errorOut = "Failed to create response handle."
        Exit Function
    End If

    ' --------------------------------------------------------
    ' 2. Copy VBA Headers -> Native Response
    ' --------------------------------------------------------
    If Not Me.Headers.CopyToResponseHandle(VtResponseHandle) Then

        errorOut = "CopyToResponseHandle failed."

        Debug.Print "[SendStatus] " & errorOut

        GoTo CleanUp

    End If

    ' --------------------------------------------------------
    ' 3. Send Response
    ' --------------------------------------------------------
    If VtHttpaServerSendResponse( _
            s.IID, _
            VtResponseHandle, _
            SendErrorOut) Then

        s.HttpaService.AddActivity_Response Me, VT_RESPONSE_STATUS
        SendStatus = True
        
        


    Else

        errorOut = SendErrorOut

        Debug.Print _
            "[SendStatus] VtHttpaServerSendResponse failed: " & _
            errorOut

    End If

CleanUp:

    If VtResponseHandle <> 0 Then

        VtHttpaResponse.VtHttpaResponseDestroy _
            VtResponseHandle

        VtResponseHandle = 0

    End If

    Exit Function

EH:

    errorOut = _
        "Error " & _
        Err.Number & _
        ": " & _
        Err.Description

    Debug.Print _
        "[SendStatus] " & errorOut

    Resume CleanUp

End Function

Private Function IsValidStatusCode(ByVal StatusCode As Long) As Boolean

IsValidStatusCode = _
        (StatusCode >= 100 And StatusCode <= 599)



End Function



Public Function SendText( _
    ByVal s As CVtHttpaSession, _
    ByVal text As String, _
    ByRef errorOut As String _
) As Boolean

    Dim VtResponseHandle As LongPtr
    Dim SendErrorOut As String
    'Dim ContentType As String

    SendText = False
    errorOut = vbNullString

    On Error GoTo EH

    ' --------------------------------------------------------
    ' 1. Status Code
    ' --------------------------------------------------------
    If Not IsValidStatusCode(Me.StatusCode) Then
        Me.StatusCode = 200
    End If

    ' --------------------------------------------------------
    ' 2. Content-Type
    ' --------------------------------------------------------
    If Not Me.Headers.Exists("Content-Type") Then

        Call Me.Headers.SetHeader("Content-Type", "text/plain; charset=utf-8")

    End If

    'ContentType = Me.Headers.GetHeader("Content-Type")

    ' --------------------------------------------------------
    ' 3. Create Response Handle
    ' --------------------------------------------------------
    VtResponseHandle = _
        VtHttpaResponse.VtHttpaResponseCreateWithText(text)

    If VtResponseHandle = 0 Then

        errorOut = "Failed to create response handle."
        Exit Function

    End If

    ' --------------------------------------------------------
    ' 4. Set Status Code
    ' --------------------------------------------------------
    Call VtHttpaResponse.VtHttpaResponseSetStatusCode(VtResponseHandle, Me.StatusCode)
    
    ' --------------------------------------------------------
    ' 5. Copy VBA Headers -> Native Response
    ' --------------------------------------------------------
    If Not Me.Headers.CopyToResponseHandle(VtResponseHandle) Then

        errorOut = "CopyToResponseHandle failed."

        Debug.Print _
            "[SendText] " & errorOut

        GoTo CleanUp

    End If

    ' --------------------------------------------------------
    ' 6. Send Response
    ' --------------------------------------------------------
    If VtHttpaServerSendResponse( _
            s.IID, _
            VtResponseHandle, _
            SendErrorOut) Then

        s.HttpaService.AddActivity_Response Me, VT_RESPONSE_TEXT
        SendText = True
        

    Else

        errorOut = SendErrorOut

        Debug.Print _
            "[SendText] VtHttpaServerSendResponse failed: " & _
            errorOut

    End If

CleanUp:

    If VtResponseHandle <> 0 Then

        VtHttpaResponse.VtHttpaResponseDestroy _
            VtResponseHandle

        VtResponseHandle = 0

    End If

    Exit Function

EH:

    errorOut = _
        "Error " & _
        Err.Number & _
        ": " & _
        Err.Description

    Debug.Print _
        "[SendText] " & errorOut

    Resume CleanUp

End Function


Public Function SendFile( _
    ByVal s As CVtHttpaSession, _
    ByVal FilePath As String, _
    ByRef errorOut As String _
) As Boolean

    Dim VtResponseHandle As LongPtr
    Dim SendErrorOut As String

    SendFile = False
    errorOut = vbNullString

    On Error GoTo EH

    ' --------------------------------------------------------
    ' 1. File Path
    ' --------------------------------------------------------
    If Len(Trim$(FilePath)) = 0 Then
        errorOut = "FilePath is empty."
        Exit Function
    End If

    ' --------------------------------------------------------
    ' 2. Status Code
    ' --------------------------------------------------------
    If Not IsValidStatusCode(Me.StatusCode) Then
        Me.StatusCode = 200
    End If

    ' --------------------------------------------------------
    ' 3. Create Native Response
    '
    ' File existence / path validation is handled
    ' by the native C engine.
    ' --------------------------------------------------------
    VtResponseHandle = _
        VtHttpaResponse.VtHttpaResponseCreateWithFile(FilePath)

    If VtResponseHandle = 0 Then

        errorOut = _
            "Failed to create file response: " & _
            FilePath

        Exit Function
    End If

    ' --------------------------------------------------------
    ' 4. Set Status Code
    ' --------------------------------------------------------
    Call VtHttpaResponse.VtHttpaResponseSetStatusCode(VtResponseHandle, Me.StatusCode)


    ' --------------------------------------------------------
    ' 5. Copy VBA Headers -> Native Response
    ' --------------------------------------------------------
    If Not Me.Headers.CopyToResponseHandle(VtResponseHandle) Then

        errorOut = "CopyToResponseHandle failed."

        Debug.Print _
            "[SendFile] " & errorOut

        GoTo CleanUp
    End If

    ' --------------------------------------------------------
    ' 6. Send Response
    ' --------------------------------------------------------
    If VtHttpaServerSendResponse( _
            s.IID, _
            VtResponseHandle, _
            SendErrorOut) Then
        s.HttpaService.AddActivity_Response Me, VT_RESPONSE_FILE

        SendFile = True

    Else

        errorOut = SendErrorOut

        Debug.Print _
            "[SendFile] VtHttpaServerSendResponse failed: " & _
            errorOut

    End If

CleanUp:

    If VtResponseHandle <> 0 Then

        VtHttpaResponse.VtHttpaResponseDestroy _
            VtResponseHandle

        VtResponseHandle = 0

    End If

    Exit Function

EH:

    errorOut = _
        "Error " & _
        Err.Number & _
        ": " & _
        Err.Description

    Debug.Print _
        "[SendFile] " & errorOut

    Resume CleanUp

End Function



Public Function SendMultipart( _
    ByVal s As CVtHttpaSession, _
    ByVal MultipartSubtype As String, _
    ByRef errorOut As String _
) As Boolean

    Dim ResponseHandle As LongPtr
    Dim ContentPart As CVtHttpaContentPart
    Dim ContentHandle As LongPtr
    Dim i As Long

    On Error GoTo ErrorHandler

    SendMultipart = False
    errorOut = vbNullString
    ResponseHandle = 0

    ' --------------------------------------------------------
    ' 1. Validate
    ' --------------------------------------------------------
    Debug.Print "SendMultiPart:1"
    If s Is Nothing Then
        errorOut = "Session is Nothing."
        Exit Function
    End If

    If mContents Is Nothing Then
        errorOut = "Contents is Nothing."
        Exit Function
    End If

    If mContents.count <= 0 Then
        errorOut = "No content parts."
        Exit Function
    End If

    MultipartSubtype = Trim$(MultipartSubtype)

    If Len(MultipartSubtype) = 0 Then
        errorOut = "MultipartSubtype is empty."
        Exit Function
    End If

    ' --------------------------------------------------------
    ' 2. Status Code
    ' --------------------------------------------------------
    Debug.Print "SendMultiPart:2"
    If Not IsValidStatusCode(Me.StatusCode) Then
        Me.StatusCode = 200
    End If

    ' --------------------------------------------------------
    ' 3. Create Multipart Response
    ' --------------------------------------------------------
    Debug.Print "SendMultiPart:3"
    
    ResponseHandle = _
        VtHttpaResponseCreateAsMultipart(MultipartSubtype)

    If ResponseHandle = 0 Then
        errorOut = _
            "Create multipart response failed. (" & _
            MultipartSubtype & ")"
        GoTo CleanUp
    End If

    Call VtHttpaResponse.VtHttpaResponseSetStatusCode(ResponseHandle, Me.StatusCode)


    ' --------------------------------------------------------
    ' 4. Copy Response Headers
    ' --------------------------------------------------------
    Debug.Print "SendMultiPart:4"

    If Not Me.Headers.CopyToResponseHandle(ResponseHandle) Then
        errorOut = "Copy response headers failed."
        GoTo CleanUp
    End If

    ' --------------------------------------------------------
    ' 5. Add Content Parts
    ' --------------------------------------------------------
    
    Debug.Print "SendMultiPart:5->ContentsCpount=" & CStr(mContents.count)
    
    For i = 2 To mContents.count


        Set ContentPart = mContents.GetPartAt(i)
        
        Debug.Print "SendMultiPart:6->part(" & CStr(i) & ")"
        Debug.Print "SendMultiPart:6->DispositionName=" & ContentPart.ContentName
        
            

        ContentHandle = 0   'null

        If ContentPart Is Nothing Then
            errorOut = _
                "ContentPart is Nothing. Index=" & CStr(i)
            GoTo CleanUp
        End If

        If ContentPart.HasDataType(VTHTTPA_CONTENT_DATATYPE_FILE) Then

            ContentHandle = VtHttpaResponse.VtHttpaResponseAddContentFile( _
                    ResponseHandle, _
                    ContentPart.FilePath, _
                    ContentPart.ContentType, _
                    ContentPart.ContentData.ContentDisposition.DispName, _
                    ContentPart.ContentData.ContentDisposition.DispFilename)

        ElseIf ContentPart.HasDataType(VTHTTPA_CONTENT_DATATYPE_TEXT) Then

            ContentHandle = VtHttpaResponse.VtHttpaResponseAddContentText( _
                    ResponseHandle, _
                    ContentPart.GetText(), _
                    ContentPart.ContentType, _
                    ContentPart.ContentData.ContentDisposition.DispName)

        ElseIf ContentPart.HasDataType(VTHTTPA_CONTENT_DATATYPE_BYTES) Then

            errorOut = _
                "Bytes content is not yet implemented. Index=" & _
                CStr(i)

            GoTo CleanUp

        Else

            errorOut = _
                "Unsupported content data type. Index=" & CStr(i) & ",datatype=" & CStr(ContentPart.DataType)

            GoTo CleanUp

        End If

        If ContentHandle <> 0 Then
           ContentPart.Headers.CopyToContentHandle (ContentHandle)
        Else
            errorOut = _
                "Add content part failed. Index=" & CStr(i)
            GoTo CleanUp
        End If

        Set ContentPart = Nothing

    Next i

    ' --------------------------------------------------------
    ' 6. Send
    ' --------------------------------------------------------
    Debug.Print "SendMultiPart:6"
    Dim Res As Boolean
    
    Res = _
        VtHttpaServerSendResponse( _
            s.IID, _
            ResponseHandle, _
            errorOut)

    Debug.Print "SendMultiPart:7"
    If Not Res Then
        GoTo CleanUp
    End If

    s.HttpaService.AddActivity_Response Me, VT_RESPONSE_MULTIPART

    SendMultipart = True

CleanUp:

    Set ContentPart = Nothing

    If ResponseHandle <> 0 Then
        VtHttpaResponseDestroy ResponseHandle
        ResponseHandle = 0
    End If

    Exit Function

ErrorHandler:

    errorOut = _
        "CVtHttpaResponse.SendMultipart error: " & _
        CStr(Err.Number) & ", " & Err.Description

    Resume CleanUp

End Function


Public Function AddFile( _
    ByVal FilePath As String, _
    Optional ByVal ContentType As String, _
    Optional ByVal ContentName As String = vbNullString _
) As CVtHttpaContentPart

    Dim ContentPart As CVtHttpaContentPart

    Set AddFile = Nothing

    Set ContentPart = New CVtHttpaContentPart

    If Not ContentPart.InitAsFile(FilePath, ContentType) Then
        Debug.Print "AddFile: InitAsFile() Failed"
        GoTo CleanUp
    End If

    ContentPart.ContentName = ContentName

    If Not mContents.AddPart(ContentPart) Then
        Debug.Print "AddFile: AddPart() Failed"
        GoTo CleanUp
    End If

    
    Set AddFile = ContentPart

CleanUp:

    Set ContentPart = Nothing

End Function

Public Function AddText( _
    ByVal text As String, _
    Optional ByVal ContentType As String = vbNullString, _
    Optional ByVal ContentName As String = vbNullString _
) As CVtHttpaContentPart

    Dim ContentPart As CVtHttpaContentPart

    Set AddText = Nothing

    Set ContentPart = New CVtHttpaContentPart

    If Not ContentPart.InitAsText(text, ContentType) Then
        Debug.Print "AddText: InitAsText() Failed"
        GoTo CleanUp
    End If

    ContentPart.ContentName = ContentName
    
    

    If Not mContents.AddPart(ContentPart) Then
        Debug.Print "AddText: AddPart() Failed"
        GoTo CleanUp
    End If

    Set AddText = ContentPart

CleanUp:

    Set ContentPart = Nothing

End Function


Public Function AddBytes( _
    ByRef Bytes() As Byte, _
    Optional ByVal ContentType As String = vbNullString, _
    Optional ByVal ContentName As String = vbNullString _
) As CVtHttpaContentPart

    Dim ContentPart As CVtHttpaContentPart

    Set AddBytes = Nothing

    Set ContentPart = New CVtHttpaContentPart

    If Not ContentPart.InitAsBytes(Bytes, ContentType) Then
        Debug.Print "AddBytes: InitAsBytes() Failed"
        GoTo CleanUp
    End If

    ContentPart.ContentName = ContentName

    If Not mContents.AddPart(ContentPart) Then
        Debug.Print "AddBytes: AddPart() Failed"
        GoTo CleanUp
    End If

    Set AddBytes = ContentPart

CleanUp:

    Set ContentPart = Nothing

End Function
''' <summary>
''' Response 정보를 한 줄 문자열로 반환합니다.
''' </summary>
Public Function ToSummary() As String

    ToSummary = _
        "HTTP Response: " & _
        GetStatusLine() & _
        ", ContentCount=" & _
        CStr(mContents.count)

End Function


