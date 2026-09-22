Attribute VB_Name = "VtHttpaResponse"
Option Explicit


#If VBA7 And Win64 Then

#If VTNODE_DEBUG = 1 Then  ' Debug Mode


'============================================================
' HTTP RESPONSE ACCESS
'============================================================

' Create a new response handle from an event.
Private Declare PtrSafe Function vtnode_httpa_response_create_from_event Lib "vtnode64d.dll" ( _
    ByVal EventId As LongLong _
) As LongPtr

' Create a new response handle with status code and reason phrase.
Private Declare PtrSafe Function vtnode_httpa_response_create_with_status Lib "vtnode64d.dll" ( _
    ByVal StatusCode As Long, _
    ByVal Reason As LongPtr _
) As LongPtr

' Create a new response handle with file content.
' FilePath is the path to the file to be sent as response body.
' ContentType is the MIME type of the file, e.g. "application/json", "text/plain", etc.
Private Declare PtrSafe Function vtnode_httpa_response_create_with_file Lib "vtnode64d.dll" ( _
    ByVal FilePath As LongPtr, _
    ByVal ContentType As LongPtr _
) As LongPtr

' Create a new response handle with text content (null-terminated UTF-8 string).
' ContentType is the MIME type of the text, e.g. "application/json", "text/plain", etc.
Private Declare PtrSafe Function vtnode_httpa_response_create_with_text Lib "vtnode64d.dll" ( _
    ByVal text As LongPtr, _
    ByVal ContentType As LongPtr _
) As LongPtr

' Create a new response handle as multipart/form-data or multipart/mixed.
' MultipartSubType is the subtype of multipart, e.g. "form-data" or "mixed".
Private Declare PtrSafe Function vtnode_httpa_response_create_as_multipart Lib "vtnode64d.dll" ( _
    ByVal MultipartSubtype As LongPtr _
) As LongPtr

' Destroy a response handle.
Private Declare PtrSafe Sub vtnode_httpa_response_destroy Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr _
)

' Set response status code.
Private Declare PtrSafe Sub vtnode_httpa_response_set_status_code Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal StatusCode As Long _
)

' Set response reason phrase.
Private Declare PtrSafe Sub vtnode_httpa_response_set_reason_phrase Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal Reason As LongPtr _
)

' Set HTTP version, e.g. "HTTP/1.1".
Private Declare PtrSafe Sub vtnode_httpa_response_set_version Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal Version As LongPtr _
)

' Set response line, e.g. "HTTP/1.1 200 OK".
Private Declare PtrSafe Sub vtnode_httpa_response_set_response_line Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal ResponseLine As LongPtr _
)

' Get response status code.
Private Declare PtrSafe Function vtnode_httpa_response_get_status_code Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr _
) As Long

' Get response reason phrase.
' Returned string is read-only UTF-8 and owned by response.
Private Declare PtrSafe Function vtnode_httpa_response_get_reason_phrase Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr _
) As LongPtr

' Get HTTP version.
' Returned string is read-only UTF-8 and owned by response.
Private Declare PtrSafe Function vtnode_httpa_response_get_version Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr _
) As LongPtr

' Get response line.
' Returned string is read-only UTF-8 and owned by response.
Private Declare PtrSafe Function vtnode_httpa_response_get_response_line Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr _
) As LongPtr

' Add or set header value by header name.
Private Declare PtrSafe Function vtnode_httpa_response_add_header Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Set header value by header name. If header already exists, it will be replaced.
Private Declare PtrSafe Function vtnode_httpa_response_set_header Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Remove header by header name. If header does not exist, do nothing.
Private Declare PtrSafe Sub vtnode_httpa_response_remove_header Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As LongPtr _
)

' Get header value by header name.
' Return NULL if not found.
' Returned string is read-only UTF-8 and owned by response.
Private Declare PtrSafe Function vtnode_httpa_response_get_header Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As LongPtr _
) As LongPtr

' Get header by 1-based index.
' Return FALSE if not found.
' Returned strings are read-only UTF-8 and owned by response.
Private Declare PtrSafe Function vtnode_httpa_response_get_header_at Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal index As Long, _
    ByRef HeaderNameOut As LongPtr, _
    ByRef HeaderValueOut As LongPtr _
) As Long

' Get content data by 1-based index.
' Return FALSE if not found.
' OutData points to file path or memory buffer.
Private Declare PtrSafe Function vtnode_httpa_response_get_content_at Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal index As Long _
) As LongPtr


' Add a new content part to the response.
' Response must be created by vtnode_httpa_response_create_as_multipart().
Private Declare PtrSafe Function vtnode_httpa_response_content_add_file Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal FilePath As LongPtr, _
    ByVal ContentType As LongPtr, _
    ByVal ContentDispositionName As LongPtr, _
    ByVal ContentDispositionFilename As LongPtr _
) As LongPtr

' Add a new text content part to the response.
' Response must be created by vtnode_httpa_response_create_as_multipart().
Private Declare PtrSafe Function vtnode_httpa_response_content_add_text Lib "vtnode64d.dll" ( _
    ByVal Response As LongPtr, _
    ByVal text As LongPtr, _
    ByVal ContentType As LongPtr, _
    ByVal ContentDispositionName As LongPtr _
) As LongPtr

#Else  ' Rleased Mode


' Create a new response handle from an event.
Private Declare PtrSafe Function vtnode_httpa_response_create_from_event Lib "vtnode64.dll" ( _
    ByVal EventId As LongLong _
) As LongPtr

' Create a new response handle with status code and reason phrase.
Private Declare PtrSafe Function vtnode_httpa_response_create_with_status Lib "vtnode64.dll" ( _
    ByVal StatusCode As Long, _
    ByVal Reason As LongPtr _
) As LongPtr

' Create a new response handle with file content.
' FilePath is the path to the file to be sent as response body.
' ContentType is the MIME type of the file, e.g. "application/json", "text/plain", etc.
Private Declare PtrSafe Function vtnode_httpa_response_create_with_file Lib "vtnode64.dll" ( _
    ByVal FilePath As LongPtr, _
    ByVal ContentType As LongPtr _
) As LongPtr

' Create a new response handle with text content (null-terminated UTF-8 string).
' ContentType is the MIME type of the text, e.g. "application/json", "text/plain", etc.
Private Declare PtrSafe Function vtnode_httpa_response_create_with_text Lib "vtnode64.dll" ( _
    ByVal text As LongPtr, _
    ByVal ContentType As LongPtr _
) As LongPtr

' Create a new response handle as multipart/form-data or multipart/mixed.
' MultipartSubType is the subtype of multipart, e.g. "form-data" or "mixed".
Private Declare PtrSafe Function vtnode_httpa_response_create_as_multipart Lib "vtnode64.dll" ( _
    ByVal MultipartSubtype As LongPtr _
) As LongPtr

' Destroy a response handle.
Private Declare PtrSafe Sub vtnode_httpa_response_destroy Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr _
)

' Set response status code.
Private Declare PtrSafe Sub vtnode_httpa_response_set_status_code Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal StatusCode As Long _
)

' Set response reason phrase.
Private Declare PtrSafe Sub vtnode_httpa_response_set_reason_phrase Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal Reason As LongPtr _
)

' Set HTTP version, e.g. "HTTP/1.1".
Private Declare PtrSafe Sub vtnode_httpa_response_set_version Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal Version As LongPtr _
)

' Set response line, e.g. "HTTP/1.1 200 OK".
Private Declare PtrSafe Sub vtnode_httpa_response_set_response_line Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal ResponseLine As LongPtr _
)

' Get response status code.
Private Declare PtrSafe Function vtnode_httpa_response_get_status_code Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr _
) As Long

' Get response reason phrase.
' Returned string is read-only UTF-8 and owned by response.
Private Declare PtrSafe Function vtnode_httpa_response_get_reason_phrase Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr _
) As LongPtr

' Get HTTP version.
' Returned string is read-only UTF-8 and owned by response.
Private Declare PtrSafe Function vtnode_httpa_response_get_version Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr _
) As LongPtr

' Get response line.
' Returned string is read-only UTF-8 and owned by response.
Private Declare PtrSafe Function vtnode_httpa_response_get_response_line Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr _
) As LongPtr

' Add or set header value by header name.
Private Declare PtrSafe Function vtnode_httpa_response_add_header Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Set header value by header name. If header already exists, it will be replaced.
Private Declare PtrSafe Function vtnode_httpa_response_set_header Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Remove header by header name. If header does not exist, do nothing.
Private Declare PtrSafe Sub vtnode_httpa_response_remove_header Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As LongPtr _
)

' Get header value by header name.
' Return NULL if not found.
' Returned string is read-only UTF-8 and owned by response.
Private Declare PtrSafe Function vtnode_httpa_response_get_header Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As LongPtr _
) As LongPtr

' Get header by 1-based index.
' Return FALSE if not found.
' Returned strings are read-only UTF-8 and owned by response.
Private Declare PtrSafe Function vtnode_httpa_response_get_header_at Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal index As Long, _
    ByRef HeaderNameOut As LongPtr, _
    ByRef HeaderValueOut As LongPtr _
) As Long

' Get content data by 1-based index.
' Return FALSE if not found.
' OutData points to file path or memory buffer.
Private Declare PtrSafe Function vtnode_httpa_response_get_content_at Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal index As Long _
) As LongPtr


' Add a new content part to the response.
' Response must be created by vtnode_httpa_response_create_as_multipart().
Private Declare PtrSafe Function vtnode_httpa_response_content_add_file Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal FilePath As LongPtr, _
    ByVal ContentType As LongPtr, _
    ByVal ContentDispositionName As LongPtr, _
    ByVal ContentDispositionFilename As LongPtr _
) As LongPtr

' Add a new text content part to the response.
' Response must be created by vtnode_httpa_response_create_as_multipart().
Private Declare PtrSafe Function vtnode_httpa_response_content_add_text Lib "vtnode64.dll" ( _
    ByVal Response As LongPtr, _
    ByVal text As LongPtr, _
    ByVal ContentType As LongPtr, _
    ByVal ContentDispositionName As LongPtr _
) As LongPtr




#End If

#End If




' ==============================================================================
' HTTP RESPONSE WRAPPER
' ==============================================================================

''' <summary>
''' EventId로부터 HTTP Response 핸들을 생성합니다.
''' </summary>
''' <param name="EventId">HTTP 이벤트 ID</param>
''' <returns>생성된 Response 핸들. 실패 시 0</returns>
Public Function VtHttpaResponseCreateFromEvent( _
    ByVal EventId As LongLong _
) As LongPtr

    Dim ResponseHandle As LongPtr

    On Error GoTo EH

    VtHttpaResponseCreateFromEvent = 0

    If EventId = 0 Then
        Debug.Print _
            "[VtHttpaResponseCreateFromEvent] Invalid EventId."
        Exit Function
    End If

    ResponseHandle = _
        vtnode_httpa_response_create_from_event(EventId)

    If ResponseHandle = 0 Then
        Debug.Print _
            "[VtHttpaResponseCreateFromEvent] " & _
            "Failed to create response from EventId: " & EventId
    End If

    VtHttpaResponseCreateFromEvent = ResponseHandle
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaResponseCreateFromEvent] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseCreateFromEvent = 0

End Function


''' <summary>
''' 상태 코드와 Reason Phrase를 사용하여 HTTP Response 핸들을 생성합니다.
''' </summary>
''' <param name="StatusCode">HTTP 상태 코드. 예: 200, 404, 500</param>
''' <param name="Reason">Reason Phrase. 비어 있으면 NULL 전달</param>
''' <returns>생성된 Response 핸들. 실패 시 0</returns>
Public Function VtHttpaResponseCreateWithStatus( _
    ByVal StatusCode As Long, _
    Optional ByVal Reason As String = vbNullString _
) As LongPtr

    Dim utf8Reason() As Byte
    Dim pReason As LongPtr
    Dim ResponseHandle As LongPtr

    On Error GoTo EH

    VtHttpaResponseCreateWithStatus = 0

    If StatusCode < 100 Or StatusCode > 999 Then
        Debug.Print _
            "[VtHttpaResponseCreateWithStatus] " & _
            "Invalid status code: " & StatusCode
        Exit Function
    End If

    If Len(Reason) > 0 Then
        utf8Reason = _
            VtCharUtils.VtStringToUtf8Bytes(Reason)

        pReason = VarPtr(utf8Reason(0))
    Else
        pReason = 0
    End If

    ResponseHandle = _
        vtnode_httpa_response_create_with_status( _
            StatusCode, _
            pReason)

    If ResponseHandle = 0 Then
        Debug.Print _
            "[VtHttpaResponseCreateWithStatus] " & _
            "Failed to create response. StatusCode=" & _
            StatusCode
    End If

    VtHttpaResponseCreateWithStatus = ResponseHandle
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaResponseCreateWithStatus] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseCreateWithStatus = 0

End Function


''' <summary>
''' 파일을 Response Body로 갖는 HTTP Response 핸들을 생성합니다.
''' </summary>
''' <param name="FilePath">전송할 파일 경로</param>
''' <param name="ContentType">
''' 파일의 MIME 타입. 비어 있으면 application/octet-stream 사용
''' </param>
''' <returns>생성된 Response 핸들. 실패 시 0</returns>
Public Function VtHttpaResponseCreateWithFile(ByVal FilePath As String) As LongPtr

    Dim utf8FilePath() As Byte
    'Dim utf8ContentType() As Byte

    Dim ResponseHandle As LongPtr

    On Error GoTo EH

    VtHttpaResponseCreateWithFile = 0

    If Len(Trim$(FilePath)) = 0 Then
        Debug.Print _
            "[VtHttpaResponseCreateWithFile] File path is empty."
        Exit Function
    End If

    If Len(dir$( _
        FilePath, _
        vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then

        Debug.Print _
            "[VtHttpaResponseCreateWithFile] " & _
            "File not found: " & FilePath
        Exit Function
    End If

'    If Len(Trim$(ContentType)) = 0 Then
'        ContentType = "application/octet-stream"
'    End If

    utf8FilePath = VtCharUtils.VtStringToUtf8Bytes(FilePath)

    'utf8ContentType = VtCharUtils.VtStringToUtf8Bytes(ContentType)

    ResponseHandle = vtnode_httpa_response_create_with_file(VarPtr(utf8FilePath(0)), 0)

    If ResponseHandle = 0 Then
        Debug.Print _
            "[VtHttpaResponseCreateWithFile] " & _
            "Failed to create file response: " & FilePath
    End If

    VtHttpaResponseCreateWithFile = ResponseHandle
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaResponseCreateWithFile] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseCreateWithFile = 0

End Function


''' <summary>
''' UTF-8 텍스트를 Response Body로 갖는 HTTP Response 핸들을 생성합니다.
''' </summary>
'''
''' <param name="Text">
''' 응답 본문 문자열
''' </param>
'''
''' <returns>
''' 생성된 Response 핸들. 실패 시 0
''' </returns>
Public Function VtHttpaResponseCreateWithText( _
    ByVal text As String _
) As LongPtr

    Dim utf8Text() As Byte
    Dim TextPtr As LongPtr
    Dim ResponseHandle As LongPtr

    VtHttpaResponseCreateWithText = 0

    On Error GoTo EH

    ' --------------------------------------------------------
    ' 1. Convert Text -> UTF-8
    ' --------------------------------------------------------
    If Len(text) > 0 Then

        utf8Text = _
            VtCharUtils.VtStringToUtf8Bytes(text)

        TextPtr = VarPtr(utf8Text(0))

    Else

        TextPtr = 0

    End If

    ' --------------------------------------------------------
    ' 2. Create Native Response
    ' --------------------------------------------------------
    ResponseHandle = _
        vtnode_httpa_response_create_with_text( _
            TextPtr, _
            0)

    If ResponseHandle = 0 Then

        Debug.Print _
            "[VtHttpaResponseCreateWithText] " & _
            "Failed to create text response."

        Exit Function

    End If

    VtHttpaResponseCreateWithText = ResponseHandle
    Exit Function

EH:

    Debug.Print _
        "[VtHttpaResponseCreateWithText] Error " & _
        Err.Number & ": " & _
        Err.Description

    VtHttpaResponseCreateWithText = 0

End Function



''' <summary>
''' Multipart HTTP Response 핸들을 생성합니다.
''' </summary>
''' <param name="MultipartSubType">
''' Multipart subtype. 예: mixed, form-data, alternative
''' </param>
''' <returns>생성된 Response 핸들. 실패 시 0</returns>
Public Function VtHttpaResponseCreateAsMultipart( _
    Optional ByVal MultipartSubtype As String = "mixed" _
) As LongPtr

    Dim utf8SubType() As Byte
    Dim ResponseHandle As LongPtr

    On Error GoTo EH

    VtHttpaResponseCreateAsMultipart = 0

    If Len(Trim$(MultipartSubtype)) = 0 Then
        MultipartSubtype = "mixed"
    End If

    utf8SubType = _
        VtCharUtils.VtStringToUtf8Bytes(MultipartSubtype)

    ResponseHandle = _
        vtnode_httpa_response_create_as_multipart( _
            VarPtr(utf8SubType(0)))

    If ResponseHandle = 0 Then
        Debug.Print _
            "[VtHttpaResponseCreateAsMultipart] " & _
            "Failed to create multipart response."
    End If

    VtHttpaResponseCreateAsMultipart = ResponseHandle
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaResponseCreateAsMultipart] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseCreateAsMultipart = 0

End Function


''' <summary>
''' Response 핸들을 해제하고 변수 값을 0으로 초기화합니다.
''' </summary>
''' <param name="Response">
''' 해제할 Response 핸들. 해제 후 0으로 변경됨
''' </param>
Public Sub VtHttpaResponseDestroy( _
    ByRef Response As LongPtr _
)

    On Error GoTo EH

    If Response = 0 Then
        Exit Sub
    End If

    vtnode_httpa_response_destroy Response

    Response = 0
    Exit Sub

EH:
    Debug.Print _
        "[VtHttpaResponseDestroy] Error " & _
        Err.Number & ": " & Err.Description

End Sub





' ==============================================================================
' HTTP RESPONSE STATUS / VERSION WRAPPER
' ==============================================================================

''' <summary>
''' HTTP Response 상태 코드를 설정합니다.
''' </summary>
Public Sub VtHttpaResponseSetStatusCode( _
    ByVal Response As LongPtr, _
    ByVal StatusCode As Long _
)

    If Response = 0 Then
        Exit Sub
    End If

    vtnode_httpa_response_set_status_code _
        Response, _
        StatusCode

End Sub


''' <summary>
''' HTTP Response Reason Phrase를 설정합니다.
''' 예: "OK", "Not Found", "Internal Server Error"
''' </summary>
Public Sub VtHttpaResponseSetReasonPhrase( _
    ByVal Response As LongPtr, _
    ByVal Reason As String _
)

    Dim utf8Reason() As Byte

    If Response = 0 Then
        Exit Sub
    End If

    utf8Reason = _
        VtCharUtils.VtStringToUtf8Bytes(Reason)

    vtnode_httpa_response_set_reason_phrase _
        Response, _
        VarPtr(utf8Reason(0))

End Sub


''' <summary>
''' HTTP Response 버전을 설정합니다.
''' 예: "HTTP/1.1"
''' </summary>
Public Sub VtHttpaResponseSetVersion( _
    ByVal Response As LongPtr, _
    ByVal Version As String _
)

    Dim utf8Version() As Byte

    If Response = 0 Then
        Exit Sub
    End If

    If Len(Trim$(Version)) = 0 Then
        Version = "HTTP/1.1"
    End If

    utf8Version = _
        VtCharUtils.VtStringToUtf8Bytes(Version)

    vtnode_httpa_response_set_version _
        Response, _
        VarPtr(utf8Version(0))

End Sub


''' <summary>
''' HTTP Response Line 전체를 설정합니다.
''' 예: "HTTP/1.1 200 OK"
''' </summary>
Public Sub VtHttpaResponseSetResponseLine( _
    ByVal Response As LongPtr, _
    ByVal ResponseLine As String _
)

    Dim utf8ResponseLine() As Byte

    If Response = 0 Then
        Exit Sub
    End If

    If Len(Trim$(ResponseLine)) = 0 Then
        Exit Sub
    End If

    utf8ResponseLine = _
        VtCharUtils.VtStringToUtf8Bytes(ResponseLine)

    vtnode_httpa_response_set_response_line _
        Response, _
        VarPtr(utf8ResponseLine(0))

End Sub


''' <summary>
''' HTTP Response 상태 코드를 가져옵니다.
''' 핸들이 유효하지 않으면 0을 반환합니다.
''' </summary>
Public Function VtHttpaResponseGetStatusCode( _
    ByVal Response As LongPtr _
) As Long

    If Response = 0 Then
        VtHttpaResponseGetStatusCode = 0
        Exit Function
    End If

    VtHttpaResponseGetStatusCode = _
        vtnode_httpa_response_get_status_code(Response)

End Function


''' <summary>
''' HTTP Response Reason Phrase를 가져옵니다.
''' 반환 문자열은 Native Response 객체가 소유하므로 즉시 VBA String으로 복사합니다.
''' </summary>
Public Function VtHttpaResponseGetReasonPhrase( _
    ByVal Response As LongPtr _
) As String

    Dim pReason As LongPtr

    If Response = 0 Then
        VtHttpaResponseGetReasonPhrase = vbNullString
        Exit Function
    End If

    pReason = _
        vtnode_httpa_response_get_reason_phrase(Response)

    If pReason = 0 Then
        VtHttpaResponseGetReasonPhrase = vbNullString
    Else
        VtHttpaResponseGetReasonPhrase = _
            VtCharUtils.VtUtf8PtrToString(pReason)
    End If

End Function


''' <summary>
''' HTTP Response 버전을 가져옵니다.
''' 예: "HTTP/1.1"
''' </summary>
Public Function VtHttpaResponseGetVersion( _
    ByVal Response As LongPtr _
) As String

    Dim pVersion As LongPtr

    If Response = 0 Then
        VtHttpaResponseGetVersion = vbNullString
        Exit Function
    End If

    pVersion = _
        vtnode_httpa_response_get_version(Response)

    If pVersion = 0 Then
        VtHttpaResponseGetVersion = vbNullString
    Else
        VtHttpaResponseGetVersion = _
            VtCharUtils.VtUtf8PtrToString(pVersion)
    End If

End Function


''' <summary>
''' HTTP Response Line 전체를 가져옵니다.
''' 예: "HTTP/1.1 200 OK"
''' </summary>
Public Function VtHttpaResponseGetResponseLine( _
    ByVal Response As LongPtr _
) As String

    Dim pResponseLine As LongPtr

    If Response = 0 Then
        VtHttpaResponseGetResponseLine = vbNullString
        Exit Function
    End If

    pResponseLine = _
        vtnode_httpa_response_get_response_line(Response)

    If pResponseLine = 0 Then
        VtHttpaResponseGetResponseLine = vbNullString
    Else
        VtHttpaResponseGetResponseLine = _
            VtCharUtils.VtUtf8PtrToString(pResponseLine)
    End If

End Function




' ==============================================================================
' HTTP RESPONSE HEADER WRAPPER
' ==============================================================================

''' <summary>
''' HTTP Response에 헤더 값을 추가합니다.
''' 동일한 이름의 헤더가 이미 존재해도 새 값을 추가할 수 있습니다.
''' </summary>
Public Function VtHttpaResponseAddHeader( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As String, _
    ByVal HeaderValue As String _
) As Boolean

    Dim utf8HeaderName() As Byte
    Dim utf8HeaderValue() As Byte
    Dim nativeResult As Long

    On Error GoTo EH

    VtHttpaResponseAddHeader = False

    If Response = 0 Then
        Exit Function
    End If

    If Len(Trim$(HeaderName)) = 0 Then
        Exit Function
    End If

    utf8HeaderName = _
        VtCharUtils.VtStringToUtf8Bytes(HeaderName)

    utf8HeaderValue = _
        VtCharUtils.VtStringToUtf8Bytes(HeaderValue)

    nativeResult = _
        vtnode_httpa_response_add_header( _
            Response, _
            VarPtr(utf8HeaderName(0)), _
            VarPtr(utf8HeaderValue(0)))

    VtHttpaResponseAddHeader = (nativeResult <> 0)
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaResponseAddHeader] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseAddHeader = False

End Function


''' <summary>
''' HTTP Response 헤더 값을 설정합니다.
''' 동일한 이름의 헤더가 이미 존재하면 기존 값을 교체합니다.
''' </summary>
Public Function VtHttpaResponseSetHeader( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As String, _
    ByVal HeaderValue As String _
) As Boolean

    Dim utf8HeaderName() As Byte
    Dim utf8HeaderValue() As Byte
    Dim nativeResult As Long

    On Error GoTo EH

    VtHttpaResponseSetHeader = False

    If Response = 0 Then
        Exit Function
    End If

    If Len(Trim$(HeaderName)) = 0 Then
        Exit Function
    End If

    utf8HeaderName = _
        VtCharUtils.VtStringToUtf8Bytes(HeaderName)

    utf8HeaderValue = _
        VtCharUtils.VtStringToUtf8Bytes(HeaderValue)

    nativeResult = _
        vtnode_httpa_response_set_header( _
            Response, _
            VarPtr(utf8HeaderName(0)), _
            VarPtr(utf8HeaderValue(0)))

    VtHttpaResponseSetHeader = (nativeResult <> 0)
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaResponseSetHeader] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseSetHeader = False

End Function


''' <summary>
''' HTTP Response에서 지정한 이름의 헤더를 제거합니다.
''' 헤더가 존재하지 않으면 아무 작업도 하지 않습니다.
''' </summary>
Public Sub VtHttpaResponseRemoveHeader( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As String _
)

    Dim utf8HeaderName() As Byte

    On Error GoTo EH

    If Response = 0 Then
        Exit Sub
    End If

    If Len(Trim$(HeaderName)) = 0 Then
        Exit Sub
    End If

    utf8HeaderName = _
        VtCharUtils.VtStringToUtf8Bytes(HeaderName)

    vtnode_httpa_response_remove_header _
        Response, _
        VarPtr(utf8HeaderName(0))

    Exit Sub

EH:
    Debug.Print _
        "[VtHttpaResponseRemoveHeader] Error " & _
        Err.Number & ": " & Err.Description

End Sub


''' <summary>
''' HTTP Response에서 지정한 이름의 헤더 값을 가져옵니다.
''' 헤더가 없거나 핸들이 유효하지 않으면 vbNullString을 반환합니다.
''' </summary>
Public Function VtHttpaResponseGetHeader( _
    ByVal Response As LongPtr, _
    ByVal HeaderName As String _
) As String

    Dim utf8HeaderName() As Byte
    Dim pHeaderValue As LongPtr

    On Error GoTo EH

    VtHttpaResponseGetHeader = vbNullString

    If Response = 0 Then
        Exit Function
    End If

    If Len(Trim$(HeaderName)) = 0 Then
        Exit Function
    End If

    utf8HeaderName = _
        VtCharUtils.VtStringToUtf8Bytes(HeaderName)

    pHeaderValue = _
        vtnode_httpa_response_get_header( _
            Response, _
            VarPtr(utf8HeaderName(0)))

    If pHeaderValue <> 0 Then
        VtHttpaResponseGetHeader = _
            VtCharUtils.VtUtf8PtrToString(pHeaderValue)
    End If

    Exit Function

EH:
    Debug.Print _
        "[VtHttpaResponseGetHeader] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseGetHeader = vbNullString

End Function


''' <summary>
''' 1-based 인덱스로 HTTP Response 헤더 이름과 값을 가져옵니다.
''' 성공하면 True를 반환하고 HeaderNameOut, HeaderValueOut에 값을 저장합니다.
''' 실패하면 False를 반환하고 출력 문자열을 비웁니다.
''' </summary>
Public Function VtHttpaResponseGetHeaderAt( _
    ByVal Response As LongPtr, _
    ByVal index As Long, _
    ByRef HeaderNameOut As String, _
    ByRef HeaderValueOut As String _
) As Boolean

    Dim pHeaderName As LongPtr
    Dim pHeaderValue As LongPtr
    Dim nativeResult As Long

    On Error GoTo EH

    HeaderNameOut = vbNullString
    HeaderValueOut = vbNullString
    VtHttpaResponseGetHeaderAt = False

    If Response = 0 Then
        Exit Function
    End If

    If index < 1 Then
        Exit Function
    End If

    nativeResult = _
        vtnode_httpa_response_get_header_at( _
            Response, _
            index, _
            pHeaderName, _
            pHeaderValue)

    If nativeResult = 0 Then
        Exit Function
    End If

    If pHeaderName <> 0 Then
        HeaderNameOut = _
            VtCharUtils.VtUtf8PtrToString(pHeaderName)
    End If

    If pHeaderValue <> 0 Then
        HeaderValueOut = _
            VtCharUtils.VtUtf8PtrToString(pHeaderValue)
    End If

    VtHttpaResponseGetHeaderAt = True
    Exit Function

EH:
    HeaderNameOut = vbNullString
    HeaderValueOut = vbNullString

    Debug.Print _
        "[VtHttpaResponseGetHeaderAt] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseGetHeaderAt = False

End Function







''' <summary>
''' Response의 지정된 1-based Index에 해당하는 Content Handle을 반환합니다.
''' Response가 유효하지 않거나 Index가 범위를 벗어나면 0을 반환합니다.
''' </summary>
Public Function VtHttpaResponseGetContentAt( _
    ByVal Response As LongPtr, _
    ByVal index As Long _
) As LongPtr

    Dim ContentHandle As LongPtr

    On Error GoTo EH

    VtHttpaResponseGetContentAt = 0

    If Response = 0 Then
        Exit Function
    End If

    If index < 1 Then
        Exit Function
    End If

    ContentHandle = _
        vtnode_httpa_response_get_content_at(Response, index)

    VtHttpaResponseGetContentAt = ContentHandle
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaResponseGetContentAt] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseGetContentAt = 0

End Function


''' <summary>
''' Multipart Response에 파일 Content를 추가하고 Content Handle을 반환합니다.
''' 실패하면 0을 반환합니다.
''' </summary>
Public Function VtHttpaResponseAddContentFile( _
    ByVal Response As LongPtr, _
    ByVal FilePath As String, _
    Optional ByVal ContentType As String = "application/octet-stream", _
    Optional ByVal ContentDispositionName As String = vbNullString, _
    Optional ByVal ContentDispositionFilename As String = vbNullString _
) As LongPtr

    Dim utf8FilePath() As Byte
    Dim utf8ContentType() As Byte
    Dim utf8DispositionName() As Byte
    Dim utf8DispositionFilename() As Byte

    Dim pFilePath As LongPtr
    Dim pContentType As LongPtr
    Dim pDispositionName As LongPtr
    Dim pDispositionFilename As LongPtr

    Dim ContentHandle As LongPtr

    On Error GoTo EH

    VtHttpaResponseAddContentFile = 0

    ' Validate Response handle.
    If Response = 0 Then
        Debug.Print _
            "[VtHttpaResponseAddContentFile] Invalid Response handle."
        Exit Function
    End If

    ' Validate file path.
    If Len(Trim$(FilePath)) = 0 Then
        Debug.Print _
            "[VtHttpaResponseAddContentFile] FilePath is empty."
        Exit Function
    End If

    ' Verify that the file exists.
    If Len(dir$( _
        FilePath, _
        vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then

        Debug.Print _
            "[VtHttpaResponseAddContentFile] File not found: " & _
            FilePath

        Exit Function
    End If

    ' Apply default Content-Type.
    If Len(Trim$(ContentType)) = 0 Then
        ContentType = "application/octet-stream"
    End If

    ' Convert required parameters to null-terminated UTF-8.
    utf8FilePath = _
        VtCharUtils.VtStringToUtf8Bytes(FilePath)

    utf8ContentType = _
        VtCharUtils.VtStringToUtf8Bytes(ContentType)

    pFilePath = VarPtr(utf8FilePath(0))
    pContentType = VarPtr(utf8ContentType(0))

    ' Convert optional Content-Disposition name.
    If Len(ContentDispositionName) > 0 Then

        utf8DispositionName = _
            VtCharUtils.VtStringToUtf8Bytes( _
                ContentDispositionName)

        pDispositionName = _
            VarPtr(utf8DispositionName(0))

    Else
        pDispositionName = 0
    End If

    ' Convert optional Content-Disposition filename.
    If Len(ContentDispositionFilename) > 0 Then

        utf8DispositionFilename = _
            VtCharUtils.VtStringToUtf8Bytes( _
                ContentDispositionFilename)

        pDispositionFilename = _
            VarPtr(utf8DispositionFilename(0))

    Else
        pDispositionFilename = 0
    End If

    ' Add file content and receive Content Handle.
    ContentHandle = _
        vtnode_httpa_response_content_add_file( _
            Response, _
            pFilePath, _
            pContentType, _
            pDispositionName, _
            pDispositionFilename)

    If ContentHandle = 0 Then
        Debug.Print _
            "[VtHttpaResponseAddContentFile] " & _
            "Failed to add file content: " & _
            FilePath

        Exit Function
    End If

    VtHttpaResponseAddContentFile = ContentHandle
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaResponseAddContentFile] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseAddContentFile = 0

End Function



''' <summary>
''' Multipart Response에 Text Content를 추가하고 Content Handle을 반환합니다.
''' 실패하면 0을 반환합니다.
''' </summary>
Public Function VtHttpaResponseAddContentText( _
    ByVal Response As LongPtr, _
    ByVal text As String, _
    Optional ByVal ContentType As String = _
        "text/plain; charset=utf-8", _
    Optional ByVal ContentDispositionName As String = _
        vbNullString _
) As LongPtr

    Dim utf8Text() As Byte
    Dim utf8ContentType() As Byte
    Dim utf8DispositionName() As Byte

    Dim pText As LongPtr
    Dim pContentType As LongPtr
    Dim pDispositionName As LongPtr

    Dim ContentHandle As LongPtr

    On Error GoTo EH

    VtHttpaResponseAddContentText = 0

    ' Validate Response handle.
    If Response = 0 Then
        Debug.Print _
            "[VtHttpaResponseAddContentText] Invalid Response handle."
        Exit Function
    End If

    ' Apply default Content-Type.
    If Len(Trim$(ContentType)) = 0 Then
        ContentType = "text/plain; charset=utf-8"
    End If

    ' Convert Text to a null-terminated UTF-8 byte array.
    utf8Text = _
        VtCharUtils.VtStringToUtf8Bytes(text)

    pText = VarPtr(utf8Text(0))

    ' Convert Content-Type to a null-terminated UTF-8 byte array.
    utf8ContentType = _
        VtCharUtils.VtStringToUtf8Bytes(ContentType)

    pContentType = VarPtr(utf8ContentType(0))

    ' Convert optional Content-Disposition name.
    If Len(ContentDispositionName) > 0 Then

        utf8DispositionName = _
            VtCharUtils.VtStringToUtf8Bytes( _
                ContentDispositionName)

        pDispositionName = _
            VarPtr(utf8DispositionName(0))

    Else
        pDispositionName = 0
    End If

    ' Add text content and receive Content Handle.
    ContentHandle = _
        vtnode_httpa_response_content_add_text( _
            Response, _
            pText, _
            pContentType, _
            pDispositionName)

    If ContentHandle = 0 Then
        Debug.Print _
            "[VtHttpaResponseAddContentText] " & _
            "Failed to add text content."

        Exit Function
    End If

    VtHttpaResponseAddContentText = ContentHandle
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaResponseAddContentText] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaResponseAddContentText = 0

End Function

