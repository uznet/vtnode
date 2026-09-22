Attribute VB_Name = "VtHttpaRequest"
Option Explicit

#If VBA7 And Win64 Then

#If VTNODE_DEBUG = 1 Then  ' Debug Mode


' Create a new request handle from an event.
Private Declare PtrSafe Function vtnode_httpa_request_create_from_event Lib "vtnode64d.dll" ( _
    ByVal EventId As LongLong _
) As LongPtr

' Create a new empty request handle.
Private Declare PtrSafe Function vtnode_httpa_request_create Lib "vtnode64d.dll" () As LongPtr

' Create a new request handle with file content.
' FilePath is the path to the file to be sent as request body (null-terminated UTF-8 string).
' ContentType is the MIME type of the file, e.g. "application/json", "text/plain", etc.
Private Declare PtrSafe Function vtnode_httpa_request_create_with_file Lib "vtnode64d.dll" ( _
    ByVal FilePath As LongPtr, _
    ByVal ContentType As LongPtr _
) As LongPtr

' Create a new request handle with text content (null-terminated UTF-8 string).
' ContentType is the MIME type of the text, e.g. "application/json", "text/plain", etc.
' Text is the text content to be sent as request body.
Private Declare PtrSafe Function vtnode_httpa_request_create_with_text Lib "vtnode64d.dll" ( _
    ByVal text As LongPtr, _
    ByVal ContentType As LongPtr _
) As LongPtr

' Create a new request handle as multipart/form-data or multipart/mixed request.
' MultipartSubType is the subtype of multipart, e.g. "form-data" or "mixed".
Private Declare PtrSafe Function vtnode_httpa_request_create_as_multipart Lib "vtnode64d.dll" ( _
    ByVal MultipartSubtype As LongPtr _
) As LongPtr

' Destroy a request handle.
Private Declare PtrSafe Sub vtnode_httpa_request_destroy Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr _
)

' url is the full URL including scheme, host, port, path, and query string.
Private Declare PtrSafe Sub vtnode_httpa_request_set_fullurl Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal url As LongPtr _
)

' uri is the path and query string part of the URL, e.g. "/path/to/resource?query=param"
Private Declare PtrSafe Sub vtnode_httpa_request_set_uri Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal Uri As LongPtr _
)

' path is the path part of the URL, e.g. "/path/to/resource"
Private Declare PtrSafe Sub vtnode_httpa_request_set_path Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal Path As LongPtr _
)

' query is the query string part of the URL, e.g. "query=param"
Private Declare PtrSafe Sub vtnode_httpa_request_set_query Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal Query As LongPtr _
)

' method is the HTTP method, e.g. "GET", "POST"
Private Declare PtrSafe Sub vtnode_httpa_request_set_method Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal Method As LongPtr _
)

' version is the HTTP version, e.g. "HTTP/1.1"
Private Declare PtrSafe Sub vtnode_httpa_request_set_version Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal Version As LongPtr _
)

' Get request line (e.g. "GET /path/to/resource HTTP/1.1").
Private Declare PtrSafe Function vtnode_httpa_request_get_request_line Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get request method (e.g. "GET", "POST").
Private Declare PtrSafe Function vtnode_httpa_request_get_method Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get request version (e.g. "HTTP/1.1").
Private Declare PtrSafe Function vtnode_httpa_request_get_version Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get full URL (e.g. "http://example.com/path/to/resource?query=param").
Private Declare PtrSafe Function vtnode_httpa_request_get_fullurl Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get request URI (e.g. "/path/to/resource?query=param").
Private Declare PtrSafe Function vtnode_httpa_request_get_uri Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get request path (e.g. "/path/to/resource").
Private Declare PtrSafe Function vtnode_httpa_request_get_path Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get request query string (e.g. "query=param").
Private Declare PtrSafe Function vtnode_httpa_request_get_query Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr


' Add or set header value by header name.
Private Declare PtrSafe Function vtnode_httpa_request_add_header Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Set header value by header name. If header already exists, it will be replaced.
Private Declare PtrSafe Function vtnode_httpa_request_set_header Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Remove header by header name. If header does not exist, do nothing.
Private Declare PtrSafe Sub vtnode_httpa_request_remove_header Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal HeaderName As LongPtr _
)

' Get header value by header name.
' Return NULL if not found.
' Returned string is read-only UTF-8 and owned by request.
Private Declare PtrSafe Function vtnode_httpa_request_get_header Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal HeaderName As LongPtr _
) As LongPtr

' Get header by 1-base index.
' Return FALSE if not found.
' Returned string is read-only UTF-8 and owned by request.
Private Declare PtrSafe Function vtnode_httpa_request_get_header_at Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal index As Long, _
    ByRef HeaderName As LongPtr, _
    ByRef HeaderValue As LongPtr _
) As Long

' Get request contents count (multipart/form-data etc).
Private Declare PtrSafe Function vtnode_httpa_request_content_get_count Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr _
) As Long

' Get content data by 1-based index. Return FALSE if not found.
' outData is a pointer to the content data.
' If the content is file-backed, outData will point to the file path string.
' If the content is memory-backed, outData will point to the memory buffer.
Private Declare PtrSafe Function vtnode_httpa_request_content_get_at Lib "vtnode64d.dll" (ByVal Request As LongPtr, ByVal index As Long) As LongPtr


' Add a new content part to the request.
' Request must be created by vtnode_httpa_request_create_as_multipart().
Private Declare PtrSafe Function vtnode_httpa_request_content_add_file Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal FilePath As LongPtr, _
    ByVal ContentType As LongPtr, _
    ByVal ContentDispositionName As LongPtr, _
    ByVal ContentDispositionFilename As LongPtr _
) As LongPtr

' Add a new content part to the request with text content.
' Request must be created by vtnode_httpa_request_create_as_multipart().
Private Declare PtrSafe Function vtnode_httpa_request_content_add_text Lib "vtnode64d.dll" ( _
    ByVal Request As LongPtr, _
    ByVal text As LongPtr, _
    ByVal ContentType As LongPtr, _
    ByVal ContentDispositionName As LongPtr _
) As LongPtr



#Else  ' Release mode


' Create a new request handle from an event.
Private Declare PtrSafe Function vtnode_httpa_request_create_from_event Lib "vtnode64.dll" ( _
    ByVal EventId As LongLong _
) As LongPtr

' Create a new empty request handle.
Private Declare PtrSafe Function vtnode_httpa_request_create Lib "vtnode64.dll" () As LongPtr

' Create a new request handle with file content.
' FilePath is the path to the file to be sent as request body (null-terminated UTF-8 string).
' ContentType is the MIME type of the file, e.g. "application/json", "text/plain", etc.
Private Declare PtrSafe Function vtnode_httpa_request_create_with_file Lib "vtnode64.dll" ( _
    ByVal FilePath As LongPtr, _
    ByVal ContentType As LongPtr _
) As LongPtr

' Create a new request handle with text content (null-terminated UTF-8 string).
' ContentType is the MIME type of the text, e.g. "application/json", "text/plain", etc.
' Text is the text content to be sent as request body.
Private Declare PtrSafe Function vtnode_httpa_request_create_with_text Lib "vtnode64.dll" ( _
    ByVal text As LongPtr, _
    ByVal ContentType As LongPtr _
) As LongPtr

' Create a new request handle as multipart/form-data or multipart/mixed request.
' MultipartSubType is the subtype of multipart, e.g. "form-data" or "mixed".
Private Declare PtrSafe Function vtnode_httpa_request_create_as_multipart Lib "vtnode64.dll" ( _
    ByVal MultipartSubtype As LongPtr _
) As LongPtr

' Destroy a request handle.
Private Declare PtrSafe Sub vtnode_httpa_request_destroy Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr _
)

' url is the full URL including scheme, host, port, path, and query string.
Private Declare PtrSafe Sub vtnode_httpa_request_set_fullurl Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal url As LongPtr _
)

' uri is the path and query string part of the URL, e.g. "/path/to/resource?query=param"
Private Declare PtrSafe Sub vtnode_httpa_request_set_uri Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal Uri As LongPtr _
)

' path is the path part of the URL, e.g. "/path/to/resource"
Private Declare PtrSafe Sub vtnode_httpa_request_set_path Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal Path As LongPtr _
)

' query is the query string part of the URL, e.g. "query=param"
Private Declare PtrSafe Sub vtnode_httpa_request_set_query Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal Query As LongPtr _
)

' method is the HTTP method, e.g. "GET", "POST"
Private Declare PtrSafe Sub vtnode_httpa_request_set_method Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal Method As LongPtr _
)

' version is the HTTP version, e.g. "HTTP/1.1"
Private Declare PtrSafe Sub vtnode_httpa_request_set_version Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal Version As LongPtr _
)

' Get request line (e.g. "GET /path/to/resource HTTP/1.1").
Private Declare PtrSafe Function vtnode_httpa_request_get_request_line Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get request method (e.g. "GET", "POST").
Private Declare PtrSafe Function vtnode_httpa_request_get_method Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get request version (e.g. "HTTP/1.1").
Private Declare PtrSafe Function vtnode_httpa_request_get_version Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get full URL (e.g. "http://example.com/path/to/resource?query=param").
Private Declare PtrSafe Function vtnode_httpa_request_get_fullurl Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get request URI (e.g. "/path/to/resource?query=param").
Private Declare PtrSafe Function vtnode_httpa_request_get_uri Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get request path (e.g. "/path/to/resource").
Private Declare PtrSafe Function vtnode_httpa_request_get_path Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr

' Get request query string (e.g. "query=param").
Private Declare PtrSafe Function vtnode_httpa_request_get_query Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr _
) As LongPtr


' Add or set header value by header name.
Private Declare PtrSafe Function vtnode_httpa_request_add_header Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Set header value by header name. If header already exists, it will be replaced.
Private Declare PtrSafe Function vtnode_httpa_request_set_header Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Remove header by header name. If header does not exist, do nothing.
Private Declare PtrSafe Sub vtnode_httpa_request_remove_header Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal HeaderName As LongPtr _
)

' Get header value by header name.
' Return NULL if not found.
' Returned string is read-only UTF-8 and owned by request.
Private Declare PtrSafe Function vtnode_httpa_request_get_header Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal HeaderName As LongPtr _
) As LongPtr

' Get header by 1-base index.
' Return FALSE if not found.
' Returned string is read-only UTF-8 and owned by request.
Private Declare PtrSafe Function vtnode_httpa_request_get_header_at Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal index As Long, _
    ByRef HeaderName As LongPtr, _
    ByRef HeaderValue As LongPtr _
) As Long

' Get request contents count (multipart/form-data etc).
Private Declare PtrSafe Function vtnode_httpa_request_content_get_count Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr _
) As Long

' Get content data by 1-based index. Return FALSE if not found.
' outData is a pointer to the content data.
' If the content is file-backed, outData will point to the file path string.
' If the content is memory-backed, outData will point to the memory buffer.
Private Declare PtrSafe Function vtnode_httpa_request_content_get_at Lib "vtnode64.dll" (ByVal Request As LongPtr, ByVal index As Long) As LongPtr

' Add a new content part to the request.
' Request must be created by vtnode_httpa_request_create_as_multipart().
Private Declare PtrSafe Function vtnode_httpa_request_content_add_file Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal FilePath As LongPtr, _
    ByVal ContentType As LongPtr, _
    ByVal ContentDispositionName As LongPtr, _
    ByVal ContentDispositionFilename As LongPtr _
) As LongPtr

' Add a new content part to the request with text content.
' Request must be created by vtnode_httpa_request_create_as_multipart().
Private Declare PtrSafe Function vtnode_httpa_request_content_add_text Lib "vtnode64.dll" ( _
    ByVal Request As LongPtr, _
    ByVal text As LongPtr, _
    ByVal ContentType As LongPtr, _
    ByVal ContentDispositionName As LongPtr _
) As LongPtr


#End If

#End If



' =========================================================================
' 2. VBA Wrapper 함수 (응용 레이어 인터페이스)
' =========================================================================
''' <summary>
''' EventId로부터 HTTP Request 핸들(포인터)을 생성합니다.
''' </summary>
''' <param name="eventId">WAS/엔진 이벤트 핸들러로부터 전달받은 64비트 Event ID</param>
''' <returns>생성된 HTTP Request 객체 포인터 (실패 시 0 / NullPtr)</returns>
Public Function VtHttpaRequestCreateFromEvent(ByVal EventId As LongLong) As LongPtr
    Dim hRequest As LongPtr
    
    ' 1. 유효성 검사 (0 또는 음수 eventId 예외 처리)
    If EventId = 0 Then
        Debug.Print "[VtHttaRequestCreateFromEvent] Error: Invalid EventId (" & EventId & ")"
        VtHttpaRequestCreateFromEvent = 0
        Exit Function
    End If
    
    ' 2. Native DLL 함수 호출
    On Error GoTo EH
    hRequest = vtnode_httpa_request_create_from_event(EventId)
    
    ' 3. 반환 포인터 검증 및 로깅
    If hRequest = 0 Then
        Debug.Print "[VtHttaRequestCreateFromEvent] Warning: Failed to create Request handle from EventId (" & EventId & ")"
    Else
        ' 성공 시 획득한 포인터 반환
        VtHttpaRequestCreateFromEvent = hRequest
    End If
    
    Exit Function

EH:
    Debug.Print "[VtHttaRequestCreateFromEvent] Critical Error: " & Err.Description & " (Num: " & Err.Number & ")"
    VtHttpaRequestCreateFromEvent = 0
End Function



Public Function VtHttpaRequestCreateWithFile(ByVal FilePath As String, ByVal ContentType As String) As LongPtr

    Dim hRequest As LongPtr
    Dim u8FilePath() As Byte
    Dim u8ContentType() As Byte
    
    
    On Error GoTo EH
    
    ' 1. Guard Clause: 경로 유효성 검사
    If Len(Trim$(FilePath)) = 0 Then
        Debug.Print "[VtHttpaRequestCreateWithFile] Error: File path is empty."
        VtHttpaRequestCreateWithFile = 0
        Exit Function
    End If
    
    ' 2. VBA String -> UTF-8 Byte Array 변환 (C DLL과의 호환성)
    u8FilePath = VtCharUtils.VtStringToUtf8Bytes(FilePath)
    
    If Len(Trim$(ContentType)) = 0 Then
        ' Content-Type이 빈 문자열인 경우 기본값 처리 또는 Null 전달
        u8ContentType = VtCharUtils.VtStringToUtf8Bytes("application/octet-stream")
    Else
        u8ContentType = VtCharUtils.VtStringToUtf8Bytes(ContentType)
    End If
    
    
    
    ' 3. Native DLL 호출 (바이트 배열의 첫 번째 요소 주소 전달)
    hRequest = vtnode_httpa_request_create_with_file( _
        VarPtr(u8FilePath(0)), _
        VarPtr(u8ContentType(0)) _
    )
    
    ' 4. 반환값 검증 및 반환
    If hRequest = 0 Then
        Debug.Print "[VtHttpaRequestCreateWithFile] Failed to create HTTP request object from file: " & FilePath
    End If
    
    VtHttpaRequestCreateWithFile = hRequest
    Exit Function

EH:
    Debug.Print "[VtHttpaRequestCreateWithFile] Exception: " & Err.Description
    VtHttpaRequestCreateWithFile = 0
    
End Function


' -----------------------------------------------------------------------------
' Wrapper: VtHttpaRequestCreateWithText
' -----------------------------------------------------------------------------
''' <summary>
''' 텍스트 본문(Text)과 Content-Type을 받아 HTTP Request 핸들을 생성합니다.
''' </summary>
''' <param name="Text">전송할 텍스트 본문 (예: JSON 문자열, 일반 HTML/Text 등)</param>
''' <param name="ContentType">미디어 타입 (기본값: "text/plain; charset=utf-8")</param>
''' <returns>생성된 Request 객체 포인터(LongPtr). 실패 시 0(Null Pointer) 반환</returns>
Public Function VtHttpaRequestCreateWithText( _
    ByVal text As String, _
    Optional ByVal ContentType As String = "text/plain; charset=utf-8" _
) As LongPtr

    On Error GoTo EH
    
    ' 1. 기본값 널 체크 핸들링
    If Len(Trim$(ContentType)) = 0 Then
        ContentType = "text/plain; charset=utf-8"
    End If
    
    ' 2. VBA Unicode String -> UTF-8 Byte Array 인코딩 변환
    Dim textBytes() As Byte
    Dim contentTypeBytes() As Byte
    
    textBytes = VtCharUtils.VtStringToUtf8Bytes(text)
    contentTypeBytes = VtCharUtils.VtStringToUtf8Bytes(ContentType)
    
    ' 3. DLL C API 호출 (포인터 전달)
    Dim hRequest As LongPtr
    hRequest = vtnode_httpa_request_create_with_text( _
        VarPtr(textBytes(0)), _
        VarPtr(contentTypeBytes(0)) _
    )
    
    ' 4. 결과 검증
    If hRequest = 0 Then
        Debug.Print "[ERR] VtHttpaRequestCreateWithText: Failed to create request handle from DLL."
    End If
    
    VtHttpaRequestCreateWithText = hRequest
    Exit Function

EH:
    Debug.Print "[ERR] VtHttpaRequestCreateWithText Exception: " & Err.Description
    VtHttpaRequestCreateWithText = 0
End Function
' ----------------------------------------------------------------
' [Wrapper] VtHttpaRequestCreateAsMultipart
' ----------------------------------------------------------------
''' <summary>
''' Multipart 형태의 HTTP Request 객체를 생성합니다.
''' </summary>
''' <param name="subType">Multipart 세부 타입 (기본값: "form-data", 필요시 "mixed", "alternative" 등 가능)</param>
''' <returns>생성된 Request 객체 핸들/포인터 (실패 시 0)</returns>
Public Function VtHttpaRequestCreateAsMultipart(Optional ByVal subType As String = "form-data") As LongPtr
    On Error GoTo EH
    
    ' 1. 입력 인자 검증 및 기본값 보장
    If Trim$(subType) = vbNullString Then
        subType = "form-data"
    End If
    
    ' 2. UTF-8 인코딩 변환 (Null-terminated Byte Array)
    Dim u8SubType() As Byte
    u8SubType = VtCharUtils.VtStringToUtf8Bytes(subType)
    
    ' 3. DLL C API 호출 (버퍼의 첫 번째 포인터 전달)
    Dim hReq As LongPtr
    hReq = vtnode_httpa_request_create_as_multipart(VarPtr(u8SubType(0)))
    
    ' 4. 결과 반환
    VtHttpaRequestCreateAsMultipart = hReq
    Exit Function

EH:
    Debug.Print "[ERR] VtHttpaRequestCreateAsMultipart Exception: " & Err.Description
    VtHttpaRequestCreateAsMultipart = 0
End Function

' ==============================================================================
' [Wrapper Sub] VtHttpaRequestDestroy
' ------------------------------------------------------------------------------
' 설명 : C/C++ DLL에서 할당한 Request 객체 메모리를 안전하게 해제합니다.
' 매개변수 :
'   - reqHandle : [ByRef] 해제할 Request 포인터 변수.
'                 (ByRef로 받아 해제 후 0(Null)으로 자동 초기화하여 Double-Free 방지)
' ==============================================================================
Public Sub VtHttpaRequestDestroy(ByRef reqHandle As LongPtr)
    On Error GoTo EH

    ' 1. Guard Clause : 포인터가 이미 Null(0)인 경우 처리 Skip
    If reqHandle = 0 Then
        ' Debug.Print "[VtHttpaRequestDestroy] Already NULL pointer. Skip."
        Exit Sub
    End If

    ' 2. DLL Native Sub 호출 (메모리 해제 수행)
    Call vtnode_httpa_request_destroy(reqHandle)

    ' 3. Dangling Pointer 및 Double Free 방지를 위한 포인터 초기화
    reqHandle = 0

    Exit Sub

EH:
    Debug.Print "[ERR][VtHttpaRequestDestroy] Error Number: " & Err.Number & " / " & Err.Description
End Sub



' ==========================================
' VBA Wrapper Subroutines (SETTERs)
' ==========================================

''' <summary>Set full URL (e.g., "http://example.com/path?query=1")</summary>
Public Sub VtHttpaRequestSetFullUrl(ByVal reqHandle As LongPtr, ByVal url As String)
    If reqHandle = 0 Then Exit Sub
    Dim u8Bytes() As Byte
    u8Bytes = StringToUtf8Bytes(url)
    Call vtnode_httpa_request_set_fullurl(reqHandle, VarPtr(u8Bytes(0)))
End Sub

''' <summary>Set URI path and query (e.g., "/path/to/resource?query=param")</summary>
Public Sub VtHttpaRequestSetUri(ByVal reqHandle As LongPtr, ByVal Uri As String)
    If reqHandle = 0 Then Exit Sub
    Dim u8Bytes() As Byte
    u8Bytes = StringToUtf8Bytes(Uri)
    Call vtnode_httpa_request_set_uri(reqHandle, VarPtr(u8Bytes(0)))
End Sub

''' <summary>Set Path only (e.g., "/path/to/resource")</summary>
Public Sub VtHttpaRequestSetPath(ByVal reqHandle As LongPtr, ByVal pathStr As String)
    If reqHandle = 0 Then Exit Sub
    Dim u8Bytes() As Byte
    u8Bytes = StringToUtf8Bytes(pathStr)
    Call vtnode_httpa_request_set_path(reqHandle, VarPtr(u8Bytes(0)))
End Sub

''' <summary>Set Query String (e.g., "query=param")</summary>
Public Sub VtHttpaRequestSetQuery(ByVal reqHandle As LongPtr, ByVal queryStr As String)
    If reqHandle = 0 Then Exit Sub
    Dim u8Bytes() As Byte
    u8Bytes = StringToUtf8Bytes(queryStr)
    Call vtnode_httpa_request_set_query(reqHandle, VarPtr(u8Bytes(0)))
End Sub

''' <summary>Set HTTP Method (e.g., "GET", "POST", "PUT")</summary>
Public Sub VtHttpaRequestSetMethod(ByVal reqHandle As LongPtr, ByVal MethodStr As String)
    If reqHandle = 0 Then Exit Sub
    Dim u8Bytes() As Byte
    u8Bytes = StringToUtf8Bytes(UCase$(MethodStr))
    Call vtnode_httpa_request_set_method(reqHandle, VarPtr(u8Bytes(0)))
End Sub

''' <summary>Set HTTP Version (e.g., "HTTP/1.1")</summary>
Public Sub VtHttpaRequestSetVersion(ByVal reqHandle As LongPtr, ByVal versionStr As String)
    If reqHandle = 0 Then Exit Sub
    Dim u8Bytes() As Byte
    u8Bytes = StringToUtf8Bytes(versionStr)
    Call vtnode_httpa_request_set_version(reqHandle, VarPtr(u8Bytes(0)))
End Sub


' ==========================================
' VBA Wrapper Functions (GETTERs)
' ==========================================

''' <summary>Get Request Line (e.g., "GET /path/to/resource HTTP/1.1")</summary>
Public Function VtHttpaRequestGetRequestLine(ByVal reqHandle As LongPtr) As String
    
    If reqHandle <> 0 Then
       VtHttpaRequestGetRequestLine = VtCharUtils.VtUtf8PtrToString(vtnode_httpa_request_get_request_line(reqHandle))
    Else
    
       VtHttpaRequestGetRequestLine = vbNullString
    End If
    
    

End Function

''' <summary>Get HTTP Method</summary>
Public Function VtHttpaRequestGetMethod(ByVal reqHandle As LongPtr) As String
    
    If reqHandle <> 0 Then
        VtHttpaRequestGetMethod = VtCharUtils.VtUtf8PtrToString(vtnode_httpa_request_get_method(reqHandle))
    Else
        VtHttpaRequestGetMethod = vbNullString
    End If
End Function

''' <summary>Get HTTP Version</summary>
Public Function VtHttpaRequestGetVersion(ByVal reqHandle As LongPtr) As String
    If reqHandle <> 0 Then
        VtHttpaRequestGetVersion = VtCharUtils.VtUtf8PtrToString(vtnode_httpa_request_get_version(reqHandle))
    Else
       VtHttpaRequestGetVersion = vbNullString
    End If
    
End Function

''' <summary>Get Full URL</summary>
Public Function VtHttpaRequestGetFullUrl(ByVal reqHandle As LongPtr) As String
    If reqHandle <> 0 Then
       VtHttpaRequestGetFullUrl = VtCharUtils.VtUtf8PtrToString(vtnode_httpa_request_get_fullurl(reqHandle))
    Else
       VtHttpaRequestGetFullUrl = vbNullString
      
    End If
End Function

''' <summary>Get Request URI</summary>
Public Function VtHttpaRequestGetUri(ByVal reqHandle As LongPtr) As String
    If reqHandle <> 0 Then
    VtHttpaRequestGetUri = VtCharUtils.VtUtf8PtrToString(vtnode_httpa_request_get_uri(reqHandle))
    Else
      VtHttpaRequestGetUri = vbNullString
    End If
End Function

''' <summary>Get Path</summary>
Public Function VtHttpaRequestGetPath(ByVal reqHandle As LongPtr) As String
    If reqHandle <> 0 Then
      VtHttpaRequestGetPath = VtCharUtils.VtUtf8PtrToString(vtnode_httpa_request_get_path(reqHandle))
    Else
      VtHttpaRequestGetPath = vbNullString
    End If
End Function

''' <summary>Get Query String</summary>
Public Function VtHttpaRequestGetQuery(ByVal reqHandle As LongPtr) As String
    If reqHandle <> 0 Then
       VtHttpaRequestGetQuery = VtCharUtils.VtUtf8PtrToString(vtnode_httpa_request_get_query(reqHandle))
    Else
     VtHttpaRequestGetQuery = vbNullString
    End If
End Function




' ==============================================================================
' VBA High-Level Wrapper Functions
' ==============================================================================

''' <summary>
''' HTTP 요청에 헤더를 추가합니다.
''' </summary>
Public Function VtHttpRequestAddHeader( _
    ByVal reqHandle As LongPtr, _
    ByVal HeaderName As String, _
    ByVal HeaderValue As String _
) As Boolean
    On Error GoTo EH
    If reqHandle = 0 Or Len(HeaderName) = 0 Then Exit Function
    
    Dim u8Name() As Byte, u8Value() As Byte
    u8Name = VtCharUtils.VtStringToUtf8Bytes(HeaderName)
    u8Value = VtCharUtils.VtStringToUtf8Bytes(HeaderValue)
    
    Dim Res As Long
    Res = vtnode_httpa_request_add_header(reqHandle, VarPtr(u8Name(0)), VarPtr(u8Value(0)))
    VtHttpRequestAddHeader = (Res <> 0)
    Exit Function
EH:
    VtHttpRequestAddHeader = False
End Function

''' <summary>
''' HTTP 요청의 헤더 값을 설정합니다. (기존 헤더가 존재하면 덮어씀)
''' </summary>
Public Function VtHttpRequestSetHeader( _
    ByVal reqHandle As LongPtr, _
    ByVal HeaderName As String, _
    ByVal HeaderValue As String _
) As Boolean
    On Error GoTo EH
    If reqHandle = 0 Or Len(HeaderName) = 0 Then Exit Function
    
    Dim u8Name() As Byte, u8Value() As Byte
    u8Name = VtCharUtils.VtStringToUtf8Bytes(HeaderName)
    u8Value = VtCharUtils.VtStringToUtf8Bytes(HeaderValue)
    
    Dim Res As Long
    Res = vtnode_httpa_request_set_header(reqHandle, VarPtr(u8Name(0)), VarPtr(u8Value(0)))
    VtHttpRequestSetHeader = (Res <> 0)
    Exit Function
EH:
    VtHttpRequestSetHeader = False
End Function

''' <summary>
''' 헤더 이름으로 지정된 헤더를 제거합니다.
''' </summary>
Public Sub VtHttpRequestRemoveHeader( _
    ByVal reqHandle As LongPtr, _
    ByVal HeaderName As String _
)
    On Error GoTo EH
    If reqHandle = 0 Or Len(HeaderName) = 0 Then Exit Sub
    
    Dim u8Name() As Byte
    u8Name = VtCharUtils.VtStringToUtf8Bytes(HeaderName)
    
    Call vtnode_httpa_request_remove_header(reqHandle, VarPtr(u8Name(0)))
EH:
End Sub

''' <summary>
''' 헤더 이름으로 헤더 값을 가져옵니다. (없을 경우 vbNullString 반환)
''' </summary>
Public Function VtHttpRequestGetHeader( _
    ByVal reqHandle As LongPtr, _
    ByVal HeaderName As String _
) As String
    On Error GoTo EH
    If reqHandle = 0 Or Len(HeaderName) = 0 Then
        VtHttpRequestGetHeader = vbNullString
        Exit Function
    End If
    
    Dim u8Name() As Byte
    u8Name = VtCharUtils.VtStringToUtf8Bytes(HeaderName)
    
    Dim valPtr As LongPtr
    valPtr = vtnode_httpa_request_get_header(reqHandle, VarPtr(u8Name(0)))
    
    ' DLL이 소유한 Read-only UTF-8 포인터를 VBA String으로 복사 변환
    VtHttpRequestGetHeader = VtUtf8PtrToString(valPtr)
    Exit Function
EH:
    VtHttpRequestGetHeader = vbNullString
End Function

''' <summary>
''' 1-based 인덱스로 헤더의 이름과 값을 함께 조회합니다.
''' 성공 시 True 반환 및 outHeaderName, outHeaderValue에 값 할당.
''' </summary>
Public Function VtHttpRequestGetHeaderAt( _
    ByVal reqHandle As LongPtr, _
    ByVal index As Long, _
    ByRef outHeaderName As String, _
    ByRef outHeaderValue As String _
) As Boolean
    On Error GoTo EH
    outHeaderName = vbNullString
    outHeaderValue = vbNullString
    
    If reqHandle = 0 Or index < 1 Then Exit Function
    
    Dim ptrName As LongPtr
    Dim ptrVal As LongPtr
    Dim Res As Long
    
    ' 1-base index를 DLL로 그대로 전달
    Res = vtnode_httpa_request_get_header_at(reqHandle, index, ptrName, ptrVal)
    
    If Res <> 0 Then
        outHeaderName = VtCharUtils.VtUtf8PtrToString(ptrName)
        outHeaderValue = VtCharUtils.VtUtf8PtrToString(ptrVal)
        VtHttpRequestGetHeaderAt = True
    Else
        VtHttpRequestGetHeaderAt = False
    End If
    Exit Function
EH:
    VtHttpRequestGetHeaderAt = False
End Function



''' <summary>
''' 멀티파트 요청 내에 포함된 Part(콘텐츠) 개수를 반환합니다.
''' </summary>
Public Function VtHttpaRequestGetContentCount(ByVal reqHandle As LongPtr) As Long
    If reqHandle <> 0 Then
    VtHttpaRequestGetContentCount = vtnode_httpa_request_content_get_count(reqHandle)
    Else
    VtHttpaRequestGetContentCount = 0
    End If
    
End Function


''' <summary>
''' Returns the content handle at the specified 1-based index.
''' Return 0 if the index is out of range or the request is invalid.
''' </summary>
Public Function VtHttpaRequestGetContentAt( _
    ByVal Request As LongPtr, _
    ByVal index As Long _
) As LongPtr

    Dim ContentHandle As LongPtr

    On Error GoTo EH

    VtHttpaRequestGetContentAt = 0
    
    

    If Request = 0 Then
        Exit Function
    End If

    If index < 1 Then
        Exit Function
    End If

   
    
 
    ContentHandle = vtnode_httpa_request_content_get_at(Request, index)
        


    VtHttpaRequestGetContentAt = ContentHandle
    Exit Function

EH:

    Debug.Print _
        "[VtHttpaRequestGetContentAt] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaRequestGetContentAt = 0

End Function


''' <summary>
''' Multipart Request에 파일 Content를 추가하고 Content Handle을 반환합니다.
''' 실패하면 0을 반환합니다.
''' </summary>
Public Function VtHttpaRequestAddContentFile( _
    ByVal Request As LongPtr, _
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

    VtHttpaRequestAddContentFile = 0

    ' Validate Request handle.
    If Request = 0 Then
        Debug.Print _
            "[VtHttpaRequestAddContentFile] Invalid Request handle."
        Exit Function
    End If

    ' Validate file path.
    If Len(Trim$(FilePath)) = 0 Then
        Debug.Print _
            "[VtHttpaRequestAddContentFile] FilePath is empty."
        Exit Function
    End If

    ' Verify that the file exists.
    If Len(dir$( _
        FilePath, _
        vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then

        Debug.Print _
            "[VtHttpaRequestAddContentFile] File not found: " & _
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

    ' Add file content and receive its Content Handle.
    ContentHandle = _
        vtnode_httpa_request_content_add_file( _
            Request, _
            pFilePath, _
            pContentType, _
            pDispositionName, _
            pDispositionFilename)

    If ContentHandle = 0 Then
        Debug.Print _
            "[VtHttpaRequestAddContentFile] " & _
            "Failed to add file content: " & _
            FilePath

        Exit Function
    End If

    VtHttpaRequestAddContentFile = ContentHandle
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaRequestAddContentFile] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaRequestAddContentFile = 0

End Function




''' <summary>
''' Multipart Request에 Text Content를 추가하고 Content Handle을 반환합니다.
''' 실패하면 0을 반환합니다.
''' </summary>
Public Function VtHttpaRequestAddContentText( _
    ByVal Request As LongPtr, _
    ByVal text As String, _
    Optional ByVal ContentType As String = "text/plain; charset=utf-8", _
    Optional ByVal ContentDispositionName As String = vbNullString _
) As LongPtr

    Dim utf8Text() As Byte
    Dim utf8ContentType() As Byte
    Dim utf8DispositionName() As Byte

    Dim pText As LongPtr
    Dim pContentType As LongPtr
    Dim pDispositionName As LongPtr

    Dim ContentHandle As LongPtr

    On Error GoTo EH

    VtHttpaRequestAddContentText = 0

    ' Validate Request handle.
    If Request = 0 Then
        Debug.Print _
            "[VtHttpaRequestAddContentText] Invalid Request handle."
        Exit Function
    End If

    ' Apply default Content-Type.
    If Len(Trim$(ContentType)) = 0 Then
        ContentType = "text/plain; charset=utf-8"
    End If

    ' Convert text to a null-terminated UTF-8 byte array.
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

    ' Add text content and receive its Content Handle.
    ContentHandle = _
        vtnode_httpa_request_content_add_text( _
            Request, _
            pText, _
            pContentType, _
            pDispositionName)

    If ContentHandle = 0 Then
        Debug.Print _
            "[VtHttpaRequestAddContentText] " & _
            "Failed to add text content."

        Exit Function
    End If

    VtHttpaRequestAddContentText = ContentHandle
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaRequestAddContentText] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaRequestAddContentText = 0

End Function


