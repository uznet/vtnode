Attribute VB_Name = "VtHttpaClient"
Option Explicit

#If VBA7 And Win64 Then

    ' 64비트 환경 및 xhttpa64d.dll 라이브러리 선언
    ' String -> LongPtr 변경 (직접 주소 전달)
 #If VTNODE_DEBUG = 1 Then
    Declare PtrSafe Function vtnode_httpa_client_get Lib "vtnode64d.dll" ( _
        ByVal UziId As LongLong, _
        ByVal Uri As LongPtr, _
        ByVal optionsId As LongLong) As Boolean

    Declare PtrSafe Function vtnode_httpa_client_post_text Lib "vtnode64d.dll" ( _
        ByVal UziId As LongLong, _
        ByVal Uri As LongPtr, _
        ByVal text As LongPtr, _
        ByVal TextLen As Long, _
        ByVal ContentType As LongPtr, _
        ByVal optionsId As LongLong) As Boolean

    Declare PtrSafe Function vtnode_httpa_client_post_file Lib "vtnode64d.dll" ( _
        ByVal UziId As LongLong, _
        ByVal Uri As LongPtr, _
        ByVal FilePath As LongPtr, _
        ByVal ContentType As LongPtr, _
        ByVal optionsId As LongLong) As Boolean

    ' Start POST multipart/form-data request.
    ' uri: String의 메모리 주소 (LongPtr)
    ' contentsId: 콘텐츠 묶음의 핸들 (LongLong)
    ' optionsId: 요청 옵션의 핸들 (LongLong)
    
    Declare PtrSafe Function vtnode_httpa_client_post_contents Lib "vtnode64d.dll" ( _
        ByVal UziId As LongLong, _
        ByVal Uri As LongPtr, _
        ByVal contentsId As LongLong, _
        ByVal optionsId As LongLong) As Boolean

    ' --- HTTP Response 정보 추출 ---
    
    ' Response Status Code 가져오기 (예: 200, 404, 500)
    Private Declare PtrSafe Function vtnode_httpa_event_get_response_status Lib "vtnode64d.dll" ( _
        ByVal eid As LongLong) As Long
    
    ' Response Reason Phrase 가져오기 (예: "OK", "Not Found")
    Private Declare PtrSafe Function vtnode_httpa_event_get_response_reason Lib "vtnode64d.dll" ( _
        ByVal eid As LongLong) As LongPtr
    
    ' Response Header 추출 (이름 기준)
    Private Declare PtrSafe Function vtnode_httpa_event_get_response_header Lib "vtnode64d.dll" ( _
        ByVal eid As LongLong, _
        ByVal hname As LongPtr, _
        ByRef hvaluePtr As LongPtr) As Long
    
    ' Response Header 추출 (인덱스 기준, 1-based)
    Private Declare PtrSafe Function vtnode_httpa_event_get_response_header_by_index Lib "vtnode64d.dll" ( _
        ByVal eid As LongLong, _
        ByVal index As Long, _
        ByRef hnamePtr As LongPtr, _
        ByRef hvaluePtr As LongPtr) As Long
         ' 64비트 환경
    Private Declare PtrSafe Function vtnode_httpa_event_get_response_contents Lib "vtnode64d.dll" ( _
        ByVal eid As LongLong) As LongLong

  #Else  ' released mode
    
    Declare PtrSafe Function vtnode_httpa_client_get Lib "vtnode64.dll" ( _
        ByVal UziId As LongLong, _
        ByVal Uri As LongPtr, _
        ByVal optionsId As LongLong) As Boolean

    Declare PtrSafe Function vtnode_httpa_client_post_text Lib "vtnode64.dll" ( _
        ByVal UziId As LongLong, _
        ByVal Uri As LongPtr, _
        ByVal text As LongPtr, _
        ByVal TextLen As Long, _
        ByVal ContentType As LongPtr, _
        ByVal optionsId As LongLong) As Boolean

    Declare PtrSafe Function vtnode_httpa_client_post_file Lib "vtnode64.dll" ( _
        ByVal UziId As LongLong, _
        ByVal Uri As LongPtr, _
        ByVal FilePath As LongPtr, _
        ByVal ContentType As LongPtr, _
        ByVal optionsId As LongLong) As Boolean

    ' Start POST multipart/form-data request.
    ' uri: String의 메모리 주소 (LongPtr)
    ' contentsId: 콘텐츠 묶음의 핸들 (LongLong)
    ' optionsId: 요청 옵션의 핸들 (LongLong)
    
    Declare PtrSafe Function vtnode_httpa_client_post_contents Lib "vtnode64.dll" ( _
        ByVal UziId As LongLong, _
        ByVal Uri As LongPtr, _
        ByVal contentsId As LongLong, _
        ByVal optionsId As LongLong) As Boolean

    ' --- HTTP Response 정보 추출 ---
    
    ' Response Status Code 가져오기 (예: 200, 404, 500)
    Private Declare PtrSafe Function vtnode_httpa_event_get_response_status Lib "vtnode64.dll" ( _
        ByVal eid As LongLong) As Long
    
    ' Response Reason Phrase 가져오기 (예: "OK", "Not Found")
    Private Declare PtrSafe Function vtnode_httpa_event_get_response_reason Lib "vtnode64.dll" ( _
        ByVal eid As LongLong) As LongPtr
    
    ' Response Header 추출 (이름 기준)
    Private Declare PtrSafe Function vtnode_httpa_event_get_response_header Lib "vtnode64.dll" ( _
        ByVal eid As LongLong, _
        ByVal hname As LongPtr, _
        ByRef hvaluePtr As LongPtr) As Long
    
    ' Response Header 추출 (인덱스 기준, 1-based)
    Private Declare PtrSafe Function vtnode_httpa_event_get_response_header_by_index Lib "vtnode64.dll" ( _
        ByVal eid As LongLong, _
        ByVal index As Long, _
        ByRef hnamePtr As LongPtr, _
        ByRef hvaluePtr As LongPtr) As Long
         ' 64비트 환경
    Private Declare PtrSafe Function vtnode_httpa_event_get_response_contents Lib "vtnode64.dll" ( _
        ByVal eid As LongLong) As LongLong
  #End If
  
#End If

' ============================================================
' VBA 래퍼 함수들
' ============================================================
Public Function VtHttpaClientGet(ByVal UziId As LongLong, ByVal Uri As String, ByVal optionsId As LongLong) As Boolean
    Dim utf8Uri() As Byte
    utf8Uri = VtStringToUtf8Bytes(Uri)
    VtHttpaClientGet = vtnode_httpa_client_get(UziId, VarPtr(utf8Uri(0)), optionsId)
End Function

Public Function VtHttpaClientPostText(ByVal UziId As LongLong, ByVal Uri As String, ByVal text As String, ByVal ContentType As String, ByVal optionsId As LongLong) As Boolean
    Dim utf8Uri() As Byte
    Dim utf8Text() As Byte
    Dim utf8TextLen As Long
    Dim utf8ContentType() As Byte
    utf8Uri = VtStringToUtf8Bytes(Uri)
    utf8Text = VtStringToUtf8Bytes(text)
    utf8TextLen = (UBound(utf8Text) - LBound(utf8Text) + 1) - 1
    utf8ContentType = VtStringToUtf8Bytes(ContentType)
    VtHttpaClientPostText = vtnode_httpa_client_post_text(UziId, VarPtr(utf8Uri(0)), VarPtr(utf8Text(0)), utf8TextLen, VarPtr(utf8ContentType(0)), optionsId)
End Function

Public Function VtHttpaClientPostFile(ByVal UziId As LongLong, ByVal Uri As String, ByVal FilePath As String, ByVal ContentType As String, ByVal optionsId As LongLong) As Boolean
    Dim utf8Uri() As Byte
    Dim utf8FilePath() As Byte
    Dim utf8ContentType() As Byte
    
    utf8Uri = VtStringToUtf8Bytes(Uri)
    utf8FilePath = VtStringToUtf8Bytes(FilePath)
    utf8ContentType = VtStringToUtf8Bytes(ContentType)
    VtHttpaClientPostFile = vtnode_httpa_client_post_file(UziId, VarPtr(utf8Uri(0)), VarPtr(utf8FilePath(0)), VarPtr(utf8ContentType(0)), optionsId)
End Function

Public Function VtHttpaClientPostContents(ByVal UziId As LongLong, ByVal Uri As String, ByVal contentsId As LongLong, ByVal optionsId As LongLong) As Boolean
    Dim utf8Uri() As Byte
    utf8Uri = VtStringToUtf8Bytes(Uri)
    VtHttpaClientPostContents = vtnode_httpa_client_post_contents(UziId, VarPtr(utf8Uri(0)), contentsId, optionsId)
End Function


Public Function VtHttpaEventGetResponseStatus(ByVal EventId As LongLong) As Long
    VtHttpaEventGetResponseStatus = vtnode_httpa_event_get_response_status(EventId)
End Function

Public Function VtHttpaEventGetResponseReason(ByVal EventId As LongLong) As String
    Dim pReason As LongPtr
    pReason = vtnode_httpa_event_get_response_reason(EventId)
    VtHttpaEventGetResponseReason = VtUtf8PtrToString(pReason)
End Function

Public Function VtHttpaEventGetResponseHeader(ByVal EventId As LongLong, ByVal HeaderName As String) As String
    Dim utf8Header() As Byte
    Dim pValue As LongPtr
    Dim result As Long

    utf8Header = VtStringToUtf8Bytes(HeaderName)
    result = vtnode_httpa_event_get_response_header(EventId, VarPtr(utf8Header(0)), pValue)

    If result <> 0 Then
        VtHttpaEventGetResponseHeader = VtUtf8PtrToString(pValue)
    Else
        VtHttpaEventGetResponseHeader = ""
    End If
End Function

Public Function VtHttpaEventGetResponseHeaderInfoByIndex(ByVal EventId As LongLong, ByVal index As Long) As String
    Dim pName As LongPtr
    Dim pValue As LongPtr
    Dim result As VTHttpaHeaderInfo

    result = vtnode_httpa_event_get_response_header_by_index(EventId, index, pName, pValue)

    If result <> 0 Then
        VtHttpaEventGetResponseHeaderByIndex = VtUtf8PtrToString(pName) & ": " & VtUtf8PtrToString(pValue)
    Else
        VtHttpaEventGetResponseHeaderByIndex = ""
    End If
End Function

Public Function VtHttpaEventGetResponseContents(ByVal EventId As LongLong) As LongLong
    VtHttpaEventGetResponseContents = vtnode_httpa_event_get_response_contents(EventId)
End Function
