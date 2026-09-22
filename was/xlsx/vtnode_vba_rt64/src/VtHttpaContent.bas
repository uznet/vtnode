Attribute VB_Name = "VtHttpaContent"
Option Explicit

'============================================================
' HTTP CONTENT ACCESS
'============================================================

#If VBA7 And Win64 Then


#If VTNODE_DEBUG = 1 Then  ' Debug Mode


' Get content data.
' OutData points to either a file path string or a memory buffer.
' OutDataLen receives the data length in bytes.
' OutIsFile receives TRUE when the content is file-backed.
Private Declare PtrSafe Function vtnode_httpa_content_get_data Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByRef OutData As LongPtr, _
    ByRef OutDataLen As Long, _
    ByRef OutIsFile As Long _
) As Long

' Get Content-Disposition information.
' Output buffers receive null-terminated UTF-8 strings.
Private Declare PtrSafe Function vtnode_httpa_content_get_disposition Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByVal OutType As LongPtr, _
    ByVal OutTypeSize As Long, _
    ByVal OutName As LongPtr, _
    ByVal OutNameSize As Long, _
    ByVal OutFilename As LongPtr, _
    ByVal OutFilenameSize As Long _
) As Long

' Add a content header value.
Private Declare PtrSafe Function vtnode_httpa_content_add_header Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Replace all existing values of the specified content header.
Private Declare PtrSafe Function vtnode_httpa_content_set_header Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Remove all values of the specified content header.
Private Declare PtrSafe Function vtnode_httpa_content_remove_header Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr _
) As Long

' Remove a content header value by zero-based value index.
Private Declare PtrSafe Function vtnode_httpa_content_remove_header_at Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal index As Long _
) As Long

' Return the number of registered content header names.
Private Declare PtrSafe Function vtnode_httpa_content_get_header_count Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr _
) As Long

' Return the number of values for the specified content header.
Private Declare PtrSafe Function vtnode_httpa_content_get_header_value_count Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr _
) As Long

' Return the first value of the specified content header.
' Returned string is read-only UTF-8 and owned by content.
Private Declare PtrSafe Function vtnode_httpa_content_get_header Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr _
) As LongPtr

' Return a content header name/value by 1-based header index.
' Return NULL if Index is out of range.
' HeaderValueOut receives a read-only UTF-8 string pointer owned by content.
Private Declare PtrSafe Function vtnode_httpa_content_get_header_at Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByVal index As Long, _
    ByRef HeaderValueOut As LongPtr _
) As LongPtr

' Return the content header name at the specified 1-based position.
' Returned string is read-only UTF-8 and owned by content.
Private Declare PtrSafe Function vtnode_httpa_content_get_header_name_at Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByVal index As Long _
) As LongPtr

' Return a content header value at the specified index.
' HeaderName is a null-terminated UTF-8 string.
' Returned string is read-only UTF-8 and owned by content.
Private Declare PtrSafe Function vtnode_httpa_content_get_header_value_at Lib "vtnode64d.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal index As Long _
) As LongPtr



Private Declare PtrSafe Function vtnode_httpa_header_get_merge_policy Lib "vtnode64d.dll" (ByVal HeaderName As LongPtr) As Long

#Else


Private Declare PtrSafe Function vtnode_httpa_content_get_data Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByRef OutData As LongPtr, _
    ByRef OutDataLen As Long, _
    ByRef OutIsFile As Long _
) As Long

' Get Content-Disposition information.
' Output buffers receive null-terminated UTF-8 strings.
Private Declare PtrSafe Function vtnode_httpa_content_get_disposition Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByVal OutType As LongPtr, _
    ByVal OutTypeSize As Long, _
    ByVal OutName As LongPtr, _
    ByVal OutNameSize As Long, _
    ByVal OutFilename As LongPtr, _
    ByVal OutFilenameSize As Long _
) As Long

' Add a content header value.
Private Declare PtrSafe Function vtnode_httpa_content_add_header Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Replace all existing values of the specified content header.
Private Declare PtrSafe Function vtnode_httpa_content_set_header Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal HeaderValue As LongPtr _
) As Long

' Remove all values of the specified content header.
Private Declare PtrSafe Function vtnode_httpa_content_remove_header Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr _
) As Long

' Remove a content header value by zero-based value index.
Private Declare PtrSafe Function vtnode_httpa_content_remove_header_at Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal index As Long _
) As Long

' Return the number of registered content header names.
Private Declare PtrSafe Function vtnode_httpa_content_get_header_count Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr _
) As Long

' Return the number of values for the specified content header.
Private Declare PtrSafe Function vtnode_httpa_content_get_header_value_count Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr _
) As Long

' Return the first value of the specified content header.
' Returned string is read-only UTF-8 and owned by content.
Private Declare PtrSafe Function vtnode_httpa_content_get_header Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr _
) As LongPtr

' Return a content header name/value by 1-based header index.
' Return NULL if Index is out of range.
' HeaderValueOut receives a read-only UTF-8 string pointer owned by content.
Private Declare PtrSafe Function vtnode_httpa_content_get_header_at Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByVal index As Long, _
    ByRef HeaderValueOut As LongPtr _
) As LongPtr

' Return the content header name at the specified 1-based position.
' Returned string is read-only UTF-8 and owned by content.
Private Declare PtrSafe Function vtnode_httpa_content_get_header_name_at Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByVal index As Long _
) As LongPtr

' Return a content header value at the specified index.
' HeaderName is a null-terminated UTF-8 string.
' Returned string is read-only UTF-8 and owned by content.
Private Declare PtrSafe Function vtnode_httpa_content_get_header_value_at Lib "vtnode64.dll" ( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As LongPtr, _
    ByVal index As Long _
) As LongPtr



Private Declare PtrSafe Function vtnode_httpa_header_get_merge_policy Lib "vtnode64.dll" (ByVal HeaderName As LongPtr) As Long


#End If
#End If



''' <summary>
''' Content Handle의 데이터를 VTHttpaContentData 구조체로 반환합니다.
'''
''' File Content:
'''     IsFile = True
'''     FilePath 사용
'''
''' Memory Content:
'''     IsFile = False
'''     bytes() 사용
''' </summary>
Public Function VtHttpaContentGetData( _
    ByVal Content As LongPtr _
) As VTHttpaContentData

    Dim result As VTHttpaContentData

    Dim pData As LongPtr
    Dim DataLen As Long
    Dim isFileFlag As Long

    Dim nativeResult As Long

    On Error GoTo EH

    ' --------------------------------------------------------
    ' Initialize Result
    ' --------------------------------------------------------
    result.IsValid = False
    result.DataType = VTHTTPA_CONTENT_DATATYPE_NONE
    result.ByteLen = 0

    result.text = vbNullString
    result.FilePath = vbNullString

    Erase result.Bytes

    ' --------------------------------------------------------
    ' Validate Handle
    ' --------------------------------------------------------
    If Content = 0 Then
        VtHttpaContentGetData = result
        Exit Function
    End If

    ' --------------------------------------------------------
    ' Get Native Content Data
    ' --------------------------------------------------------
    nativeResult = _
        vtnode_httpa_content_get_data( _
            Content, _
            pData, _
            DataLen, _
            isFileFlag)
            
   ' Debug.Print "VtHttpaContentGetData:DataLen=" & CStr(DataLen)
    

    If nativeResult = 0 Then
        VtHttpaContentGetData = result
        Exit Function
    End If

    ' --------------------------------------------------------
    ' File-backed Content
    ' --------------------------------------------------------
    If isFileFlag <> 0 Then

        result.DataType = VTHTTPA_CONTENT_DATATYPE_FILE

        If pData <> 0 Then

            result.FilePath = _
                VtCharUtils.VtUtf8PtrToString(pData)

        End If

    ' --------------------------------------------------------
    ' Memory-backed Content
    ' --------------------------------------------------------
    Else

        result.DataType = VTHTTPA_CONTENT_DATATYPE_BYTES

        If pData <> 0 And DataLen > 0 Then

            result.Bytes = VtCharUtils.VtPtrToBytes(pData, DataLen)
            
        
'            Dim ByteLen As Long
'            ByteLen = VtCharUtils.VtByteArrayLength(result.Bytes)
'
'            Debug.Print "VtHttpaContentGetData:ByteLen=" & CStr(ByteLen)


        End If

    End If

    ' --------------------------------------------------------
    ' Content-Disposition
    ' --------------------------------------------------------
    result.ContentDisposition = VtHttpaContentGetDisposition(Content)

    result.IsValid = True

    VtHttpaContentGetData = result
    Exit Function

EH:

    Debug.Print _
        "[VtHttpaContentGetData] Error " & _
        Err.Number & ": " & Err.Description

    result.IsValid = False
    result.DataType = VTHTTPA_CONTENT_DATATYPE_NONE

    VtHttpaContentGetData = result

End Function

''' <summary>
''' Returns the Content-Disposition information of a content handle.
''' </summary>
Public Function VtHttpaContentGetDisposition( _
    ByVal Content As LongPtr _
) As VTHttpaContentDisposition

    Const BUFFER_SIZE As Long = 512

    Dim result As VTHttpaContentDisposition

    Dim typeBuffer(0 To BUFFER_SIZE - 1) As Byte
    Dim nameBuffer(0 To BUFFER_SIZE - 1) As Byte
    Dim filenameBuffer(0 To BUFFER_SIZE - 1) As Byte

    Dim nativeResult As Long

    On Error GoTo EH

    If Content = 0 Then
        VtHttpaContentGetDisposition = result
        Exit Function
    End If

    nativeResult = _
        vtnode_httpa_content_get_disposition( _
            Content, _
            VarPtr(typeBuffer(0)), _
            BUFFER_SIZE, _
            VarPtr(nameBuffer(0)), _
            BUFFER_SIZE, _
            VarPtr(filenameBuffer(0)), _
            BUFFER_SIZE)

    If nativeResult = 0 Then
        VtHttpaContentGetDisposition = result
        Exit Function
    End If

    result.DispType = _
        VtCharUtils.VtUtf8BytesToString(typeBuffer)

    result.DispName = _
        VtCharUtils.VtUtf8BytesToString(nameBuffer)

    result.DispFilename = _
        VtCharUtils.VtUtf8BytesToString(filenameBuffer)

    VtHttpaContentGetDisposition = result
    Exit Function

EH:

    Debug.Print _
        "[VtHttpaContentGetDisposition] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaContentGetDisposition = result

End Function



''' <summary>
''' Content Header 값을 추가합니다.
''' 같은 이름의 Header가 존재하면 값을 추가합니다.
''' </summary>
Public Function VtHttpaContentAddHeader( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As String, _
    ByVal HeaderValue As String _
) As Boolean

    Dim utf8HeaderName() As Byte
    Dim utf8HeaderValue() As Byte
    Dim nativeResult As Long

    On Error GoTo EH

    VtHttpaContentAddHeader = False

    If Content = 0 Then
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
        vtnode_httpa_content_add_header( _
            Content, _
            VarPtr(utf8HeaderName(0)), _
            VarPtr(utf8HeaderValue(0)))

    VtHttpaContentAddHeader = (nativeResult <> 0)
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaContentAddHeader] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaContentAddHeader = False

End Function


''' <summary>
''' Content Header 값을 설정합니다.
''' 같은 이름의 기존 Header 값은 모두 교체됩니다.
''' </summary>
Public Function VtHttpaContentSetHeader( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As String, _
    ByVal HeaderValue As String _
) As Boolean

    Dim utf8HeaderName() As Byte
    Dim utf8HeaderValue() As Byte
    Dim nativeResult As Long

    On Error GoTo EH

    VtHttpaContentSetHeader = False

    If Content = 0 Then
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
        vtnode_httpa_content_set_header( _
            Content, _
            VarPtr(utf8HeaderName(0)), _
            VarPtr(utf8HeaderValue(0)))

    VtHttpaContentSetHeader = (nativeResult <> 0)
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaContentSetHeader] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaContentSetHeader = False

End Function


''' <summary>
''' 지정한 이름의 Content Header 값을 모두 제거합니다.
''' </summary>
Public Function VtHttpaContentRemoveHeader( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As String _
) As Boolean

    Dim utf8HeaderName() As Byte
    Dim nativeResult As Long

    On Error GoTo EH

    VtHttpaContentRemoveHeader = False

    If Content = 0 Then
        Exit Function
    End If

    If Len(Trim$(HeaderName)) = 0 Then
        Exit Function
    End If

    utf8HeaderName = _
        VtCharUtils.VtStringToUtf8Bytes(HeaderName)

    nativeResult = _
        vtnode_httpa_content_remove_header( _
            Content, _
            VarPtr(utf8HeaderName(0)))

    VtHttpaContentRemoveHeader = (nativeResult <> 0)
    Exit Function

EH:
    Debug.Print _
        "[VtHttpaContentRemoveHeader] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaContentRemoveHeader = False

End Function



''' <summary>
''' 등록된 Content Header 이름의 개수를 반환합니다.
''' Content Handle이 유효하지 않거나 오류가 발생하면 0을 반환합니다.
''' </summary>
Public Function VtHttpaContentGetHeaderCount( _
    ByVal Content As LongPtr _
) As Long

    On Error GoTo EH

    VtHttpaContentGetHeaderCount = 0

    If Content = 0 Then
        Exit Function
    End If

    VtHttpaContentGetHeaderCount = _
        vtnode_httpa_content_get_header_count(Content)

    Exit Function

EH:
    Debug.Print _
        "[VtHttpaContentGetHeaderCount] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaContentGetHeaderCount = 0

End Function


''' <summary>
''' 지정한 이름의 Content Header에서 첫 번째 값을 반환합니다.
''' Header가 없거나 Content Handle이 유효하지 않으면 vbNullString을 반환합니다.
''' </summary>
Public Function VtHttpaContentGetHeader( _
    ByVal Content As LongPtr, _
    ByVal HeaderName As String _
) As String

    Dim utf8HeaderName() As Byte
    Dim pHeaderValue As LongPtr

    On Error GoTo EH

    VtHttpaContentGetHeader = vbNullString

    If Content = 0 Then
        Exit Function
    End If

    If Len(Trim$(HeaderName)) = 0 Then
        Exit Function
    End If

    utf8HeaderName = _
        VtCharUtils.VtStringToUtf8Bytes(HeaderName)

    pHeaderValue = _
        vtnode_httpa_content_get_header( _
            Content, _
            VarPtr(utf8HeaderName(0)))

    If pHeaderValue <> 0 Then
        VtHttpaContentGetHeader = _
            VtCharUtils.VtUtf8PtrToString(pHeaderValue)
    End If

    Exit Function

EH:
    Debug.Print _
        "[VtHttpaContentGetHeader] Error " & _
        Err.Number & ": " & Err.Description

    VtHttpaContentGetHeader = vbNullString

End Function



''' <summary>
''' 1-based 인덱스로 Content Header 이름과 값을 가져옵니다.
''' 성공하면 True를 반환하고 HeaderNameOut, HeaderValueOut에 값을 저장합니다.
''' </summary>
Public Function VtHttpaContentGetHeaderAt( _
    ByVal Content As LongPtr, _
    ByVal index As Long, _
    ByRef HeaderNameOut As String, _
    ByRef HeaderValueOut As String _
) As Boolean

    Dim pHeaderName As LongPtr
    Dim pHeaderValue As LongPtr

    On Error GoTo EH

    HeaderNameOut = vbNullString
    HeaderValueOut = vbNullString
    VtHttpaContentGetHeaderAt = False

    If Content = 0 Then
        Exit Function
    End If

    If index < 1 Then
        Exit Function
    End If

    pHeaderName = _
        vtnode_httpa_content_get_header_at( _
            Content, _
            index, _
            pHeaderValue)

    If pHeaderName = 0 Then
        Exit Function
    End If

    HeaderNameOut = _
        VtCharUtils.VtUtf8PtrToString(pHeaderName)

    If pHeaderValue <> 0 Then
        HeaderValueOut = _
            VtCharUtils.VtUtf8PtrToString(pHeaderValue)
    End If

    VtHttpaContentGetHeaderAt = True
    Exit Function

EH:

    Debug.Print _
        "[VtHttpaContentGetHeaderAt] Error " & _
        Err.Number & ": " & Err.Description

    HeaderNameOut = vbNullString
    HeaderValueOut = vbNullString
    VtHttpaContentGetHeaderAt = False

End Function





Public Function VtHttpaHeaderGetMergePolicy( _
    ByVal HeaderName As String _
) As VTHttpaHeaderMergePolicy

    Dim NameBytes() As Byte

    NameBytes = VtCharUtils.VtStringToUtf8Bytes(HeaderName)

    VtHttpaHeaderGetMergePolicy = vtnode_httpa_header_get_merge_policy(VarPtr(NameBytes(0)))

End Function

