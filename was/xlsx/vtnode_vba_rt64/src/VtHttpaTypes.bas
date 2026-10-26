Attribute VB_Name = "VtHttpaTypes"

Option Explicit



' ============================================================================
' VTNode HTTPA VBA Types
' ============================================================================
'
' This module defines application-level Enum values and types.
'
' Native Enum values are defined in VtHttpaNativeTypes.bas.
' Do not modify native Enum definitions.
'
' ============================================================================

Public Enum VTHttpaSessionRole

    VTHTTPA_SESSION_ROLE_NONE = 0
    VTHTTPA_SESSION_ROLE_CLIENT
    VTHTTPA_SESSION_ROLE_SERVER
    
End Enum
' Don't Modify.
' it must be equal C Enum Value
Public Enum VTHttpaConnectionMode
    VTHTTPA_CONNECTION_MODE_UNSET = 0
    VTHTTPA_CONNECTION_MODE_CLOSE       ' default connection mode. connection will be closed after response is sent.
    VTHTTPA_CONNECTION_MODE_KEEP_ALIVE  ' connection will be kept alive after response is sent, and can be reused for next request.
End Enum



Public Enum VTHttpaMethod
    VTHTTPA_METHOD_UNKNOWN = 0
    VTHTTPA_METHOD_GET
    VTHTTPA_METHOD_POST
    VTHTTPA_METHOD_PUT
    VTHTTPA_METHOD_DELETE
    VTHTTPA_METHOD_PATCH
    VTHTTPA_METHOD_HEAD
    VTHTTPA_METHOD_OPTIONS
    ' 필요에 따라 다른 HTTP 메서드 추가 가능
End Enum

Public Enum VTHttpaWebsocketState

    VTHTTPA_WEBSOCKET_STATE_CLOSED = 0    ' closed
    VTHTTPA_WEBSOCKET_STATE_CONNECTING  ' connecting
    VTHTTPA_WEBSOCKET_STATE_CONNECTED  ' connected
    VTHTTPA_WEBSOCKET_STATE_CLOSING     ' closing
    
End Enum

Public Enum VTHttpaContentDataType
    VTHTTPA_CONTENT_DATATYPE_NONE = 0
    VTHTTPA_CONTENT_DATATYPE_TEXT = 1
    VTHTTPA_CONTENT_DATATYPE_BYTES = 2
    VTHTTPA_CONTENT_DATATYPE_FILE = 4
    
End Enum

' 헤더 정보를 담는 타입
Public Type VTHttpaHeaderInfo
    Name As String
    value As String
    IsValid As Boolean  ' 헤더 존재 여부
End Type

Public Type VTHttpaQueryParam
    key As String
    value As String
    IsValid As Boolean  ' 파라미터 존재 여부
End Type


Public Type VTHttpaContentDisposition
    DispType As String     ' e.g. "form-data", "attachment"
    DispName As String     ' e.g. "upload_file", "username"
    DispFilename As String ' e.g. "report.pdf"
End Type



Public Type VTHttpaContentData
    IsValid As Boolean
    DataType As VTHttpaContentDataType
    
    
    Bytes() As Byte
    FilePath As String
    text As String
    ByteLen As Long
    
        
    ContentType As String
    ContentDisposition As VTHttpaContentDisposition
    
End Type



Public Type VTHttpaEngineStartResult

    success As Boolean
    emsg    As String
        
End Type

Public Function VTHttpaEventTypeToString(EventType As VTHttpaEventType) As String
    Select Case EventType
        Case VTHTTPA_EVENT_ERROR
            VTHttpaEventTypeToString = "ERROR"
        Case VTHTTPA_EVENT_NET_ACCEPTED
            VTHttpaEventTypeToString = "NET_ACCEPTED"
        Case VTHTTPA_EVENT_NET_CONNECTING
            VTHttpaEventTypeToString = "NET_CONNECTING"
        Case VTHTTPA_EVENT_NET_CONNECT_FAILED
            VTHttpaEventTypeToString = "NET_CONNECT_FAILED"
        Case VTHTTPA_EVENT_NET_CONNECTED
            VTHttpaEventTypeToString = "NET_CONNECTED"
        Case VTHTTPA_EVENT_NET_DISCONNECTED
            VTHttpaEventTypeToString = "NET_DISCONNECTED"
        Case VTHTTPA_EVENT_NET_ERROR
            VTHttpaEventTypeToString = "NET_ERROR"
        Case VTHTTPA_EVENT_HT_ERROR
            VTHttpaEventTypeToString = "HT_ERROR"
        Case VTHTTPA_EVENT_HT_TIMEOUT
            VTHttpaEventTypeToString = "HT_TIMEOUT"
        Case VTHTTPA_EVENT_HT_TRANSACTION_BEGIN
            VTHttpaEventTypeToString = "HT_TRANSACTION_BEGIN"
        Case VTHTTPA_EVENT_HT_TRANSACTION_END
            VTHttpaEventTypeToString = "HT_TRANSACTION_END"
        Case VTHTTPA_EVENT_HT_REQUEST_COMPLETED
            VTHttpaEventTypeToString = "HT_REQUEST_COMPLETED"
        Case VTHTTPA_EVENT_HT_RESPONSE_COMPLETED
            VTHttpaEventTypeToString = "HT_RESPONSE_COMPLETED"
        Case VTHTTPA_EVENT_WS_ESTABLISHED
            VTHttpaEventTypeToString = "WS_ESTABLISHED"
        Case VTHTTPA_EVENT_WS_TEXT_RECEIVED
            VTHttpaEventTypeToString = "WS_TEXT_RECEIVED"
        Case VTHTTPA_EVENT_WS_BINARY_RECEIVED
            VTHttpaEventTypeToString = "WS_BINARY_RECEIVED"
        Case VTHTTPA_EVENT_WS_CLOSED
            VTHttpaEventTypeToString = "WS_CLOSED"
        Case VTHTTPA_EVENT_WS_TIMEOUT
            VTHttpaEventTypeToString = "WS_TIMEOUT"
        Case VTHTTPA_EVENT_WS_ERROR
            VTHttpaEventTypeToString = "WS_ERROR"
        Case Else
            VTHttpaEventTypeToString = "UNKNOWN"
    End Select
End Function


Public Function VTHttpaMethodToString(Method As VTHttpaMethod) As String
    Select Case Method
        Case VTHTTPA_METHOD_GET
            VTHttpaMethodToString = "GET"
        Case VTHTTPA_METHOD_POST
            VTHttpaMethodToString = "POST"
        Case VTHTTPA_METHOD_PUT
            VTHttpaMethodToString = "PUT"
        Case VTHTTPA_METHOD_DELETE
            VTHttpaMethodToString = "DELETE"
        Case VTHTTPA_METHOD_PATCH
            VTHttpaMethodToString = "PATCH"
        Case VTHTTPA_METHOD_HEAD
            VTHttpaMethodToString = "HEAD"
        Case VTHTTPA_METHOD_OPTIONS
            VTHttpaMethodToString = "OPTIONS"
        Case Else
            VTHttpaMethodToString = "UNKNOWN"
    End Select
End Function

Public Function VTHttpaMethodFromString(MethodStr As String) As VTHttpaMethod
    Select Case UCase$(MethodStr)
        Case "GET"
            VTHttpaMethodFromString = VTHTTPA_METHOD_GET
        Case "POST"
            VTHttpaMethodFromString = VTHTTPA_METHOD_POST
        Case "PUT"
            VTHttpaMethodFromString = VTHTTPA_METHOD_PUT
        Case "DELETE"
            VTHttpaMethodFromString = VTHTTPA_METHOD_DELETE
        Case "PATCH"
            VTHttpaMethodFromString = VTHTTPA_METHOD_PATCH
        Case "HEAD"
            VTHttpaMethodFromString = VTHTTPA_METHOD_HEAD
        Case "OPTIONS"
            VTHttpaMethodFromString = VTHTTPA_METHOD_OPTIONS
        Case Else
            VTHttpaMethodFromString = VTHTTPA_METHOD_UNKNOWN
    End Select
End Function
