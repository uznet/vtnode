Attribute VB_Name = "VtHttpaRuntimeTypes"
Option Explicit

Public Enum VtHttpaRuntimeStatus
    VT_RUNTIME_STOPPED = 0
    VT_RUNTIME_STARTING = 10
    VT_RUNTIME_RUNNING = 20
    VT_RUNTIME_STOPPING = 30
    VT_RUNTIME_FAULTED = 90
End Enum

'===============================================================================
' VtHttpaRuntimeTypes.bas
'===============================================================================


Public Enum VtResponseActivityType
    VT_RESPONSE_STATUS = 0
    VT_RESPONSE_TEXT = 10
    VT_RESPONSE_FILE = 20
    VT_RESPONSE_MULTIPART = 30
End Enum

Public Type VtHttpaRuntimeAppInfo
    ApplicationName As String
    WorkbookPath As String
End Type

Public Type VtHttpaRuntimeNetworkInfo
    LocalIp As String
    ExternalIp As String
End Type

Public Type VtHttpaRuntimeHttpInfo
    Enabled As Boolean
    Port As Long
    HttpUrl As String
    WebSocketUrl As String
End Type

Public Type VtHttpaRuntimeHttpsInfo
    Enabled As Boolean
    Port As Long
    HttpsUrl As String
    SecureWebSocketUrl As String
    CertificateFile As String
    PrivateKeyFile As String
End Type

Public Type VtHttpaRuntimeTimingInfo
    StartedAt As Date
    UptimeSeconds As Long
End Type

Public Type VtHttpaRuntimeStatsInfo
    TotalRequests As Long
    WebSocketClients As Long
End Type

Public Type VtNodeRuntimeState
    Status As VtHttpaRuntimeStatus
    App As VtHttpaRuntimeAppInfo
    Network As VtHttpaRuntimeNetworkInfo
    Http As VtHttpaRuntimeHttpInfo
    Https As VtHttpaRuntimeHttpsInfo
    Timing As VtHttpaRuntimeTimingInfo
    Stats As VtHttpaRuntimeStatsInfo
    LatestActivity As String
    LastError As String
End Type

Public Function VtHttpaRuntimeStatusToString( _
    ByVal value As VtHttpaRuntimeStatus) As String

    Select Case value
        Case VT_RUNTIME_STOPPED
            VtHttpaRuntimeStatusToString = "STOPPED"
        Case VT_RUNTIME_STARTING
            VtHttpaRuntimeStatusToString = "STARTING"
        Case VT_RUNTIME_RUNNING
            VtHttpaRuntimeStatusToString = "RUNNING"
        Case VT_RUNTIME_STOPPING
            VtHttpaRuntimeStatusToString = "STOPPING"
        Case VT_RUNTIME_FAULTED
            VtHttpaRuntimeStatusToString = "FAULTED"
        Case Else
            VtHttpaRuntimeStatusToString = "UNKNOWN"
    End Select
End Function
