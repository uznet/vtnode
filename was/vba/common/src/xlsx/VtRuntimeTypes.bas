Attribute VB_Name = "VtRuntimeTypes"
Option Explicit

Public Enum VtRuntimeStatus
    VT_RUNTIME_STOPPED = 0
    VT_RUNTIME_STARTING = 10
    VT_RUNTIME_RUNNING = 20
    VT_RUNTIME_STOPPING = 30
    VT_RUNTIME_FAULTED = 90
End Enum

'===============================================================================
' VtRuntimeTypes.bas
'===============================================================================


Public Enum VtResponseActivityType
    VT_RESPONSE_STATUS = 0
    VT_RESPONSE_TEXT = 10
    VT_RESPONSE_FILE = 20
    VT_RESPONSE_MULTIPART = 30
End Enum

Public Type VtRuntimeAppInfo
    ApplicationName As String
    WorkbookPath As String
End Type

Public Type VtRuntimeNetworkInfo
    LocalIp As String
    ExternalIp As String
End Type

Public Type VtRuntimeHttpInfo
    Enabled As Boolean
    Port As Long
    HttpUrl As String
    WebSocketUrl As String
End Type

Public Type VtRuntimeHttpsInfo
    Enabled As Boolean
    Port As Long
    HttpsUrl As String
    SecureWebSocketUrl As String
    CertificateFile As String
    PrivateKeyFile As String
End Type

Public Type VtRuntimeTimingInfo
    StartedAt As Date
    UptimeSeconds As Long
End Type

Public Type VtRuntimeStatsInfo
    TotalRequests As Long
    WebSocketClients As Long
End Type

Public Type VtNodeRuntimeState
    Status As VtRuntimeStatus
    App As VtRuntimeAppInfo
    Network As VtRuntimeNetworkInfo
    Http As VtRuntimeHttpInfo
    Https As VtRuntimeHttpsInfo
    Timing As VtRuntimeTimingInfo
    Stats As VtRuntimeStatsInfo
    LatestActivity As String
    LastError As String
End Type

Public Function VtRuntimeStatusToString( _
    ByVal value As VtRuntimeStatus) As String

    Select Case value
        Case VT_RUNTIME_STOPPED
            VtRuntimeStatusToString = "STOPPED"
        Case VT_RUNTIME_STARTING
            VtRuntimeStatusToString = "STARTING"
        Case VT_RUNTIME_RUNNING
            VtRuntimeStatusToString = "RUNNING"
        Case VT_RUNTIME_STOPPING
            VtRuntimeStatusToString = "STOPPING"
        Case VT_RUNTIME_FAULTED
            VtRuntimeStatusToString = "FAULTED"
        Case Else
            VtRuntimeStatusToString = "UNKNOWN"
    End Select
End Function
