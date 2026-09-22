VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "CVtHttpaSession"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Option Explicit


Private mIid As LongLong
Private mKey As String
Private mRole   As VTHttpaSessionRole   ' VTHTTPA_SESSION_ROLE_CLIENT or VTHTTPA_SESSION_ROLE_SERVER

Private mHttpaService As CVtHttpaService

Private mEventId   As LongLong 'last event id
Private mEventType As VTHttpaEventType  ' last event type

Private mSocket As CVtHttpaSocket
Private mRequest As CVtHttpaRequest
Private mResponse As CVtHttpaResponse
Private mWebsocket As CVtHttpaWebsocket
Private mErrorMsg As String
Private mTimeoutReason As String


Private mRequestSent As Boolean
Private mResponseSent As Boolean
Private mIsWebsocket  As Boolean




Private mCustomData As Object


Private Sub Clear()
          
    mErrorMsg = ""
    mTimeoutReason = ""
    
    mRequestSent = False
    mResponseSent = False
    
    
    
    
End Sub

Private Sub Class_Initialize()

    
    mIid = 0
    mKey = ""
    mRole = VTHTTPA_SESSION_ROLE_NONE
    
    mEventId = 0
    mEventType = VTHTTPA_EVENT_NONE
    
    Set mCustomData = Nothing
    
    Call Clear
        
    
End Sub

Private Sub Class_Terminate()
    Call Clear
    
    Set mSocket = Nothing
    Set mRequest = Nothing
    Set mResponse = Nothing
    Set mWebsocket = Nothing

End Sub

Public Sub Init(IID As LongLong, ByVal HttpaService As CVtHttpaService)
     
     
    mIid = IID
    mKey = CStr(IID)
    
    Set mHttpaService = HttpaService
    
    
End Sub


Public Property Get HttpaService() As CVtHttpaService

   Set HttpaService = mHttpaService
   
End Property

Public Sub NewTransaction()


    Call Clear
    
        
'    Set mSocket = New CVtHttpaSocket
    Set mRequest = Nothing 'New CVtHttpaRequest
    Set mResponse = Nothing 'New CVtHttpaResponse
    Set mWebsocket = Nothing

   
End Sub

Public Property Let Role(ByVal Role As VTHttpaSessionRole)
   mRole = Role
End Property
Public Property Get Role() As VTHttpaSessionRole

    Role = mRole
End Property
' ============================================================
' Properties (수정됨)
' ============================================================

Public Property Get CID() As LongLong
    CID = mIid
End Property
Public Property Get IID() As LongLong
    IID = mIid
End Property
Public Property Get key() As String

    key = mKey
End Property
Public Property Let IsRequestSent(ByVal b As Boolean)

   mRequestSent = b
   
End Property

Public Property Get IsRequestSent() As Boolean

   IsRequestSent = mRequestSent
   
End Property

Public Property Let IsResponseSent(ByVal b As Boolean)

   mResponseSent = b
   
End Property

Public Property Get IsResponseSent() As Boolean

   IsResponseSent = mResponseSent
   
End Property

' Socket Property (Set 사용 - 객체이므로)
Public Property Set Socket(ByVal sock As CVtHttpaSocket)
    Set mSocket = sock  ' ? Set 키워드 추가
End Property

Public Property Get Socket() As CVtHttpaSocket
    Set Socket = mSocket  ' ? Set 키워드 추가 (객체 반환)
End Property

' EventType Property
Public Property Get EventType() As VTHttpaEventType
    EventType = mEventType
End Property

Public Property Let EventType(ByVal etype As VTHttpaEventType)
    mEventType = etype
End Property
Public Property Get EventId() As LongLong

    EventId = mEventId
End Property

Public Property Let EventId(ByVal eid As LongLong)
   mEventId = eid
   
End Property


Public Property Set CustomData(ByVal Data As Object)

   Set mCustomData = Data
   
End Property

Public Property Get CustomData() As Object

   Set CustomData = mCustomData
End Property
' Request Property (Set 사용 - 객체이므로)
Public Property Set Request(ByVal req As CVtHttpaRequest)
    Set mRequest = req  ' ? Set 키워드 추가
End Property

Public Property Get Request() As CVtHttpaRequest
    Set Request = mRequest  ' ? Set 키워드 추가 (객체 반환)
End Property

' Response Property (오타 수정: responset → response)
Public Property Set Response(ByVal Res As CVtHttpaResponse)  ' ? 오타 수정
    Set mResponse = Res  ' ? Set 키워드 추가
End Property

Public Property Get Response() As CVtHttpaResponse
    Set Response = mResponse  ' ? Set 키워드 추가 (객체 반환)
End Property

Public Property Set Websocket(ByVal ws As CVtHttpaWebsocket)
    Set mWebsocket = ws  ' ? Set 키워드 추가
        
End Property

Public Property Get Websocket() As CVtHttpaWebsocket
    Set Websocket = mWebsocket  ' ? Set 키워드 추가 (객체 반환)
End Property


' ErrorMsg Property (Let 사용 - String이므로)
Public Property Let ErrorMsg(ByVal emsg As String)  ' ? Let으로 수정
    mErrorMsg = emsg
End Property

Public Property Get ErrorMsg() As String
    ErrorMsg = mErrorMsg
End Property

' TimeoutReason Property (Let 사용 - String이므로)
Public Property Let TimeoutReason(ByVal Reason As String)  ' ? Let으로 수정
    mTimeoutReason = Reason
End Property

Public Property Get TimeoutReason() As String
    TimeoutReason = mTimeoutReason
End Property


