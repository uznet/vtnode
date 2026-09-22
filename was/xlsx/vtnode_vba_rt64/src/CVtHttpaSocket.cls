VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mRemoteAddr As String
Private mRemotePort As Long
Private mLocalAddr  As String
Private mLocalPort  As Long

Private Sub Clear()
    mRemoteAddr = ""
    mRemotePort = 0
    
    mLocalAddr = ""
    mLocalPort = 0

End Sub
Private Sub Class_Initialize()
  Call Clear
  
    
End Sub

Private Sub Class_Terminate()
  Call Clear
End Sub

Public Sub initFromEvent(ByVal eid As LongLong)

   mRemoteAddr = VtHttpaEventGetRemoteAddr(eid)
   mRemotePort = VtHttpaEventGetRemotePort(eid)
   
   mLocalAddr = VtHttpaEventGetLocalAddr(eid)
   mLocalPort = VtHttpaEventGetLocalPort(eid)
   
   
End Sub


Public Property Get RemoteAddr() As String

   RemoteAddr = mRemoteAddr
End Property

Public Property Get RemotePort() As Long
   RemotePort = mRemotePort
End Property

Public Property Get LocalAddr() As String

   LocalAddr = mLocalAddr
End Property

Public Property Get LocalPort() As Long
   LocalPort = mLocalPort
End Property


Public Function ToSummary() As String
    ToSummary = "Remote: " & mRemoteAddr & ":" & mRemotePort & ", Local: " & mLocalAddr & ":" & mLocalPort
End Function

