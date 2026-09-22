Attribute VB_Name = "TestHttpa"
Option Explicit

'Private mApp As CHttpaMyApp
Private m_bStopEventLoop As Boolean
Private Function StartHttpaEngine() As Boolean
    Dim homeDir As String
    Dim iniFile As String
    Dim result As Boolean

    homeDir = ThisWorkbook.Path
    iniFile = VtPathCombine(homeDir, "vtnode_httpa_params.ini")

    result = VtHttpaEngineRun(homeDir, iniFile)

    Debug.Print "VtHttpaEngineRun()->" & result

   
    StartHttpaEngine = result
End Function
Private Sub StopHttpaEngine()
    Call VtHttpaEngineShutdown
End Sub



Public Sub TestCHttpa_Start2()
    Dim success As Boolean
    success = StartHttpaEngine()
End Sub


Public Sub TestCHttpa_StartEventLoop()

   

    Debug.Print "이벤트 루프 시작..."
    m_bStopEventLoop = False

    Dim Res As Boolean
    Res = StartHttpaEngine
       
    If Res Then

        If mApp Is Nothing Then
           Set mApp = New CHttpaMyApp
        End If
        
        
        Do While Not m_bStopEventLoop
            Call mApp.Poll
            DoEvents
        Loop
    
        Call StopHttpaEngine
        
        Set mApp = Nothing
                
    End If

    Debug.Print "이벤트 루프 종료됨."
End Sub

Public Sub TestCHttpa_StopEventLoop()
    m_bStopEventLoop = True
    Debug.Print "이벤트 루프 중단 명령 전달됨."
End Sub

Public Sub TestCHttpa_DestroyApp()
    Set mApp = Nothing
End Sub

Sub ClearDebug()
    Application.VBE.Windows("Immediate").SetFocus
    SendKeys "^a"
    SendKeys "{DEL}"
    
End Sub
    
Sub ClearImmediateWindow()
    Application.SendKeys "^a"
    Application.SendKeys "{DEL}"
    
End Sub

Private Sub Parse(ByVal hQuery As CVtHttpaQuery, ByVal s As String)
   hQuery.Clear
   
   Debug.Print "Parse:" & s
   
   hQuery.Parse s
   Call hQuery.ToSummary
   Dim encoded As String
   encoded = hQuery.ToString
   
   Debug.Print "encoded:" & encoded
   

End Sub
Public Sub TestQuery()

   Dim hQuery As New CVtHttpaQuery
   
   
   'Call ClearImmediateWindow
   
   hQuery.AddParam "key.a", "value.a"
   hQuery.AddParam "key.a", "value.a1"
   hQuery.AddParam "key.b", "value.b"
   
   Debug.Print "key.b=" & hQuery.GetParam("key.b")
   Debug.Print "key.a=" & hQuery.GetParam("key.a")
   
   
   
   Parse hQuery, "name=kim+seo&city=Seoul&msg=hello+world%21&tag=a&tag=b"
   Parse hQuery, "a=a&tag=a&tag=b"
   
   
   
   Set hQuery = Nothing
   
   
   
   
   
   
   
End Sub


Private Sub PrintHArray(ByVal hArray As Collection, ByVal title As String)

  Debug.Print "Print:" & title
  
 If Not hArray Is Nothing Then
    
    ' Collection이나 Array인 경우에만 루프 진입
    If IsArray(hArray) Or IsObject(hArray) Then
        Dim v As Variant
        For Each v In hArray
            ' v 자체가 객체일 경우를 대비하여 한 번 더 방어
            If IsObject(v) Then
                Debug.Print "h1=[Object: " & TypeName(v) & "]"
            Else
                ' Null이나 에러 값이 아닐 때만 출력
                If Not IsError(v) Then
                    Debug.Print "h1=" & CStr(v)
                End If
            End If
        Next
    End If
    
End If

End Sub
Public Sub TestHeader()

 Dim hHeaders As New CVtHttpaHeaders
 Dim hvalue As String
 Dim hValues As Collection
 hHeaders.AddHeader "h1", "value1"
 hHeaders.AddHeader "h2", "value2"
 
 hHeaders.AddHeader "h3", "value3"
 hHeaders.AddHeader "h1", "value11"
 
 Set hValues = hHeaders.GetHeaders("h1")
 
 PrintHArray hValues, "h1.array"
 
 
 Dim hString As String
 
 hString = hHeaders.ToString
 Debug.Print "String=" & hString
 
 hHeaders.RemoveHeader "h2"

 hString = hHeaders.ToString
 Debug.Print "h2.removed.string=" & hString
 
 
 
 
 hvalue = hHeaders.GetHeader("h1")
 Debug.Print "h1=" & hvalue
 
 
 
 
 
 
 
 hvalue = hHeaders.GetHeader("h2")
 Debug.Print "h2=" & hvalue
 
 
 
End Sub

Public Sub TestVtWasInit()


Dim success As Boolean
Dim initResult As VtWas.VTWasInitReult



success = VtWas.VtWasInit(initResult)


End Sub

Public Sub TestVtWasUninit()

Call VtWas.VtWasUnInit

End Sub


Public Sub TestGetRandomString()
    Dim Res As Long
    Dim pStr As LongPtr    ' 문자열 시작 주소를 받을 포인터
    Dim strLen As Long     ' 문자열 길이를 받을 변수
    
    
    Call VtNode.VtNodeLoadEngineDLL
    
    ' 1. DLL 함수 호출 (참조 전달을 통해 주소와 길이를 받아옴)
    Res = VtHttpaContents.vtnode_get_random_string(pStr, strLen)
    
    If Res <> 0 And pStr <> 0 And strLen > 0 Then
        
        Dim str  As String
        Dim Bytes() As Byte
        
        Debug.Print "TestGetRandomString.len=" & strLen
        
        Bytes = VtUtf8PtrToBytes(pStr, strLen)
        
        str = VtUtf8BytesToString(Bytes)
        
        
        Debug.Print "result=" & str
    Else
        MsgBox "랜덤 문자열 생성 실패 또는 잘못된 데이터 전달", vbCritical
    End If
End Sub

