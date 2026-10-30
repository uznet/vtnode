Attribute VB_Name = "VtNetwork"
Option Explicit

#If VBA7 And Win64 Then

    #If VTNODE_DEBUG = 1 Then

        Private Declare PtrSafe Function vtnode_net_get_localip Lib "vtnode64d.dll" ( _
            ByVal LocalIp As LongPtr, _
            ByVal localIpSize As Long, _
            ByVal eout As LongPtr, _
            ByVal esize As Long) As Long

        Private Declare PtrSafe Function vtnode_net_get_publicip Lib "vtnode64d.dll" ( _
            ByVal publicIp As LongPtr, _
            ByVal publicIpSize As Long, _
            ByVal eout As LongPtr, _
            ByVal esize As Long) As Long

    #Else

        Private Declare PtrSafe Function vtnode_net_get_localip Lib "vtnode64.dll" ( _
            ByVal LocalIp As LongPtr, _
            ByVal localIpSize As Long, _
            ByVal eout As LongPtr, _
            ByVal esize As Long) As Long

        Private Declare PtrSafe Function vtnode_net_get_publicip Lib "vtnode64.dll" ( _
            ByVal publicIp As LongPtr, _
            ByVal publicIpSize As Long, _
            ByVal eout As LongPtr, _
            ByVal esize As Long) As Long


    #End If

#End If

'IP 버퍼 및 에러 버퍼 크기 상수 선언 (문자 수 기준)
Private Const IP_BUFFER_SIZE As Long = 64
Private Const ERROR_BUFFER_SIZE As Long = 512

Public Function VtNetGetLocalIp(ByRef LocalIp As String, ByRef eout As String) As Boolean
    Dim ret As Long
    
    ' C의 char localIp[64] 와 100% 동일한 연속된 바이트 힙 메모리 스택 생성
    Dim bufIp() As Byte: ReDim bufIp(IP_BUFFER_SIZE - 1)
    Dim bufErr() As Byte: ReDim bufErr(ERROR_BUFFER_SIZE - 1)
    
    ' VarPtr()를 통해 바이트 배열 첫 번째 원소(index 0)의 물리 메모리 주소를 포인터로 주입
    ret = vtnode_net_get_localip(VarPtr(bufIp(0)), IP_BUFFER_SIZE, VarPtr(bufErr(0)), ERROR_BUFFER_SIZE)
    
    If ret >= 0 Then
        ' C에서 반환한 멀티바이트 바이트 스트림을 다시 VBA 유저폼용 유니코드 스트링으로 파싱
        Dim rawStr As String: rawStr = StrConv(bufIp, vbUnicode)
        LocalIp = left$(rawStr, InStr(rawStr, vbNullChar) - 1)
        eout = ""
        VtNetGetLocalIp = True
    Else
        Dim rawErr As String: rawErr = StrConv(bufErr, vbUnicode)
        LocalIp = ""
        eout = left$(rawErr, InStr(rawErr, vbNullChar) - 1)
        VtNetGetLocalIp = False
    End If
End Function

Public Function VtNetGetExternalIp(ByRef ExternalIp As String, ByRef eout As String) As Boolean

    Dim ret As Long
    
    ' C의 char localIp[64] 와 100% 동일한 연속된 바이트 힙 메모리 스택 생성
    Dim bufIp() As Byte: ReDim bufIp(IP_BUFFER_SIZE - 1)
    Dim bufErr() As Byte: ReDim bufErr(ERROR_BUFFER_SIZE - 1)
    
    ' VarPtr()를 통해 바이트 배열 첫 번째 원소(index 0)의 물리 메모리 주소를 포인터로 주입
    ret = vtnode_net_get_publicip(VarPtr(bufIp(0)), IP_BUFFER_SIZE, VarPtr(bufErr(0)), ERROR_BUFFER_SIZE)
    
    If ret >= 0 Then
        ' C에서 반환한 멀티바이트 바이트 스트림을 다시 VBA 유저폼용 유니코드 스트링으로 파싱
        Dim rawStr As String: rawStr = StrConv(bufIp, vbUnicode)
        ExternalIp = left$(rawStr, InStr(rawStr, vbNullChar) - 1)
        eout = ""
        VtNetGetExternalIp = True
    Else
        Dim rawErr As String: rawErr = StrConv(bufErr, vbUnicode)
        ExternalIp = ""
        eout = left$(rawErr, InStr(rawErr, vbNullChar) - 1)
        VtNetGetExternalIp = False
    End If

End Function



