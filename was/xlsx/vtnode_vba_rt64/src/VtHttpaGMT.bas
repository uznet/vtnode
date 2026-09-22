Attribute VB_Name = "VtHttpaGMT"
Option Explicit

#If VBA7 And Win64 Then

#If VTNODE_DEBUG = 1 Then

    ' GMT 문자열 → time_t 변환
    Private Declare PtrSafe Function vtnode_httpa_parse_gmt Lib "vtnode64d.dll" ( _
        ByVal gmt As LongPtr) As LongLong
    
    ' time_t → GMT 문자열 변환
    Private Declare PtrSafe Function vtnode_httpa_make_gmt Lib "vtnode64d.dll" ( _
        ByVal t As LongLong, _
        ByVal gmtOut As LongPtr, _
        ByVal gmtSize As Long) As Long
#Else
    ' GMT 문자열 → time_t 변환
    Private Declare PtrSafe Function vtnode_httpa_parse_gmt Lib "vtnode64.dll" ( _
        ByVal gmt As LongPtr) As LongLong
    
    ' time_t → GMT 문자열 변환
    Private Declare PtrSafe Function vtnode_httpa_make_gmt Lib "vtnode64.dll" ( _
        ByVal t As LongLong, _
        ByVal gmtOut As LongPtr, _
        ByVal gmtSize As Long) As Long

#End If
#End If

' ============================================================
' VBA 래퍼 함수들
' ============================================================

' GMT 문자열 → time_t 변환
Public Function VtHttpaParseGmt(ByVal gmtString As String) As LongLong
    Dim utf8Gmt() As Byte
    utf8Gmt = VtStringToUtf8Bytes(gmtString)
    VtHttpaParseGmt = vtnode_httpa_parse_gmt(VarPtr(utf8Gmt(0)))
End Function

' time_t → GMT 문자열 변환
Public Function VtHttpaMakeGmt(ByVal t As LongLong) As String
    Dim buffer(0 To 127) As Byte  ' GMT 문자열은 보통 30바이트 이하
    Dim result As Long

    result = vtnode_httpa_make_gmt(t, VarPtr(buffer(0)), 128)

    If result <> 0 Then
        VtHttpaMakeGmt = VtUtf8BytesToString(buffer)
    Else
        VtHttpaMakeGmt = ""
    End If
End Function
