Attribute VB_Name = "VtHttpaEvent"
' --- [ modVtnode.bas / Private Declares ] ---
Option Explicit

#If VBA7 And Win64 Then

#If VTNODE_DEBUG = 1 Then

    ' --- 이벤트 폴링 ---
    'Private Declare PtrSafe Function vtnode_httpa_event_dequeue Lib "vtnode64d.dll" (ByVal out As LongPtr, ByVal capacity As Long) As Long
    
    ' 1. 이벤트 타입 가져오기
    Private Declare PtrSafe Function vtnode_httpa_event_get_type Lib "vtnode64d.dll" (ByVal eid As LongLong) As VTHttpaEventType
    
    ' 2. 이벤트 소유 인스턴스 ID 가져오기
    Private Declare PtrSafe Function vtnode_httpa_event_get_instanceid Lib "vtnode64d.dll" (ByVal eid As LongLong) As LongLong
    
    ' 3. 이벤트 메모리 해제 (매우 중요: 처리 후 반드시 호출)
    Private Declare PtrSafe Sub vtnode_httpa_event_free Lib "vtnode64d.dll" (ByVal eid As LongLong)
    
    ' 4. 에러 메시지 포인터 가져오기
    Private Declare PtrSafe Function vtnode_httpa_event_get_error_message Lib "vtnode64d.dll" (ByVal eid As LongLong) As LongPtr
    
    ' 5. 타임아웃 사유 포인터 가져오기
    Private Declare PtrSafe Function vtnode_httpa_event_get_timeout_reason Lib "vtnode64d.dll" (ByVal eid As LongLong) As LongPtr
    
    ' 6. 원격지 주소 및 포트 정보
    Private Declare PtrSafe Function vtnode_httpa_event_get_remote_addr Lib "vtnode64d.dll" (ByVal eid As LongLong) As LongPtr
    Private Declare PtrSafe Function vtnode_httpa_event_get_remote_port Lib "vtnode64d.dll" (ByVal eid As LongLong) As Long
    
    ' 7. 로컬 주소 및 포트 정보
    Private Declare PtrSafe Function vtnode_httpa_event_get_local_addr Lib "vtnode64d.dll" (ByVal eid As LongLong) As LongPtr
    Private Declare PtrSafe Function vtnode_httpa_event_get_local_port Lib "vtnode64d.dll" (ByVal eid As LongLong) As Long
#Else

    ' --- 이벤트 폴링 ---
    Private Declare PtrSafe Function vtnode_httpa_event_dequeue Lib "vtnode64.dll" (ByVal out As LongPtr, ByVal Capacity As Long) As Long
    ' 1. 이벤트 타입 가져오기
    Private Declare PtrSafe Function vtnode_httpa_event_get_type Lib "vtnode64.dll" (ByVal eid As LongLong) As VTHttpaEventType
    
    ' 2. 이벤트 소유 인스턴스 ID 가져오기
    Private Declare PtrSafe Function vtnode_httpa_event_get_instanceid Lib "vtnode64.dll" (ByVal eid As LongLong) As LongLong
    
    ' 3. 이벤트 메모리 해제 (매우 중요: 처리 후 반드시 호출)
    Private Declare PtrSafe Sub vtnode_httpa_event_free Lib "vtnode64.dll" (ByVal eid As LongLong)
    
    ' 4. 에러 메시지 포인터 가져오기
    Private Declare PtrSafe Function vtnode_httpa_event_get_error_message Lib "vtnode64.dll" (ByVal eid As LongLong) As LongPtr
    
    ' 5. 타임아웃 사유 포인터 가져오기
    Private Declare PtrSafe Function vtnode_httpa_event_get_timeout_reason Lib "vtnode64.dll" (ByVal eid As LongLong) As LongPtr
    
    ' 6. 원격지 주소 및 포트 정보
    Private Declare PtrSafe Function vtnode_httpa_event_get_remote_addr Lib "vtnode64.dll" (ByVal eid As LongLong) As LongPtr
    Private Declare PtrSafe Function vtnode_httpa_event_get_remote_port Lib "vtnode64.dll" (ByVal eid As LongLong) As Long
    
    ' 7. 로컬 주소 및 포트 정보
    Private Declare PtrSafe Function vtnode_httpa_event_get_local_addr Lib "vtnode64.dll" (ByVal eid As LongLong) As LongPtr
    Private Declare PtrSafe Function vtnode_httpa_event_get_local_port Lib "vtnode64.dll" (ByVal eid As LongLong) As Long

#End If

#End If


'Public Function VtHttpaEventDequeueInto(ByRef arr() As LongLong) As Long
'  Dim Capacity As Long
'
'  Capacity = UBound(arr) - LBound(arr) + 1
'
'  VtHttpaEventDequeueInto = vtnode_httpa_event_dequeue(VarPtr(arr(LBound(arr))), Capacity)
'
'End Function

' 1개의 이벤트를 꺼내어 LongLong 핸들로 반환합니다.
' 반환값이 0이면 처리 대기 중인 이벤트가 없음을 의미합니다.
'Public Function VtHttpaEventDequeue() As LongLong
'    Dim singleArr(0 To 0) As LongLong
'    Dim Count As Long
'
'    ' 앞서 만든 DequeueInto를 활용하여 1개의 요소를 채웁니다.
'    Count = VtHttpaEventDequeueInto(singleArr)
'
'    ' 이벤트가 성공적으로 추출되었다면 해당 ID를 반환
'    If Count > 0 Then
'        VtHttpaEventDequeue = singleArr(0)
'    Else
'        VtHttpaEventDequeue = 0
'    End If
'End Function

Public Function VtHttpaEventGetType(ByVal EventId As LongLong) As VTHttpaEventType
    VtHttpaEventGetType = vtnode_httpa_event_get_type(EventId)
End Function
Public Function VtHttpaEventGetInstanceId(ByVal EventId As LongLong) As LongLong
    VtHttpaEventGetInstanceId = vtnode_httpa_event_get_instanceid(EventId)
End Function
Public Sub VtHttpaEventFree(ByVal EventId As LongLong)
    Call vtnode_httpa_event_free(EventId)
End Sub

Public Function VtHttpaEventGetErrorMessage(ByVal EventId As LongLong) As String
    Dim pMsg As LongPtr
    pMsg = vtnode_httpa_event_get_error_message(EventId)
    VtHttpaEventGetErrorMessage = VtUtf8PtrToString(pMsg)
End Function

Public Function VtHttpaEventGetTimeoutReason(ByVal EventId As LongLong) As String
    Dim pReason As LongPtr
    pReason = vtnode_httpa_event_get_timeout_reason(EventId)
    VtHttpaEventGetTimeoutReason = VtUtf8PtrToString(pReason)
End Function

Public Function VtHttpaEventGetRemoteAddr(ByVal EventId As LongLong) As String
    Dim pAddr As LongPtr
    pAddr = vtnode_httpa_event_get_remote_addr(EventId)
    VtHttpaEventGetRemoteAddr = VtUtf8PtrToString(pAddr)
End Function

Public Function VtHttpaEventGetRemotePort(ByVal EventId As LongLong) As Long
    VtHttpaEventGetRemotePort = vtnode_httpa_event_get_remote_port(EventId)
End Function


Public Function VtHttpaEventGetLocalAddr(ByVal EventId As LongLong) As String
    Dim pAddr As LongPtr
    pAddr = vtnode_httpa_event_get_local_addr(EventId)
    VtHttpaEventGetLocalAddr = VtUtf8PtrToString(pAddr)
End Function

Public Function VtHttpaEventGetLocalPort(ByVal EventId As LongLong) As Long
    VtHttpaEventGetLocalPort = vtnode_httpa_event_get_local_port(EventId)
End Function
