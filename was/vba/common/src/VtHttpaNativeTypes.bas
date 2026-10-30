Attribute VB_Name = "VtHttpaNativeTypes"
Option Explicit
' ============================================================================
' VTNode HTTPA Native Types
' ============================================================================
'
' IMPORTANT: DO NOT MODIFY THE ENUM VALUES IN THIS MODULE.
'
' These Enum values must exactly match the corresponding C Enum values
' defined in the VTNode DLL.
'
' Changing these values may cause unexpected behavior or compatibility
' issues between the VBA application and the native DLL.
'
' This module is maintained by VTNode and should only be updated when
' the corresponding native API definitions are changed.
'
' For application-level Enum definitions, use VtHttpaTypes.bas instead.
'
' ============================================================================
Public Enum VTHttpaEventType
    ' --- 공통 및 에러 ---
    VTHTTPA_EVENT_NONE = 0
    VTHTTPA_EVENT_ERROR = 1        ' 에러 발생 (메시지 포인터 전달)

    ' --- 네트워크 수준 (L4) ---
    VTHTTPA_EVENT_NET_ACCEPTED = 2    ' 서버가 연결을 수락함
    VTHTTPA_EVENT_NET_CONNECTING = 3  ' 클라이언트 접속 시작
    VTHTTPA_EVENT_NET_CONNECT_FAILED = 4 ' 클라이언트 접속 실패
    VTHTTPA_EVENT_NET_CONNECTED = 5   ' 클라이언트 접속 성공
    VTHTTPA_EVENT_NET_DISCONNECTED = 6 ' 연결 끊김
    VTHTTPA_EVENT_NET_ERROR = 7       ' 네트워크 레벨 에러

    ' --- HTTP 트랜잭션 수준 (L7) ---
    VTHTTPA_EVENT_HT_ERROR = 8        ' HTTP 트랜잭션 에러
    VTHTTPA_EVENT_HT_TIMEOUT = 9      ' HTTP 트랜잭션 타임아웃

    VTHTTPA_EVENT_HT_TRANSACTION_BEGIN = 10 ' 트랜잭션 시작 (헤더 파싱 완료)
    VTHTTPA_EVENT_HT_TRANSACTION_END = 11 ' 트랜잭션 종료 (요청/응답 완료)

    ' --- HTTP 서버측 ---
    VTHTTPA_EVENT_HT_REQUEST_COMPLETED = 12 ' 요청 수신 완료 (Body 포함)

    ' --- HTTP 클라이언트측 ---
    VTHTTPA_EVENT_HT_RESPONSE_COMPLETED = 13 ' 응답 수신 완료 (Body 포함)

    ' --- WebSocket ---
    VTHTTPA_EVENT_WS_ESTABLISHED = 14  ' WebSocket 연결 확립
    VTHTTPA_EVENT_WS_TEXT_RECEIVED = 15 '텍스트 메시지 수신 (재조립 완료)
    VTHTTPA_EVENT_WS_BINARY_RECEIVED = 16 ' 바이너리 메시지 수신 (재조립 완료)
    VTHTTPA_EVENT_WS_CLOSED = 17       ' WebSocket 닫힘
    VTHTTPA_EVENT_WS_TIMEOUT = 18      ' WebSocket 타임아웃
    VTHTTPA_EVENT_WS_ERROR = 19        ' WebSocket 에러
End Enum

' Don't Modify
' it must be equal C Enum Value
Public Enum VTHttpaHeaderMergePolicy
    VTHTTPA_MERGE_SINGLE = 0      ' 1개만 허용(중복=에러)
    VTHTTPA_MERGE_COMMA = 1         ' 여러 개면 "a, b"로 병합
    VTHTTPA_MERGE_SEMICOLON = 2     ' Cookie처럼 "a; b"로 병합
    VTHTTPA_MERGE_MULTI_KEEP = 3    ' 여러 개를 그대로 보존(요청에선 거의 안 씀)
End Enum


