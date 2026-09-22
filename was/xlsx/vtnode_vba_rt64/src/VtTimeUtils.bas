Attribute VB_Name = "VtTimeUtils"
Option Explicit

' C time_t (Unix Epoch 초) → VBA Date 변환
Public Function VtCTimeToDate(ByVal cTime As LongLong) As Date
    On Error GoTo ErrorHandler

    ' Unix Epoch (1970-01-01) = VBA Date 25569
    ' 86400 = 하루 초 수
    VtCTimeToDate = CDbl(cTime) / 86400# + 25569#
    Exit Function

ErrorHandler:
    ' 오버플로우 등 예외 시 0 반환 (1899-12-30)
    VtCTimeToDate = 0
End Function

' VBA Date → C time_t (Unix Epoch 초) 변환
Public Function VtDateToCTime(ByVal vbaDate As Date) As LongLong
    On Error GoTo ErrorHandler

    ' Unix Epoch 이전 날짜 체크 (1970-01-01 = 25569)
    If CDbl(vbaDate) < 25569# Then
        ' Unix Epoch 이전은 0 반환 또는 에러
        VtDateToCTime = 0
        Exit Function
    End If

    ' 계산
#If VBA7 Then
        VtDateToCTime = CLngLng((CDbl(vbaDate) - 25569#) * 86400#)
#Else
    ' 32bit 환경 (2038년 문제 존재)
    VtDateToCTime = CLng((CDbl(vbaDate) - 25569#) * 86400#)
#End If

    Exit Function

ErrorHandler:
    ' 오버플로우 시 0 반환
    VtDateToCTime = 0
End Function
