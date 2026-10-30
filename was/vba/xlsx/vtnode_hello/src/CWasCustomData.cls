VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWasCustomData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'Begin
'  MultiUse = -1  ' True
'End
'Attribute VB_Name = "CVtHttpaSession" ' (또는 해당 클래스 명칭)
Option Explicit

' [교정] 행 번호는 문자열이 아니라 숫자로 제어해야 메모리 및 인덱싱 성능에 유리합니다.
Private mSheetRowIndex As Long        ' Row of Sheet


Private mRequestTime As SYSTEMTIME
Private mRequestTimestamp As String
' [신규] 정밀 소요 시간 측정을 위한 내부 멤버 변수 (VBA Timer 스펙 대응)
Private mBeginTime     As Double      ' 트랜잭션 시작 타임스탬프
Private mEndTime       As Double      ' 트랜잭션 종료 타임스탬프




Private mResponseStatus    As String

Private mIsRequest As Boolean


Private Sub Class_Initialize()
    mSheetRowIndex = 0
    mBeginTime = 0#
    mEndTime = 0#
    mResponseStatus = ""
    
End Sub

Private Sub Class_Terminate()
    ' 객체 소멸 시 자원 정리 공간
End Sub

Public Property Get IsRequest() As Boolean

    IsRequest = mIsRequest
    
End Property

Public Property Let IsRequest(ByVal bIsRequest As Boolean)

    mIsRequest = bIsRequest
    
End Property


' ==========================================================================
' 1. 기존 SheetRowIndex 프로퍼티 (타입 변환 교정 및 명시적 반환 타입 정의)
' ==========================================================================
Public Property Get SheetRowIndex() As Long
    SheetRowIndex = mSheetRowIndex
End Property

Public Property Let SheetRowIndex(ByVal rowIndex As Long)
    mSheetRowIndex = rowIndex
End Property


' ==========================================================================
' 2. 신규 트랜잭션 시간 세팅용 프로퍼티 (Begin / End)
' ==========================================================================
Public Property Get BeginTime() As Double
    BeginTime = mBeginTime
    
End Property

Public Property Let BeginTime(ByVal vBegin As Double)
    mBeginTime = vBegin

    ' Win32 API를 통해 정밀한 시스템 시간 획득
    Call GetSystemTime(mRequestTime)
     ' 시차 조정(UTC -> KST +9) 및 포맷팅
    ' 단, 단순히 가볍게 초 단위까지만 관리하고 싶다면 Format(Now, "HH:mm:ss")로 대체하셔도 됩니다.
    Dim ms As String
    Dim kstHour As Integer
    kstHour = (mRequestTime.wHour + 9) Mod 24
    
    ms = Right("00" & mRequestTime.wMilliseconds, 3)
    mRequestTimestamp = Format(kstHour, "00") & ":" & _
                          Format(mRequestTime.wMinute, "00") & ":" & _
                          Format(mRequestTime.wSecond, "00") & "." & ms

End Property

Public Property Get EndTime() As Double
    EndTime = mEndTime
End Property

Public Property Let EndTime(ByVal vEnd As Double)
    mEndTime = vEnd
End Property

Public Property Get RequestTimestamp() As String

    RequestTimestamp = mRequestTimestamp
    
End Property

' ==========================================================================
' 3. [핵심 브리지] 소요 시간 계산 읽기 전용 프로퍼티 (ElapsedTime)
' ==========================================================================
''' <summary>
''' 트랜잭션의 순수 소요 시간(ms 단위 포맷 적용 가능 데이터)을 초 단위 소수로 반환합니다.
''' 예: 0.021 -> 21ms 소요됨을 의미
''' </summary>
Public Property Get ElapsedTime() As Double
    If mEndTime >= mBeginTime And mBeginTime > 0 Then
        ' 자정(Midnight)을 넘어갈 때 Timer 리셋 방어 코드 포함
        ElapsedTime = mEndTime - mBeginTime
    Else
        ElapsedTime = 0#
    End If
End Property


Public Property Let ResponseStatus(ByVal sts As String)
    
    mResponseStatus = sts
End Property

Public Property Get ResponseStatus() As String
    ResponseStatus = mResponseStatus
    
End Property
