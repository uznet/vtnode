Attribute VB_Name = "MyHttpaNetEvent"
Option Explicit


Private Const MODULE_NAME As String = "MyHttpaNetEvent"

Private mStartedAt As LongLong

Public Sub Init()

    'Set mFormVTNode = frmVTNode
    mStartedAt = VtTimeUtils.VtGetCurrentCTime()

End Sub


Public Sub Uninit()

    mStartedAt = 0

End Sub






