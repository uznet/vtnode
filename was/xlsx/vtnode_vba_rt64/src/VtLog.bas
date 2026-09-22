Attribute VB_Name = "VtLog"
Option Explicit

Private Const MODULE_NAME As String = "VtLog"
Private mVtLog As CVtLog

Public Sub VtLogInitialize(Optional ByVal LogDirectory As String = "")
   Const PROC_NAME As String = "VtLogInitialize"
    If mVtLog Is Nothing Then
        Set mVtLog = New CVtLog
        
        Call mVtLog.info(MODULE_NAME, PROC_NAME, "VtLog Initialized")
        
    End If
    If Len(Trim$(LogDirectory)) > 0 Then mVtLog.LogDirectory = LogDirectory
End Sub

Public Function VtLogInstance() As CVtLog
    If mVtLog Is Nothing Then VtLogInitialize
    Set VtLogInstance = mVtLog
End Function

Public Sub VtLogClose()
     Const PROC_NAME As String = "VtLogClose"
    If Not mVtLog Is Nothing Then
         
        Call mVtLog.info(MODULE_NAME, PROC_NAME, "VtLog Close")
        mVtLog.CloseLogFile
        Set mVtLog = Nothing
    End If
End Sub

Public Sub VtLogInfo(ByVal moduleName As String, ByVal procName As String, Optional msg As String = "")
    If mVtLog Is Nothing Then VtLogInitialize
    
    mVtLog.info moduleName, procName, msg
    
    
End Sub

Public Sub VtLogError(ByVal moduleName As String, ByVal procName As String, Optional msg As String = "")
    If mVtLog Is Nothing Then VtLogInitialize
    
    mVtLog.Error moduleName, procName, msg

End Sub

Public Sub VtLogTrace(ByVal moduleName As String, ByVal procName As String, Optional msg As String = "")

    If mVtLog Is Nothing Then VtLogInitialize
    
    mVtLog.Trace moduleName, procName, msg

End Sub
Public Sub VtLogWarn(ByVal moduleName As String, ByVal procName As String, Optional msg As String = "")

    If mVtLog Is Nothing Then VtLogInitialize
    
    mVtLog.Warn moduleName, procName, msg

End Sub


Public Sub VtLogDemo()
    VtLogInitialize
    VtLogInfo "table-order.xlsm", "VtLogDemo", "Runtime initialized"
    VtLogWarn "table-order.xlsm", "VtLogDemo", "HTTPS certificate expires soon"
    VtLogError "table-order.xlsm", "VtLogDemo", "Example error message"
End Sub
