VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
    Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
    Private Declare PtrSafe Function FlushFileBuffers Lib "kernel32" (ByVal hFile As LongPtr) As Long
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
#Else
    Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
    Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
#End If

Private Const FILE_APPEND_DATA As Long = &H4
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const OPEN_ALWAYS As Long = 4
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

Private mLogDirectory As String
Private mCurrentLogDate As Date
Private mCurrentLogPath As String
Private mIsOpen As Boolean

Private mConsoleOut As Boolean

#If VBA7 Then
    Private mFileHandle As LongPtr
#Else
    Private mFileHandle As Long
#End If

Private Sub Class_Initialize()
    mLogDirectory = ThisWorkbook.Path & Application.PathSeparator & "logs"
    mConsoleOut = True
    
    
    
End Sub

Private Sub Class_Terminate()
    CloseLogFile
End Sub

Public Property Get LogDirectory() As String
    LogDirectory = mLogDirectory
End Property

Public Property Let LogDirectory(ByVal value As String)
    Dim newDirectory As String

    newDirectory = Trim$(value)
    If Len(newDirectory) = 0 Then Err.Raise 5, "CVtLog.LogDirectory", "Log directory cannot be empty."

    If StrComp(mLogDirectory, newDirectory, vbTextCompare) <> 0 Then
        CloseLogFile
        mLogDirectory = RemoveTrailingSeparator(newDirectory)
    End If
End Property

Public Property Get CurrentLogPath() As String
    CurrentLogPath = mCurrentLogPath
End Property

Public Sub info(ByVal sourceFileName As String, ByVal functionName As String, Optional ByVal message As String = "")
    WriteLog "INFO", sourceFileName, functionName, message
End Sub

Public Sub Error(ByVal sourceFileName As String, ByVal functionName As String, Optional ByVal message As String = "")
    WriteLog "ERROR", sourceFileName, functionName, message
End Sub

Public Sub Warn(ByVal sourceFileName As String, ByVal functionName As String, Optional ByVal message As String = "")
    WriteLog "WARN", sourceFileName, functionName, message
End Sub

Public Sub Trace(ByVal sourceFileName As String, ByVal functionName As String, Optional ByVal message As String = "")
    WriteLog "DEBUG", sourceFileName, functionName, message
End Sub

Public Sub CloseLogFile()
    On Error Resume Next
    If mIsOpen Then CloseHandle mFileHandle
    mIsOpen = False
    mFileHandle = 0
    mCurrentLogPath = vbNullString
    mCurrentLogDate = 0
    On Error GoTo 0
End Sub

Private Sub WriteLog(ByVal logType As String, ByVal sourceFileName As String, ByVal functionName As String, ByVal message As String)
    Dim currentDate As Date
    Dim logLine As String

    currentDate = Date
    EnsureLogFile currentDate

    logLine = Format$(Now, "yyyy-mm-dd hh:nn:ss") _
            & " [" & UCase$(logType) & "]" _
            & " [" & SafeText(sourceFileName) & "]" _
            & " [" & SafeText(functionName) & "]"

    If Len(message) > 0 Then logLine = logLine & " " & NormalizeMessage(message)

    AppendUtf8Line logLine
    
    
    If mConsoleOut Then
       Debug.Print logLine
    End If
    
    
    
End Sub

Private Sub EnsureLogFile(ByVal logDate As Date)
    If mIsOpen Then
        If DateValue(mCurrentLogDate) = DateValue(logDate) Then Exit Sub
        CloseLogFile
    End If

    EnsureDirectoryExists mLogDirectory

    mCurrentLogDate = DateValue(logDate)
    mCurrentLogPath = mLogDirectory & Application.PathSeparator _
                    & "vtlog_" & Format$(mCurrentLogDate, "yyyy-mm-dd") & ".log"
    mFileHandle = CreateFileW(StrPtr(mCurrentLogPath), FILE_APPEND_DATA, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If mFileHandle = -1 Then
        mCurrentLogPath = vbNullString
        mCurrentLogDate = 0
        Err.Raise vbObjectError + 7101, "CVtLog.EnsureLogFile", "Unable to open UTF-8 log file."
    End If
    mIsOpen = True
End Sub

Private Sub AppendUtf8Line(ByVal logLine As String)
    Dim utf8Bytes() As Byte
    Dim byteCount As Long
    Dim bytesWritten As Long
    Dim writeResult As Long

    utf8Bytes = ToUtf8WithoutBom(logLine & vbCrLf)
    byteCount = ByteArrayLength(utf8Bytes)
    If byteCount = 0 Then Exit Sub

    writeResult = WriteFile(mFileHandle, utf8Bytes(LBound(utf8Bytes)), byteCount, bytesWritten, 0)
    If writeResult = 0 Or bytesWritten <> byteCount Then
        Err.Raise vbObjectError + 7102, "CVtLog.AppendUtf8Line", "Unable to write the complete UTF-8 log entry."
    End If

    FlushFileBuffers mFileHandle
End Sub

Private Function ToUtf8WithoutBom(ByVal value As String) As Byte()
    Const AD_TYPE_BINARY As Long = 1
    Const AD_TYPE_TEXT As Long = 2

    Dim stream As Object
    Dim result() As Byte

    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = AD_TYPE_TEXT
        .Charset = "utf-8"
        .Open
        .WriteText value
        .Position = 0
        .Type = AD_TYPE_BINARY
        .Position = 3
        result = .Read
        .Close
    End With

    ToUtf8WithoutBom = result
End Function

Private Function ByteArrayLength(ByRef value() As Byte) As Long
    On Error GoTo EmptyArray
    ByteArrayLength = UBound(value) - LBound(value) + 1
    Exit Function

EmptyArray:
    ByteArrayLength = 0
End Function

Private Sub EnsureDirectoryExists(ByVal directoryPath As String)
    If Len(dir$(directoryPath, vbDirectory)) > 0 Then Exit Sub
    MkDir directoryPath
End Sub

Private Function RemoveTrailingSeparator(ByVal value As String) As String
    Dim result As String

    result = value
    Do While Len(result) > 3
        If Right$(result, 1) <> "\" And Right$(result, 1) <> "/" Then Exit Do
        result = left$(result, Len(result) - 1)
    Loop
    RemoveTrailingSeparator = result
End Function

Private Function SafeText(ByVal value As String) As String
    If Len(Trim$(value)) = 0 Then
        SafeText = "-"
    Else
        SafeText = NormalizeMessage(value)
    End If
End Function

Private Function NormalizeMessage(ByVal value As String) As String
     NormalizeMessage = value
     
'    Dim result As String
'
'    result = Replace$(value, vbCrLf, " ")
'    result = Replace$(result, vbCr, " ")
'    result = Replace$(result, vbLf, " ")
'    NormalizeMessage = result
End Function


