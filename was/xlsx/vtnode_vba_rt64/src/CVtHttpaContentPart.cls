VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVtHttpaContentPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

' ============================================================
' CVtHttpaContentPart
'
' 역할:
'   - 하나의 HTTP/MIME Content Part를 VBA 객체로 표현
'   - Native ContentHandle에서 헤더와 데이터를 로드
'   - 파일, 텍스트, 바이너리 데이터로 새 Part 초기화
'
' 전제:
'   VTHttpaContentData 구조체가 다음 멤버를 제공한다고 가정합니다.
'
'     IsValid    As Boolean
'     IsFile     As Boolean
'     ContentType As String
'     DataLen    As Long
'     FilePath   As String
'     Bytes()    As Byte
'
' Native wrapper 함수:
'     VtHttpaContentGetData(ContentHandle) As VTHttpaContentData
'
' ============================================================

Private mHeaders As CVtHttpaHeaders
Private mContentData As VTHttpaContentData


' ============================================================
' Life cycle
' ============================================================

Private Sub Class_Initialize()

    Set mHeaders = New CVtHttpaHeaders
    Call ResetContentData

End Sub


Private Sub Class_Terminate()

    Clear
    Set mHeaders = Nothing

End Sub


' ============================================================
' Reset / Clear
' ============================================================

Public Sub Clear()

    If Not mHeaders Is Nothing Then
        mHeaders.Clear
    End If

    ResetContentData

End Sub

Public Function CreateEmpty() As Boolean

    On Error GoTo EH

    Clear


    CreateEmpty = True
    Exit Function

EH:
    CreateEmpty = False

End Function
Private Sub ResetContentData()

    On Error Resume Next

    Erase mContentData.Bytes

    mContentData.IsValid = False
    mContentData.DataType = VTHTTPA_CONTENT_DATATYPE_NONE
    mContentData.FilePath = vbNullString
    mContentData.ByteLen = 0

    mContentData.ContentType = vbNullString

    mContentData.ContentDisposition.DispType = vbNullString
    mContentData.ContentDisposition.DispName = vbNullString
    mContentData.ContentDisposition.DispFilename = vbNullString
    On Error GoTo 0

End Sub


' ============================================================
' Build from native ContentHandle
' ============================================================

Public Function InitFromContentHandle( _
    ByVal ContentHandle As LongLong) As Boolean

    On Error GoTo ErrorHandler

    Clear

    If ContentHandle = 0 Then
        Err.Raise vbObjectError + 100, _
                  "CVtHttpaContentPart.BuildFromContentHandle", _
                  "ContentHandle is zero."
    End If

    ' 헤더는 CVtHttpaHeaders가 Native handle로부터 직접 읽습니다.
    If Not mHeaders.BuildFromContentHandle(ContentHandle) Then
        Err.Raise vbObjectError + 101, _
                  "CVtHttpaContentPart.BuildFromContentHandle", _
                  "Failed to load content headers."
    End If

    ' 데이터 구조체 로드
    mContentData = VtHttpaContent.VtHttpaContentGetData(ContentHandle)

        
    
    If Not mContentData.IsValid Then
        Err.Raise vbObjectError + 102, _
                  "CVtHttpaContentPart.BuildFromContentHandle", _
                  "Failed to load content data."
    End If

    Debug.Print "InitFromContentHandle"


    

    ' ContentData에 ContentType이 비어 있으면 헤더에서 보완합니다.
    If Len(mContentData.ContentType) = 0 Then
        mContentData.ContentType = mHeaders.GetHeader("Content-Type")
    End If
    
    
    
    Dim BodyLength As Long
    BodyLength = Me.ByteLength
    
'    Debug.Print "InitFromContentHandle:IsValid=" & CStr(mContentData.IsValid)
'    Debug.Print "InitFromContentHandle:DataType=" & CStr(mContentData.DataType)
'    Debug.Print "InitFromContentHandle:Length=" & CStr(mContentData.ByteLen)
'
    
    

    

    InitFromContentHandle = True
    Exit Function

ErrorHandler:

    Debug.Print "CVtHttpaContentPart.BuildFromContentHandle error: " _
              & Err.Number & ", " & Err.Description

    Clear
    InitFromContentHandle = False

End Function



' ============================================================
' Properties
' ============================================================

Public Property Get Headers() As CVtHttpaHeaders

    Set Headers = mHeaders

End Property


Public Property Get ContentData() As VTHttpaContentData

    ContentData = mContentData

End Property

Public Property Get ContentName() As String

    ContentName = mContentData.ContentDisposition.DispName
    
End Property

Public Property Let ContentName(ByVal Name As String)
    
    mContentData.ContentDisposition.DispName = Name
        

End Property



Public Property Get IsValid() As Boolean

    IsValid = mContentData.IsValid

End Property


Public Property Get ContentType() As String

    If Not mContentData.IsValid Then
        ContentType = vbNullString
        Exit Property
    End If

    ContentType = mContentData.ContentType

End Property


Public Property Get ByteLength() As Long

    Dim utf8Bytes() As Byte

    On Error GoTo EH

    ByteLength = 0

    If Not Me.IsValid Then
        Exit Property
    End If

    If mContentData.ByteLen > 0 Then
    
       ByteLength = mContentData.ByteLen
       
       Exit Property
       
    End If
    
    'Debug.Print "ByteLength"
    ' --------------------------------------------------------
    ' File Content
    ' --------------------------------------------------------
    If Me.HasDataType(VTHTTPA_CONTENT_DATATYPE_FILE) Then

         Debug.Print "ByteLength: FAILE"

        If LenB(mContentData.FilePath) = 0 Then
            Exit Property
        End If

        If Len(dir$( _
                mContentData.FilePath, _
                vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then

            Exit Property
        End If

        mContentData.ByteLen = FileLen(mContentData.FilePath)
        
        ByteLength = mContentData.ByteLen
        
        Exit Property

    End If

    ' --------------------------------------------------------
    ' Text Content
    ' --------------------------------------------------------
    If Me.HasDataType(VTHTTPA_CONTENT_DATATYPE_TEXT) Then

         Debug.Print "ByteLength: TEXT"
         
         If Not Me.HasDataType(VTHTTPA_CONTENT_DATATYPE_BYTES) Then
           mContentData.Bytes = VtCharUtils.VtStringToUtf8Bytes(mContentData.text)
           Me.SetDataType (VTHTTPA_CONTENT_DATATYPE_BYTES)
           
         End If

         mContentData.ByteLen = VtCharUtils.VtByteArrayLength(mContentData.Bytes)
          ByteLength = mContentData.ByteLen
        Exit Property

    End If

    ' --------------------------------------------------------
    ' Bytes Content
    ' --------------------------------------------------------
    If Me.HasDataType(VTHTTPA_CONTENT_DATATYPE_BYTES) Then

        mContentData.ByteLen = VtCharUtils.VtByteArrayLength(mContentData.Bytes)
        
         Debug.Print "ByteLength: BYTES.Length=" & CStr(mContentData.ByteLen)
         
         
        
        ByteLength = mContentData.ByteLen
        Exit Property

    End If

    Exit Property

EH:

    ByteLength = 0

End Property

Public Property Get IsFile() As Boolean

    IsFile = ((Me.DataType And VTHTTPA_CONTENT_DATATYPE_FILE) <> 0)

End Property

Public Property Get IsBytes() As Boolean

    IsFile = ((Me.DataType And VTHTTPA_CONTENT_DATATYPE_BYTES) <> 0)

End Property

Public Property Get DataType() As VTHttpaContentDataType

    DataType = mContentData.DataType
    
    
End Property

Public Sub SetDataType(ByVal DType As VTHttpaContentDataType)
   
   'Dim DType As VtHttpaType.VTHttpaContentDataType
   
   DType = Me.DataType Or DType
   
   mContentData.DataType = DType
   

End Sub

Public Sub UnsetDataType(ByVal DType As VTHttpaContentDataType)

   DType = Me.DataType And Not DType
   
   mContentData.DataType = DType


End Sub


Public Function HasDataType(ByVal DType As VTHttpaContentDataType) As Boolean

    HasDataType = ((Me.DataType And DType) <> 0)

End Function
Public Property Get FilePath() As String

    

    If Not Me.IsFile Then
        FilePath = vbNullString
        Exit Property
    End If

    FilePath = mContentData.FilePath

End Property

'Private Function IsByteArrayAllocated(ByRef arr() As Byte) As Boolean
'    On Error GoTo EH
'
'    Dim lb As Long
'    Dim ub As Long
'
'    lb = LBound(arr)
'    ub = UBound(arr)
'
'    IsByteArrayAllocated = (ub >= lb)
'    Exit Function
'
'EH:
'    IsByteArrayAllocated = False
'End Function

' ============================================================
' Data access
' ============================================================

Public Function GetBytes() As Byte()

    EnsureValid "GetBytes"

    If Not Me.IsBytes Then
        Err.Raise vbObjectError + 110, _
                  "CVtHttpaContentPart.GetBytes", _
                  "The content is file-based. Use the FilePath property."
    End If

    If VtCharUtils.VtByteArrayLength(mContentData.Bytes) <= 0 Then
        Err.Raise vbObjectError + 111, _
                  "CVtHttpaContentPart.GetBytes", _
                  "The content byte array is not allocated."
    End If

    GetBytes = mContentData.Bytes

End Function



Public Function GetText() As String


    If Not Me.IsValid Then
        Err.Raise vbObjectError + 10, "VtHttpaContentPart", _
                  "Invalid ContentData: Part is not properly attached."
    End If

    If Me.IsFile Then
        Err.Raise vbObjectError + 11, "VtHttpaContentPart", _
                  "Cannot get text from file-based content. Read file using FilePath property."
    End If
    
        
    If Me.HasDataType(VTHTTPA_CONTENT_DATATYPE_TEXT) Then
    
        GetText = mContentData.text
        
        Exit Function
    
    End If
                
'    If mContentData.Text <> vbNullString Then
'
'        GetText = mContentData.Text
'        Exit Function
'
'    End If

    If VtCharUtils.VtByteArrayLength(mContentData.Bytes) <= 0 Then
       mContentData.text = ""
    Else
       mContentData.text = VtUtf8BytesToString(mContentData.Bytes)
    End If
    
    Call Me.SetDataType(VTHTTPA_CONTENT_DATATYPE_TEXT)
    GetText = mContentData.text
         
        

End Function
Public Function GetFilePath() As String
    ' 1. 유효성 체크
    If Not Me.IsValid Then
        Err.Raise vbObjectError + 10, "VtHttpaContentPart", "Invalid ContentData: Part is not properly attached."
    End If

    ' 2. 파일 모드인지 확인
    If Me.IsFile Then
    
     GetFilePath = mContentData.FilePath
    
        ' 데이터가 메모리(Byte Array)에 있는데 경로를 요구할 경우 빈 문자열 반환 혹은 에러 처리
    Else
        GetFilePath = ""
    End If

     
    ' 3. 데이터 반환 (Variant에 담긴 문자열을 String으로 변환)
'    If VBA.IsEmpty(mContentData.data) Then
'        GetFilePath = ""
'    Else
'        ' Variant에 담긴 파일 경로를 String으로 변환하여 반환
'        GetFilePath = CStr(mContentData.data)
'    End If
End Function
' 비어 있는지 확인
Public Function IsEmpty() As Boolean


    

  
    If Not mContentData.IsValid Then
        IsEmpty = True

    ElseIf mContentData.IsFile Then
        IsEmpty = (Len(mContentData.FilePath) = 0)

    Else
        IsEmpty = (VtCharUtils.VtByteArrayLength(mContentData.Bytes) = 0)

    End If

End Function


' ============================================================
' Initializer Functions (for creating new parts)
' ============================================================

' 파일로 초기화
Public Function InitAsFile( _
    ByVal FilePath As String, _
    Optional ByVal ContentType As String = vbNullString _
) As Boolean

    On Error GoTo ErrorHandler

    Clear

    FilePath = Trim$(FilePath)

    Dim FileLen As Long
    FileLen = Len(FilePath)
    
    If FileLen = 0 Then
        Err.Raise vbObjectError + 120, _
                  "CVtHttpaContentPart.InitAsFile", _
                  "FilePath is empty."
    End If

    If Len(dir$(FilePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then
        Err.Raise vbObjectError + 121, _
                  "CVtHttpaContentPart.InitAsFile", _
                  "File not found: " & FilePath
    End If

    If Len(ContentType) = 0 Then
        ContentType = "application/octet-stream"
    End If

    ' --------------------------------------------------------
    ' Content Data
    ' --------------------------------------------------------
    With mContentData

        .FilePath = FilePath
        .ByteLen = FileLen
        .DataType = VTHTTPA_CONTENT_DATATYPE_FILE
        .ContentType = ContentType
        .IsValid = True

        ' File 자체에서 결정 가능한 Content-Disposition 정보
        .ContentDisposition.DispFilename = dir$(FilePath)
        

    End With

    ' --------------------------------------------------------
    ' Headers
    ' --------------------------------------------------------
    mHeaders.SetHeader "Content-Type", ContentType

    mContentData.DataType = VTHTTPA_CONTENT_DATATYPE_FILE
    
    InitAsFile = True
    Exit Function

ErrorHandler:

    Debug.Print _
        "CVtHttpaContentPart.InitAsFile error: " & _
        Err.Number & ", " & Err.Description

    Clear
    InitAsFile = False

End Function

' 텍스트로 초기화
Public Function InitAsText( _
    ByVal text As String, _
    Optional ByVal ContentType As String = vbNullString _
) As Boolean

    Dim utf8Bytes() As Byte

    On Error GoTo ErrorHandler

    Clear

    If Len(ContentType) = 0 Then
        ContentType = "text/plain; charset=utf-8"
    End If

    utf8Bytes = VtCharUtils.VtStringToUtf8Bytes(text)

    With mContentData

        .text = text
        .Bytes = utf8Bytes
        .ByteLen = VtCharUtils.VtByteArrayLength(utf8Bytes)
        .DataType = VTHTTPA_CONTENT_DATATYPE_TEXT Or VTHTTPA_CONTENT_DATATYPE_BYTES
        .ContentType = ContentType
        .IsValid = True

    End With

    mHeaders.SetHeader "Content-Type", ContentType

    
    'Me.SetDataType (VTHTTPA_CONTENT_DATATYPE_BYTES)
    
    InitAsText = True
    Exit Function

ErrorHandler:

    Debug.Print _
        "CVtHttpaContentPart.InitAsText error: " & _
        Err.Number & ", " & Err.Description

    Clear
    InitAsText = False

End Function
' 바이너리 데이터로 초기화
Public Function InitAsBytes( _
    ByRef Bytes() As Byte, _
    Optional ByVal ContentType As String = vbNullString _
) As Boolean

    Dim ByteLen As Long

    On Error GoTo ErrorHandler

    Clear

    ByteLen = VtCharUtils.VtByteArrayLength(Bytes)

    If ByteLen <= 0 Then
        Err.Raise vbObjectError + 130, _
                  "CVtHttpaContentPart.InitAsBytes", _
                  "The source byte array is not allocated or empty."
    End If

    If Len(ContentType) = 0 Then
        ContentType = "application/octet-stream"
    End If

    With mContentData

        .text = vbNullString
        .Bytes = Bytes
        .ByteLen = VtCharUtils.VtByteArrayLength(Bytes)
        .DataType = VTHTTPA_CONTENT_DATATYPE_BYTES
        .ContentType = ContentType
        .IsValid = True

    End With

    mHeaders.SetHeader "Content-Type", ContentType
    InitAsBytes = True
    Exit Function

ErrorHandler:

    Debug.Print _
        "CVtHttpaContentPart.InitAsBytes error: " & _
        Err.Number & ", " & Err.Description

    Clear
    InitAsBytes = False

End Function


' ============================================================
' Summary
' ============================================================

Public Function ToSummary() As String

    Dim s As String
    Dim typeText As String

    On Error GoTo EH

    ' --------------------------------------------------------
    ' Data Type
    ' --------------------------------------------------------
    typeText = vbNullString

    If Me.HasDataType(VTHTTPA_CONTENT_DATATYPE_FILE) Then
        typeText = typeText & "FILE|"
    End If

    If Me.HasDataType(VTHTTPA_CONTENT_DATATYPE_TEXT) Then
        typeText = typeText & "TEXT|"
    End If

    If Me.HasDataType(VTHTTPA_CONTENT_DATATYPE_BYTES) Then
        typeText = typeText & "BYTES|"
    End If

    If Len(typeText) > 0 Then
        typeText = left$(typeText, Len(typeText) - 1)
    Else
        typeText = "NONE"
    End If

    ' --------------------------------------------------------
    ' Summary
    ' --------------------------------------------------------
    s = "ContentPart {" & _
        "IsValid=" & CStr(Me.IsValid) & _
        ", Type=" & typeText & _
        ", ContentType=""" & Me.ContentType & """" & _
        ", ByteLength=" & CStr(Me.ByteLength)

    If Me.HasDataType(VTHTTPA_CONTENT_DATATYPE_FILE) Then
        s = s & _
            ", FilePath=""" & Me.FilePath & """"
    End If

    If Len(mContentData.ContentDisposition.DispName) > 0 Then
        s = s & _
            ", Name=""" & _
            mContentData.ContentDisposition.DispName & """"
    End If

    If Len(mContentData.ContentDisposition.DispFilename) > 0 Then
        s = s & _
            ", Filename=""" & _
            mContentData.ContentDisposition.DispFilename & """"
    End If

    s = s & "}"

    ToSummary = s
    Exit Function

EH:
    ToSummary = _
        "ContentPart {Error=" & _
        CStr(Err.Number) & _
        ": " & Err.Description & "}"

End Function

