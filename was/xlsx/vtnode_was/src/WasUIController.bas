Attribute VB_Name = "WasUIController"
'Attribute VB_Name = "WasUIController"
Option Explicit

Private Const MAIN_SHEET As String = "VTNode WAS"
Private Const RUNTIME_SHEET As String = "VTNode WAS"
Private Const ACCESS_URL_CELL As String = "D14"
Private Const NETWORK_RANGE As String = "D15:H15"


' =========================================================
' 화면의 버튼 영역 위에 투명한 클릭용 Shape를 생성합니다.
' Workbook_Open에서 호출합니다.
' =========================================================
Public Sub UI_InitButtons()

    Dim ws As Worksheet

    On Error GoTo SetupError

    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET)

    Application.ScreenUpdating = False

    CreateMacroOverlay _
        ws:=ws, _
        targetAddress:="B6:D8", _
        shapeName:="vtnode_btn_start", _
        macroName:="UI_StartServer", _
        description:="Start VTNode Server"

    CreateMacroOverlay _
        ws:=ws, _
        targetAddress:="F6:H8", _
        shapeName:="vtnode_btn_stop", _
        macroName:="UI_StopServer", _
        description:="Stop VTNode Server"

    CreateMacroOverlay _
        ws:=ws, _
        targetAddress:="B10:D12", _
        shapeName:="vtnode_btn_runtime", _
        macroName:="UI_ShowRuntime", _
        description:="Show VTNode Runtime"

    CreateActionButton _
        ws:=ws, _
        targetAddress:="J7:M9", _
        shapeName:="vtnode_btn_intro", _
        macroName:="WasUIIntro.ShowIntro", _
        caption:="WHAT IS VTNODE WAS?" & vbCrLf & "INTRO", _
        description:="Open the VTNode WAS introduction"

    CreateActionButton _
        ws:=ws, _
        targetAddress:="J12:M13", _
        shapeName:="vtnode_btn_table_order", _
        macroName:="AppTableOrder.StartTableOrder", _
        caption:="TABLE ORDER" & vbCrLf & "EP01  ·  /table-order", _
        description:="Launch the Table Order application"

    CreateActionButton _
        ws:=ws, _
        targetAddress:="K17:M17", _
        shapeName:="vtnode_btn_clear_log", _
        macroName:="WasUILogger.InitLog", _
        caption:="CLEAR LOG", _
        description:="Clear request and response activity log"

    Application.ScreenUpdating = True
    Exit Sub

SetupError:
    Application.ScreenUpdating = True

    MsgBox _
        "VTNode 버튼을 생성하지 못했습니다." & vbCrLf & _
        Err.description, _
        vbExclamation, _
        "VTNode"
End Sub

Private Sub CreateActionButton( _
    ByVal ws As Worksheet, _
    ByVal targetAddress As String, _
    ByVal shapeName As String, _
    ByVal macroName As String, _
    ByVal caption As String, _
    ByVal description As String)

    Dim targetRange As Range
    Dim buttonShape As Shape
    Dim qualifiedMacroName As String

    Set targetRange = ws.Range(targetAddress)

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Set buttonShape = ws.Shapes.AddShape( _
        Type:=msoShapeRectangle, _
        Left:=targetRange.Left, _
        Top:=targetRange.Top, _
        Width:=targetRange.Width, _
        Height:=targetRange.Height)

    qualifiedMacroName = "'" & _
        Replace(ThisWorkbook.Name, "'", "''") & _
        "'!" & macroName

    With buttonShape
        .Name = shapeName
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(14, 165, 233)
        .Fill.Transparency = 0
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(2, 132, 199)
        .Line.Weight = 1.5
        .OnAction = qualifiedMacroName
        .AlternativeText = description
        .Placement = xlMoveAndSize
        .Locked = True

        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.text = caption
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Font.Name = "Aptos Display"
            .TextRange.Font.Size = 14
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
End Sub

' =========================================================
' 지정된 셀 영역과 정확히 같은 위치에 투명 Shape를 생성합니다.
' =========================================================
Private Sub CreateMacroOverlay( _
    ByVal ws As Worksheet, _
    ByVal targetAddress As String, _
    ByVal shapeName As String, _
    ByVal macroName As String, _
    ByVal description As String)

    Dim targetRange As Range
    Dim overlay As Shape
    Dim qualifiedMacroName As String

    Set targetRange = ws.Range(targetAddress)

    ' 기존 Shape가 있으면 제거한 뒤 다시 생성합니다.
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Set overlay = ws.Shapes.AddShape( _
        Type:=msoShapeRectangle, _
        Left:=targetRange.Left, _
        Top:=targetRange.Top, _
        Width:=targetRange.Width, _
        Height:=targetRange.Height)

    qualifiedMacroName = "'" & _
        Replace(ThisWorkbook.Name, "'", "''") & _
        "'!" & macroName

    With overlay
        .Name = shapeName

        ' 완전한 No Fill 대신 99% 투명 Fill을 사용해야
        ' Shape 전체 영역이 안정적으로 클릭됩니다.
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Transparency = 0.99

        .Line.Visible = msoFalse

        .OnAction = qualifiedMacroName
        .AlternativeText = description

        ' 셀 크기 변경 시 Shape도 함께 이동하고 크기가 변경됩니다.
        .Placement = xlMoveAndSize

        ' Shape 자체에는 표시 문자를 넣지 않습니다.
        .TextFrame2.TextRange.text = vbNullString

        ' 보호된 시트에서 위치가 바뀌지 않게 합니다.
        .Locked = True
    End With
End Sub

' =========================================================
' 화면 버튼에서 호출되는 매크로
' =========================================================

Public Sub UI_StartServer()

    On Error GoTo StartError

    Dim success As Boolean
    
    WasAudio.Play AUDIO_TICK
    success = WasMain.WasStartHttpa()
    
    
    If success Then
    
      Exit Sub
      
    End If
    
    
StartError:
    MsgBox _
        "서버 시작 매크로를 실행하지 못했습니다." & vbCrLf & _
        "연결 대상: VTNode_StartServer" & vbCrLf & vbCrLf & _
        Err.description, _
        vbExclamation, _
        "VTNode"
    
    
End Sub

Public Sub UI_StopServer()

    On Error GoTo StopError
    Dim success As Boolean
    
    WasAudio.Play AUDIO_TICK
    
    success = WasMain.WasStopHttpa()
    
    
    If success Then

    ' 실제 VTNode 정지 매크로 이름으로 변경할 수 있습니다.
    'Application.Run "VTNode_StopServer"
      Exit Sub
    End If
    

StopError:
    MsgBox _
        "서버 정지 매크로를 실행하지 못했습니다." & vbCrLf & _
        "연결 대상: VTNode_StopServer" & vbCrLf & vbCrLf & _
        Err.description, _
        vbExclamation, _
        "VTNode"
        
End Sub

Public Sub UI_ShowRuntime()

    Dim runtimeForm As Object

    On Error GoTo RuntimeError

    ' UserForm 이름이 다르면 아래 이름을 수정하세요.
    'Set runtimeForm = VBA.UserForms.Add("frmVTNodeRuntime")
    'runtimeForm.Show vbModeless
    


    'VtRuntimeUIControl.ShowDemoUI
    
        
    WasAudio.Play AUDIO_TICK

    WasHttpaService.ShowRuntimeStates
    
    Exit Sub

RuntimeError:
    MsgBox _
        "VTNode Runtime 화면을 열지 못했습니다." & vbCrLf & _
        "UserForm 이름을 확인하세요: frmVTNodeRuntime" & vbCrLf & vbCrLf & _
        Err.description, _
        vbExclamation, _
        "VTNode"
End Sub

Public Sub UpdateAccessUrl(ByVal target As Range, ByVal url As String)

    ' 기존 하이퍼링크 제거
    target.Hyperlinks.Delete

    ' 새 주소와 표시 문자열을 함께 등록
    target.Worksheet.Hyperlinks.Add _
        Anchor:=target, _
        Address:=url, _
        TextToDisplay:=url

End Sub

Public Sub UI_SetAccessUrl(ByVal url As String)

   Call UpdateAccessUrl(ThisWorkbook.Worksheets(MAIN_SHEET).Range(ACCESS_URL_CELL), url)
'    ThisWorkbook.Worksheets(MAIN_SHEET) _
'        .Range(ACCESS_URL_CELL).Value2 = Trim$(url)
End Sub

Public Sub UI_SetNetwork( _
    ByRef httpInfo As VtRuntimeHttpInfo, _
    ByRef httpsInfo As VtRuntimeHttpsInfo)

    Dim httpText As String
    Dim httpsText As String

    If httpInfo.enabled Then
        httpText = "HTTP " & CStr(httpInfo.Port)
    Else
        httpText = "HTTP OFF"
    End If

    If httpsInfo.enabled Then
        httpsText = "HTTPS " & CStr(httpsInfo.Port)
    Else
        httpsText = "HTTPS OFF"
    End If

    ThisWorkbook.Worksheets(MAIN_SHEET) _
        .Range(NETWORK_RANGE).Cells(1, 1).Value2 = _
        httpText & "  ·  " & httpsText
End Sub

Public Sub UI_OpenInBrowser()

    Dim serverURL As String

    On Error GoTo BrowserError

    serverURL = Trim$(CStr( _
        ThisWorkbook.Worksheets(RUNTIME_SHEET).Range(ACCESS_URL_CELL).Value2))

    If Len(serverURL) = 0 Then
        MsgBox _
            "HTTP URL이 설정되지 않았습니다.", _
            vbInformation, _
            "VTNode"
        Exit Sub
    End If

    If LCase$(Left$(serverURL, 7)) <> "http://" And _
       LCase$(Left$(serverURL, 8)) <> "https://" Then

        serverURL = "http://" & serverURL
    End If

    ThisWorkbook.FollowHyperlink _
        Address:=serverURL, _
        NewWindow:=True

    Exit Sub

BrowserError:
    MsgBox _
        "브라우저를 열지 못했습니다." & vbCrLf & _
        serverURL & "->" & Err.description, _
        vbExclamation, _
        "VTNode"
End Sub

' 버튼 위치를 수동으로 다시 맞출 때 실행합니다.
Public Sub RefreshVTNodeButtons()
    UI_InitButtons
End Sub

' 생성된 투명 Shape를 모두 제거합니다.
Public Sub RemoveVTNodeButtons()

    Dim ws As Worksheet
    Dim buttonNames As Variant
    Dim item As Variant

    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET)

    buttonNames = Array( _
        "vtnode_btn_start", _
        "vtnode_btn_stop", _
        "vtnode_btn_runtime", _
        "vtnode_btn_intro", _
        "vtnode_btn_table_order", _
        "vtnode_btn_clear_log")

    For Each item In buttonNames
        On Error Resume Next
        ws.Shapes(CStr(item)).Delete
        On Error GoTo 0
    Next item
End Sub




Public Sub CloseVTNodeRuntimeWorkbook()

    Const RUNTIME_BOOK_NAME As String = "vtnode_vba_rt64.xlsm"

    Dim runtimeBook As Workbook

    On Error Resume Next
    Set runtimeBook = Application.Workbooks(RUNTIME_BOOK_NAME)
    On Error GoTo CloseError

    If runtimeBook Is Nothing Then
        Exit Sub
    End If

    ' Runtime 내부 종료 처리가 별도로 필요하면 여기서 먼저 호출
    ' Application.Run "'" & RUNTIME_BOOK_NAME & "'!VtRtShutdown"

    runtimeBook.Close SaveChanges:=False
    Exit Sub

CloseError:
    Debug.Print _
        "CloseVTNodeRuntimeWorkbook Error " & _
        Err.Number & ": " & Err.description
End Sub



Public Sub RepairRuntimeConfigLinks()

    Dim ws As Worksheet
    Dim links As Variant
    Dim link As Variant

    Set ws = ThisWorkbook.Worksheets(MAIN_SHEET)

    ' Keep configuration formulas away from the activity log rows.
    ' Runtime and network values are written through the UI_Set* procedures.
    ws.Range(ACCESS_URL_CELL).Formula = "='Runtime Config'!B3"

    ' 남아 있는 외부 Excel 링크 제거
    links = ThisWorkbook.LinkSources(Type:=xlExcelLinks)

    If Not IsEmpty(links) Then
        For Each link In links
            If InStr(1, CStr(link), _
                     "VTNode_Get_Started.xlsm", _
                     vbTextCompare) > 0 Then

                ThisWorkbook.BreakLink _
                    Name:=CStr(link), _
                    Type:=xlLinkTypeExcelLinks
            End If
        Next link
    End If

    ThisWorkbook.Save

    MsgBox _
        "Runtime Config 외부 링크를 내부 시트 참조로 수정했습니다.", _
        vbInformation, _
        "VTNode"

End Sub



