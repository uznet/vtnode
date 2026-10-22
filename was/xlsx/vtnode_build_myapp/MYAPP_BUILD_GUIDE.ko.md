[한글](MYAPP_BUILD_GUIDE.ko.md) | [English](MYAPP_BUILD_GUIDE.en.md)

# VTNode WAS for XLSX — MyApp 직접 만들기

UI만 준비된 Excel 파일에 VBA 소스와 VTNode Runtime을 연결하여 웹 서비스를 직접 만들어 봅니다. 이 문서는 직접 테스트한 제작 과정을 단계별로 정리한 것입니다.

> **이 과정을 통해 만드는 `vtnode_myapp.xlsm`은 완성 예제인 `vtnode_hello.xlsm`과 동일한 기능과 동작을 구현한 결과물입니다.**
>
> 완성된 Hello 예제를 실행하는 대신, UI 파일에서 시작해 같은 애플리케이션을 직접 조립하는 과정입니다. 파일 이름은 다르지만, 별도의 새로운 기능을 추가하는 실습은 아닙니다.

## 준비

### 예제 다운로드와 압축 해제

이 문서는 **GitHub에서 VTNode 예제들을 다운로드한 뒤, 사용자 PC의 폴더에 정상적으로 압축을 해제한 상태**를 전제로 합니다. 압축 파일 안에서 Excel 파일을 바로 열지 말고, 먼저 전체 파일을 압축 해제합니다.

압축을 푼 위치에서 `was/xlsx` 폴더를 찾습니다. 첨부된 GitHub 폴더 구조를 기준으로 예제들은 다음과 같이 같은 위치에 나란히 있습니다.

```text
<압축을 해제한 저장소 폴더>/
└─ was/
   └─ xlsx/
      ├─ vtnode_build_myapp/   ← 이번 실습 작업 폴더
      ├─ vtnode_hello/         ← 완성된 Hello 예제
      ├─ vtnode_vba_rt64/      ← 공통 VBA 런타임
      ├─ vtnode_was/           ← WAS 데모
      └─ readme.md
```

압축을 해제한 폴더 이름과 드라이브는 사용자마다 달라도 됩니다. 예제 사이의 폴더 관계를 유지하고, 이번 실습은 **`was/xlsx/vtnode_build_myapp` 폴더에서 진행**합니다. 이후 문서에서 말하는 작업 폴더는 이 폴더입니다.

### 실습 파일 확인

작업 폴더에서 다음 파일과 폴더를 확인합니다.

| 작업 폴더 안의 항목 | 용도 |
| --- | --- |
| `vtnode_ui_only.xlsx` | UI만 준비된 시작 파일 |
| `src/`의 `.bas`, `.cls` 파일 | 애플리케이션 VBA 코드 |
| `vtnode_httpa_params.ini` | 웹 서비스 설정 |
| `www/` | 브라우저에 제공할 웹 콘텐츠 |
| `audio/` | 효과음 리소스 |

공통 VBA 런타임은 작업 폴더 안이 아니라, 나란히 있는 **`vtnode_vba_rt64` 폴더**에 있습니다.

```text
작업 폴더 기준 상대 경로:
..\vtnode_vba_rt64\vtnode_vba_rt64.xlsm
```

파일이나 폴더가 빠져 있으면 다운로드한 예제와 압축 해제 상태를 먼저 확인합니다. VTNode Runtime 설치와 Excel 매크로 실행 환경도 준비된 상태에서 진행합니다.

## 1. MyApp 파일 만들기

1. 작업 폴더의 `vtnode_ui_only.xlsx`를 Excel에서 엽니다.
2. **파일 → 다른 이름으로 저장**을 선택합니다.
3. 같은 폴더에 `vtnode_myapp.xlsm`으로 저장합니다.
4. 파일 형식은 **Excel 매크로 사용 통합 문서 (*.xlsm)**를 선택합니다.

## 2. VBA 소스 가져오기

1. **개발 도구 → Visual Basic**을 선택하거나 `Alt+F11`을 눌러 VBA 편집기를 엽니다.
2. 프로젝트 탐색기에서 **VBAProject (vtnode_myapp.xlsm)**을 선택합니다.
3. **파일 → 파일 가져오기**에서 작업 폴더 아래 `src/`의 모든 `.cls`, `.bas` 파일을 가져옵니다. 여러 번 가져와 같은 모듈이 중복되지 않도록 확인합니다.

## 3. VTNodeRuntime 참조 추가하기

1. MyApp 프로젝트를 선택한 상태에서 **도구 → 참조**를 엽니다.
2. **찾아보기**에서 작업 폴더의 상위 폴더인 `xlsx`로 이동한 뒤, `vtnode_vba_rt64` 폴더 안의 `vtnode_vba_rt64.xlsm`을 선택합니다. 파일이 보이지 않으면 파일 형식 필터를 변경합니다.
3. `VTNodeRuntime` 참조가 추가되었는지 확인하고 **확인**을 누릅니다.

이 참조를 통해 MyApp의 VBA 코드가 공통 런타임 기능을 사용합니다.

## 4. 통합 문서 열기·닫기 코드 추가하기

프로젝트 탐색기의 **Microsoft Excel 개체 → 현재_통합_문서(ThisWorkbook)**를 더블클릭하고 아래 코드를 추가합니다. 이미 `Workbook_Open` 또는 `Workbook_BeforeClose`가 있다면 같은 이벤트를 중복해서 만들지 않습니다.

- 열 때: 버튼을 연결하고 VTNode를 초기화한 뒤 현재 통합 문서를 런타임 관리 대상으로 등록합니다.
- 닫을 때: HTTP 엔진과 VTNode를 종료하고 런타임 관리 대상에서 현재 통합 문서를 분리합니다.

아래 코드는 테스트 메모의 코드를 그대로 수록했습니다.

```vb
Option Explicit

Private Sub Workbook_Open()

    On Error GoTo OpenError


Debug.Print "Workbook_Open()"
    Application.EnableEvents = False

    HelloUI.UI_InitButtons
    WasMain.WasInit

    Call VtRuntimeManager.VtRuntimeAttach(Me.Name)

OpenExit:
    Application.EnableEvents = True
    Exit Sub

OpenError:
    MsgBox _
        "VTNode 화면 초기화 중 오류가 발생했습니다." & vbCrLf & _
        Err.description, _
        vbExclamation, _
        "VTNode"

    Resume OpenExit
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error GoTo CloseError

    Application.EnableEvents = False

    Debug.Print "Workbook_BeforeClose: BEGIN"

    ' ---------------------------------------------------------
    ' 1. Native HTTP Engine 종료
    '
    ' HTTP Engine이 실행 중인 경우, 새로운 요청 수신 및
    ' 처리 작업을 중지한 후 Native Runtime을 해제합니다.
    ' ---------------------------------------------------------
    If VtHttpaEngine.VtHttpaIsRun Then

        Debug.Print "Workbook_BeforeClose: VtHttpaStop"
        VtHttpaEngine.VtHttpaStop


    End If
        
    Debug.Print "Workbook_BeforeClose: VtNodeUnInit"
    VtNode.VtNodeUnInit

    ' ---------------------------------------------------------
    ' 2. Runtime 관리 대상에서 현재 Workbook 분리
    '
    ' Native Resource 정리가 완료된 후 Runtime에 등록된
    ' 현재 Workbook의 관리 정보를 제거합니다.
    ' ---------------------------------------------------------
    Debug.Print "Workbook_BeforeClose: VtRuntimeDetach"
    VtRuntimeManager.VtRuntimeDetach Me.Name

    Debug.Print "Workbook_BeforeClose: END"

CloseExit:
    Application.EnableEvents = True
    Exit Sub

CloseError:

    Debug.Print _
        "Workbook_BeforeClose Error " & _
        Err.Number & ": " & Err.description

    Resume CloseExit

End Sub
```

## 5. 저장하고 다시 열기

1. `vtnode_myapp.xlsm`을 저장하고 닫습니다.
2. 파일을 다시 엽니다.
3. 신뢰하는 파일의 매크로 실행을 허용합니다.

다시 열면 앞에서 추가한 `Workbook_Open` 코드가 실행되어 버튼과 런타임을 초기화합니다.

## 6. 웹 서비스 실행과 확인

1. Excel 화면에서 **START**를 누릅니다.
2. 웹 서비스가 정상적으로 실행되는지 상태 표시를 확인합니다.
3. **SHOW RUNTIME**을 눌러 런타임 상태를 살펴봅니다.
4. 화면에 표시된 **ACCESS URL**을 브라우저에서 열어 웹 페이지에 접속되는지 확인합니다.

이 단계까지 완료하면 `vtnode_hello.xlsm`과 동일한 기능과 동작의 웹 애플리케이션을 직접 구성한 것입니다.

## 7. 서비스 종료하기

1. 체험이 끝나면 **STOP**을 누릅니다.
2. 서비스가 정지했는지 확인합니다.
3. 통합 문서를 닫습니다.

통합 문서의 종료 코드도 포함되어 있지만, 실습을 마칠 때는 먼저 **STOP → 정지 확인 → 통합 문서 닫기** 순서로 진행합니다.

## 완료 확인

- [ ] `vtnode_myapp.xlsm`으로 저장했다.
- [ ] `src/`의 `.cls`, `.bas` 파일을 가져왔다.
- [ ] `VTNodeRuntime` 참조를 추가했다.
- [ ] `현재_통합_문서`에 열기·닫기 이벤트 코드를 추가했다.
- [ ] 저장 후 다시 열어 START로 서비스를 실행했다.
- [ ] SHOW RUNTIME과 브라우저 접속을 확인했다.
- [ ] STOP으로 서비스를 정지한 뒤 통합 문서를 닫았다.


정리일: 2026-10-22



