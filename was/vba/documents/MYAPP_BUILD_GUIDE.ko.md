[한글](MYAPP_BUILD_GUIDE.ko.md) | [English](MYAPP_BUILD_GUIDE.en.md)

# VTNode WAS for XLSX — MyApp 직접 만들기

UI만 준비된 Excel 파일에 VBA 소스와 VTNode Runtime을 연결하여 웹 서비스를 직접 만들어 봅니다. 이 문서는 직접 테스트한 제작 과정을 단계별로 정리한 것입니다.

> **이 과정을 통해 만드는 `vtnode_myapp.xlsm`은 완성 예제인 `vtnode_hello.xlsm`과 동일한 기능과 동작을 구현한 결과물입니다.**
>
> 완성된 Hello 예제를 실행하는 대신, UI 파일에서 시작해 같은 애플리케이션을 직접 조립하는 과정입니다. 파일 이름은 다르지만, 별도의 새로운 기능을 추가하는 실습은 아닙니다.

## 준비

### 예제 다운로드와 압축 해제

이 문서는 **GitHub에서 VTNode 예제들을 다운로드한 뒤, 사용자 PC의 폴더에 정상적으로 압축을 해제한 상태**를 전제로 합니다. 압축 파일 안에서 Excel 파일을 바로 열지 말고, 먼저 전체 파일을 압축 해제합니다.

압축을 푼 위치에서 `was/vba/xlsx` 폴더를 찾습니다. 첨부된 GitHub 폴더 구조를 기준으로 예제들은 다음과 같이 같은 위치에 나란히 있습니다.

```text
<압축을 해제한 저장소 폴더>/
└─ was/
   └─ vba/
      ├─ common/
      │  ├─ src/    ← Office 공통 VBA
      │  ├─ xlsx/   ← Excel 전용
      │  └─ pptx/   ← PowerPoint 전용
      ├─ documents/
      │  ├─ MYAPP_BUILD_GUIDE.ko.md
      │  └─ MYAPP_BUILD_GUIDE.en.md
      └─ xlsx/
         ├─ vtnode_build_myapp/   ← 이번 실습 작업 폴더
         ├─ vtnode_hello/         ← 완성된 Hello 예제
         └─ vtnode_was/           ← WAS 데모
```

압축을 해제한 폴더 이름과 드라이브는 사용자마다 달라도 됩니다. 예제 사이의 폴더 관계를 유지하고, 이번 실습은 **`was/vba/xlsx/vtnode_build_myapp` 폴더에서 진행**합니다. 이후 문서에서 말하는 작업 폴더는 이 폴더입니다.

### 실습 파일 확인

작업 폴더에서 다음 파일과 폴더를 확인합니다.

| 작업 폴더 안의 항목 | 용도 |
| --- | --- |
| `vtnode_ui_only.xlsx` | UI만 준비된 시작 파일 |
| `src/`의 `.bas`, `.cls` 파일 | 애플리케이션 VBA 코드 |
| `vtnode_httpa_params.ini` | 웹 서비스 설정 |
| `www/` | 브라우저에 제공할 웹 콘텐츠 |
| `audio/` | 효과음 리소스 |

공통 소스는 **`was/vba/common/src`**(Office 공통 VBA), Excel 전용 소스는 **`was/vba/common/xlsx`**에 있습니다. 두 폴더의 소스를 MyApp 프로젝트에 직접 Import합니다. **`was/vba/common/pptx`**는 PowerPoint 전용이므로 이번 Excel 실습에서는 가져오지 않습니다. [공통 소스 안내](../common/README.ko.md)를 참고하세요.

```text
작업 폴더 기준 상대 경로:
..\..\common\src\
..\..\common\xlsx\
```

파일이나 폴더가 빠져 있으면 다운로드한 예제와 압축 해제 상태를 먼저 확인합니다. VTNode Runtime 설치와 Excel 매크로 실행 환경도 준비된 상태에서 진행합니다.

## 1. MyApp 파일 만들기

1. 작업 폴더의 `vtnode_ui_only.xlsx`를 Excel에서 엽니다.
2. **파일 → 다른 이름으로 저장**을 선택합니다.
3. 같은 폴더에 `vtnode_myapp.xlsm`으로 저장합니다.
4. 파일 형식은 **Excel 매크로 사용 통합 문서 (*.xlsm)**를 선택합니다.

## 2. 공통 VBA 런타임 소스 먼저 가져오기

1. **개발 도구 → Visual Basic**을 선택하거나 `Alt+F11`을 눌러 VBA 편집기를 엽니다.
2. 프로젝트 탐색기에서 **VBAProject (vtnode_myapp.xlsm)**을 선택합니다.
3. **파일 → 파일 가져오기**에서 작업 폴더에서 두 단계 위인 `was/vba`로 이동한 뒤 **`common/src`**를 엽니다. 공통 `.bas`, `.cls`와 필요한 의존 모듈을 가져옵니다. `ExportProjectSources.bas`는 아래의 선택 사항을 참고합니다.
4. 이어서 **`common/xlsx`**의 Excel 전용 `.bas`, `.cls`, `.frm` 소스를 가져옵니다. 같은 이름의 기존 모듈은 백업 후 교체하여 중복을 피합니다.
5. `VtRuntimeConsole.frm`을 가져올 때는 **`VtRuntimeConsole.frx`가 같은 폴더에 있는 상태**를 유지합니다. `.frx`는 폼 리소스 파일이므로 별도로 가져오지 않습니다.

### Microsoft Forms 2.0 Object Library 참조 추가

1. **VBAProject (vtnode_myapp.xlsm)**을 선택한 상태에서 **도구 → 참조**를 엽니다.
2. 목록에서 **Microsoft Forms 2.0 Object Library**를 찾아 체크합니다. 이미 체크되어 있다면 그대로 유지합니다.
3. **확인**을 눌러 적용합니다.

런타임의 UserForm을 사용하려면 이 참조가 필요합니다. `.frm` 파일을 가져온 뒤에도 참조가 선택되어 있는지 확인합니다.

공통 VBA를 MyApp 안에 직접 포함합니다. 외부 런타임 Workbook을 열거나 **도구 → 참조**에 프로젝트 참조로 추가하지 않습니다. 기존 외부 런타임 Workbook 참조가 있으면 제거합니다.

### VBA 소스 내보내기 모듈용 참조 추가

`ExportProjectSources.bas`의 `Dim comp As VBComponent` 및 `vbext_` 상수를 컴파일하려면 다음 참조도 필요합니다.

1. **도구 → 참조**를 엽니다.
2. **Microsoft Visual Basic for Applications Extensibility 5.3**을 체크하고 **확인**을 누릅니다.
3. **디버그 → VBAProject 컴파일**을 실행하여 확인합니다.

`VBComponent`에서 “사용자 정의 형식이 정의되지 않았습니다” 오류가 발생하면 이 참조를 확인합니다. 앞에서 추가한 `Microsoft Forms 2.0 Object Library`와는 별도의 참조입니다.

`ExportProjectSources`는 VBA 소스를 파일로 내보내는 도구입니다. 내보내기 기능이 필요 없다면 이 모듈을 가져오지 않아도 됩니다. 내보내기 매크로를 실제로 실행하려면 Excel 보안 센터의 **VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음** 설정도 필요하지만, 이 설정만 변경해서는 위 컴파일 오류가 해결되지 않습니다.

## 3. 애플리케이션 VBA 소스 가져오기

1. 같은 **VBAProject (vtnode_myapp.xlsm)**을 선택합니다.
2. **파일 → 파일 가져오기**에서 작업 폴더인 `vtnode_build_myapp` 아래의 **`src/`** 폴더를 엽니다.
3. 이 폴더의 모든 `.cls`, `.bas` 파일을 가져옵니다. 같은 모듈을 중복해서 가져오지 않도록 확인합니다.

가져오는 순서는 **`common/src`의 Office 공통 소스 → `common/xlsx`의 Excel 전용 소스 → `vtnode_build_myapp/src`의 앱 전용 소스**입니다. 모두 하나의 MyApp 프로젝트에 추가하고 같은 이름의 모듈을 중복 Import하지 않습니다.

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
- [ ] `common/src`의 Office 공통 소스와 `common/xlsx`의 Excel 전용 소스·UserForm을 먼저 가져왔다.
- [ ] `Microsoft Forms 2.0 Object Library` 참조를 추가했다.
- [ ] `ExportProjectSources`를 포함했다면 `Microsoft Visual Basic for Applications Extensibility 5.3` 참조를 추가했다.
- [ ] `vtnode_build_myapp/src`의 `.cls`, `.bas` 파일을 같은 프로젝트에 가져왔다.
- [ ] `현재_통합_문서`에 열기·닫기 이벤트 코드를 추가했다.
- [ ] 저장 후 다시 열어 START로 서비스를 실행했다.
- [ ] SHOW RUNTIME과 브라우저 접속을 확인했다.
- [ ] STOP으로 서비스를 정지한 뒤 통합 문서를 닫았다.


정리일: 2026-10-22






