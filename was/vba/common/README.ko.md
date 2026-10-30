# VTNode 공통 VBA 소스

공통 VBA는 각 응용 파일에 직접 포함합니다. 네이티브 VTNode Runtime은 HTTP/WebSocket 엔진을 제공하며 별도로 설치합니다.

```text
was/vba/
├─ common/
│  ├─ src/    # Office shared VBA
│  ├─ xlsx/   # Excel only
│  └─ pptx/   # PowerPoint only
└─ xlsx/      # Excel apps and examples
```

## 공통 VBA 직접 Import 및 빌드

확정된 구조는 `was/vba/common/src`(Office 공통 VBA), `was/vba/common/xlsx`(Excel 전용), `was/vba/common/pptx`(PowerPoint 전용), `was/vba/xlsx/`(Excel 앱/예제)입니다. 앱 폴더의 `src/`는 앱 전용 코드이며 공통 소스 전체를 대신하지 않습니다.

1. 대상 파일을 백업합니다. Excel UI 원본 `.xlsx`는 `.xlsm`으로 저장하고, PowerPoint는 `.pptm`, Word는 `.docm` 등 VBA를 보존하는 형식을 사용합니다.
2. 대상 파일에서 VBA 편집기를 열고 **파일 > 파일 가져오기**로 [common/src](src/)의 필요한 `.bas`·`.cls`와 의존 모듈을 해당 VBA 프로젝트에 직접 Import합니다.
3. Excel에는 [common/xlsx](xlsx/), PowerPoint에는 [common/pptx](pptx/)의 해당 호스트 전용 소스를 추가합니다. Excel 전용 코드를 다른 Office 프로젝트에 Import하지 않습니다. Word는 Office 공통 소스를 사용하고 필요한 호스트 연결 코드를 별도로 구현·검증합니다.
4. 앱의 `src/`에서 앱 전용 소스를 추가합니다. 같은 이름의 모듈은 백업 후 교체하여 중복을 피합니다. UserForm은 `.frm`과 `.frx`를 같은 폴더에 두고 `.frm`을 Import합니다. `ThisWorkbook`·시트·문서 이벤트 코드는 대상 문서 모듈에 반영합니다.
5. 외부 런타임 Workbook을 열거나 VBA 프로젝트 참조로 연결하지 않습니다. 기존의 외부 런타임 Workbook 참조는 **도구 > 참조**에서 제거하고 필요한 라이브러리 참조를 확인합니다. UI·매크로 연결도 대상 파일 내부 프로시저를 사용하도록 확인합니다.
6. **디버그 > VBAProject 컴파일** 후 저장합니다. Office와 일치하는 비트 수의 네이티브 VTNode Runtime을 준비하고, `www`·`audio`·설정 파일의 상대 배치를 유지한 상태에서 시작·종료, HTTP/WebSocket, Runtime UI를 확인합니다. Office 공통 소스라는 분류만으로 모든 Office에서의 실행이 검증된 것은 아닙니다.

[VBA project](https://github.com/uznet/vtnode/tree/main/was/vba) | [Excel apps](https://github.com/uznet/vtnode/tree/main/was/vba/xlsx)
