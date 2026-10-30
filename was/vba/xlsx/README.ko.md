[한글](./README.ko.md) | [English](./README.md)

[VBA 프로젝트](https://github.com/uznet/vtnode/tree/main/was/vba) | [현재 폴더](https://github.com/uznet/vtnode/tree/main/was/vba/xlsx)

# ⚠️ 필독 — ZIP 압축을 풀기 전에 반드시 차단 해제하세요!

## 공통 VBA 직접 Import 및 빌드

확정된 구조는 `was/vba/common/src`(Office 공통 VBA), `was/vba/common/xlsx`(Excel 전용), `was/vba/common/pptx`(PowerPoint 전용), `was/vba/xlsx/`(Excel 앱/예제)입니다. 앱 폴더의 `src/`는 앱 전용 코드이며 공통 소스 전체를 대신하지 않습니다.

1. 대상 파일을 백업합니다. Excel UI 원본 `.xlsx`는 `.xlsm`으로 저장하고, PowerPoint는 `.pptm`, Word는 `.docm` 등 VBA를 보존하는 형식을 사용합니다.
2. 대상 파일에서 VBA 편집기를 열고 **파일 > 파일 가져오기**로 [common/src](../common/src/)의 필요한 `.bas`·`.cls`와 의존 모듈을 해당 VBA 프로젝트에 직접 Import합니다.
3. Excel에는 [common/xlsx](../common/xlsx/), PowerPoint에는 [common/pptx](../common/pptx/)의 해당 호스트 전용 소스를 추가합니다. Excel 전용 코드를 다른 Office 프로젝트에 Import하지 않습니다. Word는 Office 공통 소스를 사용하고 필요한 호스트 연결 코드를 별도로 구현·검증합니다.
4. 앱의 `src/`에서 앱 전용 소스를 추가합니다. 같은 이름의 모듈은 백업 후 교체하여 중복을 피합니다. UserForm은 `.frm`과 `.frx`를 같은 폴더에 두고 `.frm`을 Import합니다. `ThisWorkbook`·시트·문서 이벤트 코드는 대상 문서 모듈에 반영합니다.
5. 외부 런타임 Workbook을 열거나 VBA 프로젝트 참조로 연결하지 않습니다. 기존의 외부 런타임 Workbook 참조는 **도구 > 참조**에서 제거하고 필요한 라이브러리 참조를 확인합니다. UI·매크로 연결도 대상 파일 내부 프로시저를 사용하도록 확인합니다.
6. **디버그 > VBAProject 컴파일** 후 저장합니다. Office와 일치하는 비트 수의 네이티브 VTNode Runtime을 준비하고, `www`·`audio`·설정 파일의 상대 배치를 유지한 상태에서 시작·종료, HTTP/WebSocket, Runtime UI를 확인합니다. Office 공통 소스라는 분류만으로 모든 Office에서의 실행이 검증된 것은 아닙니다.


> [!WARNING]
> **GitHub ZIP 다운로드 → ZIP 우클릭 → 속성 → 일반 → 차단 해제 → 적용 → 확인 → 압축 풀기**
>
> **Excel VBA 매크로의 인터넷 출처 차단을 피하려면 반드시 ZIP 압축을 풀기 전에 차단 해제하세요.** ZIP의 인터넷 보안 표시가 압축을 푼 `.xlsm` 파일에 전달될 수 있습니다. 압축을 푼 뒤 ZIP만 차단 해제해도 이미 풀린 파일의 차단은 해제되지 않습니다.
>
> **보안 주의: 저장소와 다운로드 파일을 확인하고, 신뢰할 수 있는 출처의 파일에만 적용하세요. VBA 매크로는 악성 코드를 실행할 수 있습니다. GitHub에서 받았다는 사실만으로 안전이 보장되지는 않습니다.**

---

# VTNode WAS for Excel

VTNode WAS는 HTTP와 WebSocket을 통해 Excel/VBA 애플리케이션을 웹 애플리케이션과 연결합니다. 프로젝트 설명과 예제는 [기존 XLSX/VBA 안내](./vtnode_hello/README.ko.md)를 참고하세요.

## Windows: GitHub ZIP 압축을 풀기 전

1. GitHub에서 저장소 ZIP을 다운로드한 뒤 파일 탐색기에서 찾습니다.
2. 저장소와 다운로드 파일이 신뢰할 수 있는 출처인지 확인합니다.
3. **ZIP 파일 자체**를 마우스 오른쪽 버튼으로 클릭하고 **속성(Properties)**을 선택합니다.
4. **일반(General)** 탭 아래쪽 보안 영역에서 **차단 해제(Unblock)**를 체크합니다.
5. **적용(Apply)** → **확인(OK)**을 누릅니다.
6. ZIP을 우클릭하여 **압축 풀기(Extract All)**를 실행한 뒤, 압축을 푼 폴더에서 필요한 통합 문서를 엽니다.

## 이미 압축을 풀었다면: XLSM 파일 자체 차단 해제

1. Excel에서 해당 통합 문서를 닫습니다.
2. 파일 탐색기에서 실제로 사용할 **`.xlsm` 파일**을 우클릭하고 **속성**을 선택합니다.
3. **일반 → 보안 → 차단 해제**를 체크한 뒤 **적용 → 확인**을 누릅니다.
4. 통합 문서를 다시 엽니다. Excel에서 매크로 실행 허용을 요청하면, 문서와 코드를 신뢰할 수 있는지 확인한 경우에만 허용합니다.
5. 필요한 XLSM 파일마다 개별적으로 진행합니다. 또는 원본 ZIP을 차단 해제한 후 새 폴더에 다시 압축을 풉니다. 기존 파일에 작업한 내용이 있다면 보존하세요.

## 여전히 매크로가 차단된다면

인터넷 보안 표시가 없는 파일에는 차단 해제 항목이 보이지 않을 수 있습니다. 차단 해제는 해당 표시를 제거하는 절차이며, Excel 설정이나 조직 정책에 따른 실행 제한은 남아 있을 수 있습니다. 조직의 규정을 따르고 필요하면 관리자에게 문의하세요. 문제 해결을 위해 모든 매크로를 전역 허용하거나 보안 기능을 끄지 마세요.

[Microsoft 공식 안내: 인터넷에서 가져온 Office 파일의 매크로 차단](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked)
