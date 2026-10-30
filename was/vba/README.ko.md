[English](README.md) | [한국어](README.ko.md)

# VTNode WAS VBA

## 공통 VBA 직접 Import 및 빌드

확정된 구조는 `was/vba/common/src`(Office 공통 VBA), `was/vba/common/xlsx`(Excel 전용), `was/vba/common/pptx`(PowerPoint 전용), `was/vba/xlsx/`(Excel 앱/예제)입니다. 앱 폴더의 `src/`는 앱 전용 코드이며 공통 소스 전체를 대신하지 않습니다.

1. 대상 파일을 백업합니다. Excel UI 원본 `.xlsx`는 `.xlsm`으로 저장하고, PowerPoint는 `.pptm`, Word는 `.docm` 등 VBA를 보존하는 형식을 사용합니다.
2. 대상 파일에서 VBA 편집기를 열고 **파일 > 파일 가져오기**로 [common/src](common/src/)의 필요한 `.bas`·`.cls`와 의존 모듈을 해당 VBA 프로젝트에 직접 Import합니다.
3. Excel에는 [common/xlsx](common/xlsx/), PowerPoint에는 [common/pptx](common/pptx/)의 해당 호스트 전용 소스를 추가합니다. Excel 전용 코드를 다른 Office 프로젝트에 Import하지 않습니다. Word는 Office 공통 소스를 사용하고 필요한 호스트 연결 코드를 별도로 구현·검증합니다.
4. 앱의 `src/`에서 앱 전용 소스를 추가합니다. 같은 이름의 모듈은 백업 후 교체하여 중복을 피합니다. UserForm은 `.frm`과 `.frx`를 같은 폴더에 두고 `.frm`을 Import합니다. `ThisWorkbook`·시트·문서 이벤트 코드는 대상 문서 모듈에 반영합니다.
5. 외부 런타임 Workbook을 열거나 VBA 프로젝트 참조로 연결하지 않습니다. 기존의 외부 런타임 Workbook 참조는 **도구 > 참조**에서 제거하고 필요한 라이브러리 참조를 확인합니다. UI·매크로 연결도 대상 파일 내부 프로시저를 사용하도록 확인합니다.
6. **디버그 > VBAProject 컴파일** 후 저장합니다. Office와 일치하는 비트 수의 네이티브 VTNode Runtime을 준비하고, `www`·`audio`·설정 파일의 상대 배치를 유지한 상태에서 시작·종료, HTTP/WebSocket, Runtime UI를 확인합니다. Office 공통 소스라는 분류만으로 모든 Office에서의 실행이 검증된 것은 아닙니다.


VBA 프로젝트의 기준 경로는 [was/vba](https://github.com/uznet/vtnode/tree/main/was/vba)입니다. GitHub 저장소 루트는 [uznet/vtnode](https://github.com/uznet/vtnode)입니다.

## 폴더 구조

```text
was/vba/
├─ common/
│  ├─ src/    # Office shared VBA
│  ├─ xlsx/   # Excel only
│  └─ pptx/   # PowerPoint only
├─ documents/
└─ xlsx/      # Excel apps and examples
```

Office 공통 모듈과 클래스는 common/src, Excel 전용 소스는 common/xlsx, PowerPoint 전용 소스는 common/pptx에 둡니다.

documents에는 시작 안내, Build Manual, 라이선스 및 설정 문서를 둡니다. xlsx/examples와 xlsx/note는 현재 작업 대상에서 제외합니다.

## 시작 안내

- [Excel 설정 및 ZIP 차단 해제](xlsx/README.ko.md)
- [Quick Start](documents/QUICK_START.ko.md)
- [Build Manual](documents/MYAPP_BUILD_GUIDE.ko.md)
- [Hello 예제](xlsx/vtnode_hello/README.ko.md)
- [Build My App](xlsx/vtnode_build_myapp/README.ko.md)
- [메인 데모](xlsx/vtnode_was/README.ko.md)
- [공통 런타임](common/README.ko.md)
- [문서 폴더](documents/)

각 애플리케이션의 통합 문서와 www, audio, 설정 파일의 상대 배치를 유지하세요. 애플리케이션별 src는 해당 통합 문서 전용 소스를 보관합니다. 공통 소스의 기준은 common/src와 common/xlsx입니다. common/xlsx의 UserForm은 .frm과 .frx를 함께 관리합니다.
