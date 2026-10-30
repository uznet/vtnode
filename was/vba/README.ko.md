[English](README.md) | [한국어](README.ko.md)

# VTNode WAS VBA

2026-10-30에 확인한 로컬 폴더 구조를 기준으로 작성했습니다.

```text
was/vba/
├─ common/
│  ├─ vtnode_vba_rt64.xlsm
│  └─ src/                  공통 VBA 런타임 소스
│     ├─ *.bas, *.cls, *.frm, *.frx
│     ├─ xlsx/              현재 비어 있음; Excel 전용 소스 확장용
│     └─ pptx/              현재 비어 있음; PowerPoint 전용 소스 확장용
└─ xlsx/                    Excel 애플리케이션과 예제
   ├─ vtnode_hello/
   ├─ vtnode_build_myapp/
   ├─ vtnode_was/
   ├─ examples/
   └─ wrapper/
```

현재 공통 소스 루트는 `common/src/`입니다. `xlsx/vtnode_hello/src/` 같은 애플리케이션별 소스 폴더는 별도로 유지됩니다. Office별 하위 폴더는 확장용이며 PowerPoint나 Word의 실행 지원이 완료됐다는 의미는 아닙니다. 런타임 통합 문서 이름은 아직 `vtnode_vba_rt64.xlsm`입니다.

## 시작 안내

- [Excel 설정 및 ZIP 차단 해제](xlsx/README.ko.md)
- [Hello 예제](xlsx/vtnode_hello/README.ko.md)
- [Build My App](xlsx/vtnode_build_myapp/README.ko.md)
- [메인 데모](xlsx/vtnode_was/README.ko.md)
- [공통 런타임](common/README.ko.md)

각 애플리케이션의 통합 문서와 기존 `www/`, `audio/`, 설정 파일의 상대 배치를 유지하세요. 통합 문서 기준의 소스 내보내기는 해당 문서 폴더의 `src/`를 계속 사용합니다. 애플리케이션 폴더를 이동했다는 이유만으로 이 경로를 바꿀 필요는 없습니다.

## 소스 분리 계획

공통 모듈은 `common/src/`에 두고, 실제로 Office별 모듈을 분리할 때 `xlsx/`, `pptx/` 하위 폴더를 사용합니다. 모듈 이동이나 통합 문서 이름 변경 시에는 호스트 의존성, 가져오기·내보내기 코드, 통합 문서 참조를 함께 점검한 뒤 지원할 Office에서 컴파일과 실행을 검증해야 합니다.
