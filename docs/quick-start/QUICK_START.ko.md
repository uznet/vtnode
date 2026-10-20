# VTNode WAS for XLSX 빠른 시작

[English](./QUICK_START.md) | **한국어**

VTNode Runtime을 설치하고 Excel/VBA가 브라우저 요청에 응답하는 첫 번째 예제를 실행합니다.

> 라이선스 없이도 기간과 기능 제한 없이 1채널로 실행할 수 있습니다. 먼저 기본 실행을 확인한 뒤, 더 많은 동시 연결이 필요하면 무료 4채널 라이선스를 발급받으세요.

## 전체 흐름

```text
01 VTNode Runtime 설치
        ↓
02 VTCheck64로 설치 확인
        ↓
03 GitHub에서 Excel 예제 다운로드
        ↓
04 라이선스 없이 1채널 실행
        ↓
05 브라우저 응답 확인
        ↓
선택: 무료 4채널 라이선스 발급
```

## 1. VTNode Runtime 설치

[VTNode 다운로드 페이지](https://vtnode.io/downloads/)에서 최신 Windows x64 설치 프로그램을 내려받습니다.

```text
VTNode-Setup-<version>-x64.exe
```

다운로드 페이지에서 이메일 주소와 필수 동의 항목을 입력하면 제한된 시간 동안 사용할 수 있는 다운로드 링크가 이메일로 전송됩니다.

설치 프로그램을 실행하면 다음 작업이 자동으로 처리됩니다.

- VTNode Runtime 설치
- `VTNODE_HOME` 환경 변수 설정
- Runtime 실행에 필요한 폴더와 파일 구성

`VTNODE_HOME`을 사용자가 직접 설정할 필요는 없습니다.

## 2. 설치 상태 확인

설치 프로그램의 마지막 화면에서 **VTNode 설치 상태 확인** 항목은 기본적으로 선택되어 있습니다.

```text
☑ VTNode 설치 상태 확인
```

설치를 마치면 `VTCheck64.exe`가 자동으로 실행되어 다음 항목을 검사합니다.

- `VTNODE_HOME`에 설정된 Runtime 경로
- 설치된 VTNode 패키지 버전
- Runtime 설치 경로와 필수 파일
- 최종 설치 상태

정상 출력 예시:

```text
VTNode Runtime Installation Check

VTNODE_HOME     : C:\Program Files\VTNode
Package Version : <installed-version>
Required Files  : OK
Status          : OK
```

설치 완료 화면에서 확인 항목을 해제했거나 나중에 다시 검사하려면 설치 폴더의 `VTCheck64.exe`를 직접 실행합니다.

`Status : ERROR`가 표시되면 함께 출력된 원인을 확인한 후 설치 상태를 다시 점검하세요.

## 3. Excel 예제 다운로드

GitHub에서 [VTNode WAS for XLSX 예제](https://github.com/uznet/vtnode/tree/main/was/xlsx)를 내려받습니다.

예제 XLSM 파일을 열기 전에 다음 사항을 확인합니다.

1. 압축 파일로 받았다면 먼저 압축을 해제합니다.
2. Windows가 파일을 차단한 경우 파일의 **속성**에서 차단을 해제합니다.
3. Excel에서 XLSM 파일을 엽니다.
4. 신뢰할 수 있는 파일인지 확인한 후 매크로를 허용합니다.

## 4. 라이선스 없이 실행

예제 XLSM에서 VTNode WAS를 시작합니다. 처음 실행할 때 별도의 라이선스 발급이나 코드 입력은 필요하지 않습니다.

라이선스 없는 기본 실행 조건:

- 동시 연결: 1채널
- 사용 기간 제한: 없음
- 기능 제한: 없음

처음 설치한 사용자는 이 방식으로 먼저 Runtime과 Excel/VBA의 연결을 확인하는 것을 권장합니다.

## 5. 브라우저 응답 확인

Excel 화면에 표시된 접속 주소를 PC 또는 모바일 브라우저에서 엽니다.

```text
Browser
   → VTNode Runtime
   → Excel/VBA
   → HTTP Response
   → Browser
```

브라우저 요청이 Excel/VBA로 전달되고 VBA가 생성한 결과가 브라우저에 표시되면 Quick Start가 완료된 것입니다.

## 선택 사항: 무료 4채널 라이선스

기본 1채널보다 많은 동시 연결이 필요하면 [VTNode 무료 라이선스 페이지](https://vtnode.io/license/free-trial)에서 4채널 라이선스를 발급받을 수 있습니다.

무료 라이선스는 기본 Quick Start를 완료한 뒤 적용해도 됩니다. 발급 조건과 사용 정책은 라이선스 페이지의 최신 안내를 따릅니다.

## AI로 예제 코드 확장하기

GitHub에서 내려받은 예제의 실제 VBA 모듈을 AI에 첨부한 후 다음과 같이 요청해 보세요.

> 이 VTNode 예제의 기존 API와 코드 구조를 유지하면서, `GET /status` 요청이 들어오면 `Sheet1!B2`의 값을 JSON으로 반환하도록 VBA 코드를 작성해 줘. 코드를 추가할 모듈과 위치도 설명해 줘.

실제 예제 모듈을 함께 제공하면 AI가 존재하지 않는 VTNode API를 임의로 만드는 오류를 줄일 수 있습니다. 생성된 코드는 실행하기 전에 사용 중인 예제의 API와 구조에 맞는지 확인하세요.

## 문제가 발생하면

### Excel에서 Runtime을 찾지 못하는 경우

- 설치 후 이미 실행 중이던 Excel을 완전히 종료하고 다시 실행합니다.
- `VTCheck64.exe`를 실행하여 `VTNODE_HOME`과 필수 파일 상태를 확인합니다.

### 매크로가 실행되지 않는 경우

- XLSM 파일이 차단되어 있지 않은지 확인합니다.
- Excel의 보안 경고에서 신뢰할 수 있는 파일에 대해서만 매크로를 허용합니다.

### 브라우저 응답이 없는 경우

- `VTCheck64.exe`의 최종 상태가 `OK`인지 확인합니다.
- Excel에서 VTNode WAS가 시작된 상태인지 확인합니다.
- Excel 화면에 표시된 주소를 정확히 입력했는지 확인합니다.

## 관련 링크

- [VTNode 홈페이지](https://vtnode.io/)
- [다운로드](https://vtnode.io/downloads/)
- [VTNode WAS for XLSX](https://github.com/uznet/vtnode/tree/main/was/xlsx)
- [무료 4채널 라이선스](https://vtnode.io/license/free-trial)
- [Documents](https://vtnode.io/docs)

