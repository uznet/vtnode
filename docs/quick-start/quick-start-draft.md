# VTNode WAS for XLSX Quick Start

> 설치부터 첫 번째 브라우저 응답까지 빠르게 확인합니다.

VTNode WAS for XLSX는 라이선스 없이도 기간과 기능 제한 없이 1채널로 실행할 수 있습니다. 먼저 기본 실행으로 동작을 확인한 뒤, 더 많은 동시 연결이 필요하면 무료 4채널 라이선스를 발급받으세요.

---

## 1. VTNode Runtime 설치 및 확인

### 1-1. Runtime 다운로드

VTNode 홈페이지에서 Windows 64비트 설치 프로그램을 내려받습니다.

- 설치 파일: `VTNode-Setup-0.1.0-x64.exe`
- 다운로드: [VTNode Downloads](https://vtnode.io/downloads/)

다운로드 링크는 입력한 이메일로 전송됩니다.

### 1-2. Runtime 설치

다운로드한 설치 프로그램을 실행합니다.

설치 프로그램은 다음 작업을 자동으로 처리합니다.

- VTNode Runtime 설치
- `VTNODE_HOME` 환경 변수 설정
- 실행에 필요한 폴더 및 파일 구성

사용자가 `VTNODE_HOME`을 직접 설정할 필요는 없습니다.

### 1-3. 설치 결과 확인

설치 프로그램의 마지막 화면에서 **VTNode 설치 상태 확인** 항목을 선택합니다. 이 항목은 기본적으로 선택되어 있으며, 설치를 마치면 `VTNodeCheck64.exe`가 자동으로 실행됩니다.

```text
☑ VTNode 설치 상태 확인
```

`VTNodeCheck64.exe`는 다음 항목을 검사하고 결과를 표시합니다.

- `VTNODE_HOME` 환경 변수에 설정된 Runtime 경로
- 설치된 VTNode 패키지 버전
- Runtime 설치 경로와 필수 파일
- 설치 상태

설치 완료 화면에서 확인 항목을 해제했거나 나중에 다시 점검하려면 설치 폴더의 `VTNodeCheck64.exe`를 직접 실행할 수 있습니다.

정상 출력 예시는 다음과 같습니다.

```text
VTNode Runtime Installation Check

VTNODE_HOME     : C:\Program Files\VTNode
Package Version : 0.1.0
Required Files  : OK
Status          : OK
```

`Status : ERROR`가 표시되면 출력된 원인을 확인한 후 설치를 다시 점검합니다.

---

## 2. VTNode WAS for XLSX 실행

### 2-1. Excel 예제 다운로드

GitHub에서 `VTNode WAS for XLSX` 예제 패키지를 내려받습니다.

- [VTNode WAS for XLSX on GitHub](https://github.com/uznet/vtnode/tree/main/was/xlsx)

### 2-2. Excel 파일 준비

다운로드한 예제의 XLSM 파일을 준비합니다.

1. 압축 파일을 사용했다면 먼저 압축을 해제합니다.
2. Windows가 파일을 차단한 경우 파일 속성에서 차단을 해제합니다.
3. Excel에서 XLSM 파일을 엽니다.
4. 보안 경고가 표시되면 신뢰할 수 있는 파일인지 확인한 후 매크로를 허용합니다.

### 2-3. 라이선스 없이 바로 실행

VTNode WAS를 시작합니다. 별도의 라이선스 코드를 입력하거나 발급받을 필요가 없습니다.

라이선스 없는 기본 실행 조건은 다음과 같습니다.

- 동시 연결: 1채널
- 사용 기간 제한: 없음
- 기능 제한: 없음

처음 설치한 사용자는 이 방식으로 먼저 동작을 확인하는 것을 권장합니다.

### 2-4. 브라우저 응답 확인

Excel 화면에 표시된 접속 주소를 PC 또는 모바일 브라우저에서 엽니다.

브라우저 요청이 Excel/VBA로 전달되고, VBA가 생성한 결과가 브라우저에 표시되면 Quick Start가 완료된 것입니다.

```text
Browser
   → VTNode Runtime
   → Excel/VBA
   → HTTP Response
   → Browser
```

---

## 선택 사항: 무료 4채널 라이선스

기본 1채널보다 많은 동시 연결이 필요하면 무료 라이선스를 발급받을 수 있습니다.

- 동시 연결: 4채널
- 발급 페이지: [VTNode Free License](https://vtnode.io/license/free-trial)

무료 라이선스 발급 및 활성화는 기본 Quick Start 완료 후 진행해도 됩니다. 무료 라이선스의 발급 조건과 사용 정책은 라이선스 페이지의 최신 안내를 따릅니다.

---

## 전체 흐름

```text
01 Runtime 다운로드
        ↓
02 Runtime 설치
   VTNODE_HOME 자동 설정
        ↓
03 VTNodeCheck64 자동 실행
   설치 경로·패키지 버전·필수 파일 확인
        ↓
04 GitHub 예제 다운로드
        ↓
05 Excel에서 실행
   라이선스 없이 1채널
        ↓
06 브라우저 응답 확인
        ↓
선택: 무료 라이선스 발급 → 4채널
```

---

## 홈페이지 Quick Start 요약 문구

### 01 INSTALL

**VTNode Runtime 설치**

설치 프로그램을 내려받아 실행합니다. `VTNODE_HOME`은 자동으로 설정됩니다.

### 02 VERIFY

**설치 상태 확인**

설치 마지막 화면에서 `VTNodeCheck64.exe`가 자동으로 실행되어 Runtime 경로, 설치된 패키지 버전과 필수 파일을 확인합니다.

### 03 DOWNLOAD EXAMPLE

**Excel 예제 다운로드**

GitHub에서 VTNode WAS for XLSX 예제 패키지를 내려받습니다.

### 04 RUN

**라이선스 없이 실행**

별도 신청 없이 기간과 기능 제한 없는 1채널로 바로 실행합니다.

### 05 TEST

**브라우저 응답 확인**

Excel에 표시된 주소를 열어 VBA가 처리한 응답을 확인합니다.

### OPTIONAL LICENSE

**무료 4채널 라이선스**

더 많은 동시 연결이 필요하면 무료 라이선스를 발급받습니다.

