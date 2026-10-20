# VTNode 무료 라이선스 발급 및 등록

[English](./FREE_LICENSE.md) | **한국어**

VTNode Runtime을 설치한 후 무료 라이선스를 발급받아 등록하는 방법을 설명합니다.

> VTNode는 라이선스 없이도 기간과 기능 제한 없이 1채널로 실행할 수 있습니다. 무료 라이선스에서 제공하는 추가 채널이 필요한 경우에만 이 절차를 진행하세요.

## 전체 흐름

```text
01 VTNode Runtime 설치
        ↓
02 사용할 PC에서 Reg Code 생성
        ↓
03 Reg Code 업로드 및 Meta Code 다운로드
        ↓
04 같은 PC에서 Meta Code 활성화
        ↓
05 발급된 License ID 설정
```

## 시작하기 전에

- 라이선스를 사용할 Windows PC에 VTNode Runtime을 설치합니다.
- `VTNODE_HOME` 환경 변수가 정상적으로 설정되어 있는지 확인합니다.
- 웹 신청은 다른 PC에서 진행해도 되지만, Reg Code 생성과 Meta Code 활성화는 라이선스를 사용할 동일한 PC에서 진행해야 합니다.
- Meta Code를 활성화할 때는 해당 PC가 인터넷에 연결되어 있어야 합니다.

## 1. VTNode Runtime 설치

[VTNode 다운로드 페이지](https://vtnode.io/downloads/)에서 최신 Windows x64 Runtime을 내려받아 설치합니다.

설치 프로그램이 `VTNODE_HOME` 환경 변수를 자동으로 설정합니다. 설치 후 `VTNodeCheck64.exe`를 실행하여 최종 상태가 `OK`인지 확인하세요.

## 2. Reg Code 생성

라이선스를 사용할 PC에서 `VTMakeRegCode64.exe`를 실행합니다.

정상적으로 완료되면 다음 위치에 Reg Code 파일이 생성됩니다.

```text
%VTNODE_HOME%\licenses\vtnode_regcode.vtreg
```

이 파일에는 해당 PC의 장치 식별 정보에서 생성한 해시값이 포함됩니다. 파일 내용을 수정하지 마세요.

> Reg Code 파일 생성만으로 라이선스가 발급되거나 설치되는 것은 아닙니다. 웹 발급 단계를 계속 진행해야 합니다.

## 3. Meta Code 발급 및 다운로드

웹브라우저에서 [VTNode 무료 라이선스 페이지](https://vtnode.io/license/free-trial)를 엽니다.

1. 라이선스를 발급받을 제품을 선택합니다.
2. 앞 단계에서 생성한 `vtnode_regcode.vtreg` 파일을 선택합니다.
3. 라이선스 이용 조건과 장치 해시 처리 안내를 확인하고 동의합니다.
4. Reg Code 파일을 업로드하고 발급 절차를 완료합니다.
5. 발급된 Meta Code(`.vtmeta`) 파일을 내려받습니다.

회원가입이나 이메일 인증은 필요하지 않습니다. 웹 신청을 다른 PC에서 진행했다면 내려받은 Meta Code 파일을 라이선스를 사용할 PC로 복사하세요.

> Meta Code는 실행용 라이선스 파일이 아닙니다. 반드시 대상 PC에서 활성화 단계를 완료해야 합니다.

## 4. Meta Code 활성화

라이선스를 사용할 PC를 인터넷에 연결하고 명령 프롬프트 또는 PowerShell에서 다음 명령을 실행합니다.

```powershell
VTActivateMetaCode64.exe "<Meta-Code-파일-경로>.vtmeta"
```

예시:

```powershell
VTActivateMetaCode64.exe "C:\Users\User\Downloads\<License-ID>.vtmeta"
```

### 활성화가 필요한 이유

내려받은 Meta Code는 발급 정보를 담은 중간 코드이며, Runtime에서 바로 사용할 수 있는 라이선스 파일이 아닙니다. `VTActivateMetaCode64.exe`는 다음 작업을 수행합니다.

- Meta Code가 정상적으로 발급되었고 변경되지 않았는지 확인
- 현재 PC가 Reg Code를 생성한 대상 장치와 일치하는지 확인
- VTNode 라이선스 서버에 최초 활성화 결과 등록
- 서명된 로컬 라이선스 파일(`.vtlic`) 다운로드 및 저장

이 과정을 통해 발급된 사용 권한이 대상 PC에 연결되며, 내려받은 Meta Code나 라이선스 파일을 다른 장치에 단순히 복사하여 사용하는 것을 방지합니다.

### 라이선스 사용기간이 시작되는 시점

무료 라이선스의 사용기간은 `VTActivateMetaCode64.exe`를 실행하여 **최초 온라인 활성화에 성공한 시점**부터 계산됩니다. Reg Code를 생성한 시점이나 Meta Code를 발급 또는 다운로드한 시점부터 계산되는 것이 아닙니다.

```text
Meta Code 발급·다운로드  → 라이선스 기간 시작 전
최초 활성화 성공         → 라이선스 시작일 기록
```

VTNode 라이선스 서버가 최초 활성화 시각을 기록하고 무료 라이선스 정책에 따라 사용기간을 계산합니다. 활성화 프로그램을 다시 실행하더라도 최초 시작일이 초기화되거나 사용기간이 연장되지 않습니다.

활성화 프로그램이 Meta Code와 현재 PC를 확인한 후 VTNode 라이선스 서버에 온라인 활성화를 요청합니다.

활성화에 성공하면 다음 위치에 서명된 라이선스 파일이 생성됩니다.

```text
%VTNODE_HOME%\licenses\<License-ID>.vtlic
```

생성된 License ID를 확인해 두세요. License ID는 `.vtlic` 확장자를 제외한 파일명입니다.

## 5. License ID 설정

VTNode WAS 응용프로그램에서 사용하는 `vtnode_httpa_params.ini` 파일을 열고 발급된 License ID를 설정합니다.

```ini
[license]
enabled = true
id = <License-ID>
```

예시:

```ini
[license]
enabled = true
id = bc0f9373-782d-46da-ab5c-4725b668d152
```

파일을 저장한 후 VTNode WAS 응용프로그램을 다시 시작합니다. Runtime은 `%VTNODE_HOME%\licenses` 폴더에서 해당 License ID의 라이선스 파일을 찾아 실행 시 검증합니다.

## 최종 확인

- `VTNodeCheck64.exe`의 결과가 `Status : OK`입니다.
- 라이선스를 사용할 PC에서 Reg Code를 생성했습니다.
- Reg Code를 생성한 동일한 PC에서 Meta Code를 활성화했습니다.
- 온라인 활성화가 정상적으로 완료되었습니다.
- 최초 활성화 성공 시점이 라이선스 시작일로 기록되었습니다.
- `%VTNODE_HOME%\licenses\<License-ID>.vtlic` 파일이 존재합니다.
- `vtnode_httpa_params.ini`에 동일한 License ID를 설정했습니다.

## 문제가 발생하면

### Reg Code 파일이 생성되지 않는 경우

- `VTNODE_HOME`이 정상적으로 설정되어 있는지 확인합니다.
- `VTNodeCheck64.exe`를 실행하여 Runtime 설치 상태를 확인합니다.
- `licenses` 폴더에 파일을 쓸 수 있는지 확인합니다.

### Meta Code 활성화에 실패하는 경우

- 대상 PC가 인터넷에 연결되어 있는지 확인합니다.
- 해당 PC에서 생성한 Reg Code로 발급받은 Meta Code인지 확인합니다.
- Meta Code 파일 경로가 정확한지 확인하고, 공백이 포함된 경로는 큰따옴표로 감쌉니다.
- `.vtreg` 또는 `.vtmeta` 파일 내용을 수정하지 마세요.

### 응용프로그램이 라이선스를 찾지 못하는 경우

- `%VTNODE_HOME%\licenses` 폴더에 `.vtlic` 파일이 있는지 확인합니다.
- `vtnode_httpa_params.ini`에는 `.vtlic` 확장자를 제외한 License ID를 입력합니다.
- 설정을 변경한 후 Excel과 VTNode WAS 응용프로그램을 다시 시작합니다.

## 관련 링크

- [VTNode 빠른 시작](./QUICK_START.ko.md)
- [VTNode 다운로드](https://vtnode.io/downloads/)
- [무료 라이선스 발급](https://vtnode.io/license/free-trial)
- [VTNode Documents](https://vtnode.io/docs)

