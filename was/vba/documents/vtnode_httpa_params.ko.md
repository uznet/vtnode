# vtnode_httpa_params.ini 설정 설명서

**한국어** | [English](./vtnode_httpa_params.en.md)

확인일: 2026-10-22. 대상: 이 폴더의 [vtnode_httpa_params.ini](./vtnode_httpa_params.ini).

## 근거와 적용 범위

아래 값은 대상 파일을 직접 읽어 확인한 **현재 설정값**이며, 항목을 생략했을 때의 Runtime 기본값을 뜻하지 않습니다. 설명은 INI 주석과 항목명에 근거하며, 라이선스 절차는 같은 폴더의 [무료 라이선스 안내](./FREE_LICENSE.ko.md)를 함께 확인했습니다. 네이티브 INI 파서와 실행 엔진 구현은 확인하지 못했으므로 실제 실행 시험으로 검증한 동작은 아닙니다. 불명확한 내용은 따로 표시했습니다.

## 현재 구성 요약

- HTTP 포트: `12345`. HTTPS: `ssl.port=0`으로 비활성.
- WebSocket: 활성, 메시지 상한 1MB, 프레임 상한 64KB.
- Host 정책: `example.com` 및 별칭 `www.example.com`, `api.example.com`, 임의 Host 허용은 꺼짐.
- 디버그 출력 활성, stdout 콘솔 출력 비활성, 라이선스 사용 비활성.
- `maxConnections=16`이 실제로 16개의 동시 연결을 보장하는 것은 아닙니다. 로컬 라이선스 안내는 무라이선스 실행을 1채널로 설명하며, 이 설정과 라이선스 채널 제한의 정확한 관계는 미확인입니다.

## 표기와 단위

- `[section]`은 섹션, `key=value`는 설정입니다. 점이 있는 섹션명/키(`httpserver.timeout`, `ssl.port`)는 원문 그대로 사용합니다.
- `;`로 시작하는 줄과 값 뒤 `;` 부분은 이 파일에서 주석으로 쓰입니다. `;names= chat,echo`는 활성 설정이 아닙니다.
- 원문 기준 `1KB=1024`바이트, `1MB=1024×1024`바이트입니다. 4KB=4,096, 8KB=8,192, 64KB=65,536, 1MB=1,048,576, 10MB=10,485,760바이트입니다.
- `10s`, `30s`, `60s`, `600s`는 주석 기준 초 단위입니다. 여기서 사용된 표기 외의 단위·대소문자·공백·따옴표 처리 규칙은 파서에서 확인하지 못했습니다.
- 0의 의미는 항목별로 다릅니다. 명시된 경우만 무제한/미사용으로 해석하며, 다른 항목에 일반화하지 마세요.
- idle 타임아웃은 진행이 없는 시간을 제한합니다. 전체 처리 시간 상한과 구분해야 하며, 내부 타이머의 갱신 조건은 미확인입니다.

## 설정 항목

### [http] — HTTP 요청 제한

| 항목 | 현재 값 | 설명 |
| --- | --- | --- |
| `maxRequestLineBytes` | `4KB` | 요청 시작 줄 최대 크기. 주석은 CRLF 포함을 권장합니다. |
| `maxHeaderLineBytes` | `4KB` | 단일 헤더 줄 최대 크기. CRLF를 포함합니다. |
| `maxHeaderCount` | `100` | 허용되는 최대 헤더 개수. |
| `maxHeaderTotalBytes` | `8KB` | 전체 헤더 섹션 최대 크기. 요청 시작 줄 제외, CRLFCRLF 포함. |
| `maxBodyBytes` | `10MB` | 요청 본문 최대 크기. |
| `maxTransactionsPerConnection` | `10` | 연결 하나당 최대 트랜잭션 수. 0은 무제한. |


### [httpserver.timeout] — HTTP 서버 타임아웃

| 항목 | 현재 값 | 설명 |
| --- | --- | --- |
| `requestFirstByteTimeoutSec` | `10s` | 첫 요청 줄의 첫 바이트를 기다리는 최대 시간. |
| `requestHeaderTimeoutSec` | `30s` | 첫 바이트 수신부터 전체 요청 헤더 완료까지의 시간. |
| `requestBodyIdleTimeoutSec` | `60s` | 요청 본문 수신 중 데이터 진행이 없는 시간의 제한. |
| `transactionTimeoutSec` | `60s` | 요청 수신 완료부터 응답 헤더와 본문 전송 완료까지의 제한. |
| `keepAliveTimeoutSec` | `60s` | 응답 완료 후 같은 연결에서 다음 요청을 기다리는 시간. |


### [httpclient.timeout] — HTTP 클라이언트 타임아웃

| 항목 | 현재 값 | 설명 |
| --- | --- | --- |
| `requestSendIdleTimeoutSec` | `30s` | 요청 전송 중 진행이 없는 시간의 제한. |
| `responseHeaderTimeoutSec` | `30s` | 주석상 요청 전송 완료부터 응답 헤더/첫 바이트까지의 대기(TTFB). 정확한 종료 조건은 구현 확인 필요. |
| `responseBodyIdleTimeoutSec` | `60s` | 응답 본문 수신 중 진행이 없는 시간의 제한. |
| `overallTimeoutSec` | `0` | 전체 왕복 시간 상한. 현재 0이므로 이 상한은 미사용. 다른 개별 타임아웃은 유지됩니다. |


### [websocket] — WebSocket 활성화 및 크기 제한

| 항목 | 현재 값 | 설명 |
| --- | --- | --- |
| `enabled` | `true` | WebSocket 활성. false로 비활성화. |
| `maxMessageBytes` | `1MB` | 분할 프레임을 재조립한 전체 메시지 크기 상한. |
| `maxFrameBytes` | `64KB` | 단일 프레임 크기 상한. |
| `maxMissingPongCount` | `1` | 허용되는 연속 pong 미응답 횟수. 실제 종료 임계 비교 방식은 미확인. |


### [websocket.timeout] — WebSocket 타임아웃

| 항목 | 현재 값 | 설명 |
| --- | --- | --- |
| `closeTimeoutSec` | `10s` | close 프레임 전송 후 상대 close 응답 대기 시간. |
| `idleTimeoutSec` | `600s` | 프레임 송수신이 없는 유휴 시간 제한(10분). |
| `pongTimeoutSec` | `10s` | ping 이후 pong 응답 대기 시간. ping 주기는 이 파일에 없습니다. |


### [websocket.subprotocol] — WebSocket 하위 프로토콜 협상

| 항목 | 현재 값 | 설명 |
| --- | --- | --- |
| `policy` | `optional` | 주석에 optional / required 선택지가 있습니다. 현재는 하위 프로토콜 협상을 선택적으로 요구하는 설정입니다. |
| `serverPreference` | `true` | 서버 우선 선택. false는 서버 우선 선택을 끄는 값이며 상세 선택 순서는 미확인. |
| `onUnsupported` | `ignore` | 미지원 하위 프로토콜 처리. 주석의 선택지는 ignore(무시), reject(거절). |

주석 예시: `;names= chat,echo`. 지원 이름 목록은 현재 명시되어 있지 않습니다. 목록 생략 시 내장 목록 사용 여부, `required`와 `ignore` 조합의 우선순위, 협상 실패 응답은 구현 확인이 필요합니다.


### [net] — 네트워크 및 TLS

| 항목 | 현재 값 | 설명 |
| --- | --- | --- |
| `maxConnections` | `16` | 이름상 최대 연결 수 설정. 주석이 비어 있어 적용 범위와 라이선스와의 우선순위는 미확인. |
| `port` | `12345` | 일반 HTTP 포트. 주석상 0보다 크면 활성화. |
| `ssl.port` | `0` | HTTPS 포트. 현재 비활성, 0보다 크면 활성화. |
| `ssl.certFile` | `server.crt` | 서버 인증서 파일 경로. 상대 경로 기준 디렉터리는 미확인. |
| `ssl.keyFile` | `server.key` | 서버 개인키 파일 경로. 상대 경로 기준 디렉터리는 미확인. |


### [log] — 로그 출력

| 항목 | 현재 값 | 설명 |
| --- | --- | --- |
| `debugOut` | `true` | 디버그 출력 활성. 출력 대상 및 파일 로그와의 관계는 이 파일에 없습니다. |
| `consoleOut` | `false` | 표준 출력(stdout) 콘솔 로그 비활성. |


### [host.policy] — 서버 Host 정책

| 항목 | 현재 값 | 설명 |
| --- | --- | --- |
| `primaryHost` | `example.com` | 주 호스트명. 현재 예제 도메인입니다. |
| `allowAnyHost` | `false` | 임의 Host 허용 비활성. true면 Host 헤더와 무관하게 허용(주석상 프록시/테스트용). |
| `aliasHosts` | `www.example.com , api.example.com` | 추가 호스트명 목록. 현재 두 개가 쉼표로 구분되어 있습니다. |


### [license] — 라이선스

| 항목 | 현재 값 | 설명 |
| --- | --- | --- |
| `enabled` | `false` | 라이선스 사용 비활성. 발급된 라이선스를 사용할 때 true로 설정합니다. |
| `id` | `"your-license-id"` | 발급된 License ID 자리표시자. .vtlic 확장자를 제외한 ID를 입력합니다. |

## 설정 적용과 확인

1. 실제 응용프로그램이 읽는 INI 위치를 확인하고 백업합니다. 이 quick-start 파일이 자동으로 실행 환경에 적용된다고 가정하지 마세요.
2. 필요한 항목만 변경하고 응용프로그램을 다시 시작한 뒤 결과를 확인하세요. 로컬 라이선스 안내는 재시작을 요구합니다. 다른 항목의 자동 재적용 지원 여부는 미확인이므로 재시작 후 확인하는 방식을 권장합니다.
3. HTTP 접근 시 포트와 Host 정책을 함께 확인하세요. 현재 설정의 `localhost` 또는 `127.0.0.1` 요청은 허용 목록과 일치하지 않을 수 있습니다. 아래는 로컬 서버를 대상으로 Host를 지정하는 수동 확인 예시이며, 이 문서를 작성하면서 실행하지 않았습니다.

```powershell
curl.exe -v -H "Host: example.com" http://127.0.0.1:12345/
```

루트 경로 `/`의 응답 내용과 상태는 응용프로그램 라우팅에 따라 다릅니다. Host 거절 상태 코드, 대소문자·포트·와일드카드 매칭 규칙은 확인하지 못했습니다.

### HTTPS 사용

HTTPS를 사용하려면 `ssl.port`를 사용할 양수 포트로 지정하고, 유효한 인증서와 개인키 경로를 설정해야 합니다. 현재 파일에 경로가 적혀 있어도 HTTPS는 꺼져 있습니다. 상대 경로 기준, 인증서 형식, TLS 버전, 인증서 오류 시 처리 방식은 이 INI에서 확인할 수 없습니다.

### 라이선스 사용

[무료 라이선스 안내](./FREE_LICENSE.ko.md)에 따라 대상 PC에서 Meta Code를 활성화한 후, `[license]`의 `enabled=true`와 발급된 `id`를 설정하고 응용프로그램을 다시 시작합니다. 해당 안내에 따르면 Runtime은 `%VTNODE_HOME%\licenses\<License-ID>.vtlic`을 찾아 실행 시 검증합니다. `"your-license-id"`는 예시이며 실제 발급 ID가 아닙니다.

## 문제 확인 시 참고

| 증상 | 우선 확인할 설정 |
| --- | --- |
| HTTP 연결이 안 됨 | 실행 중인 서버, `net.port`, 포트 사용 충돌, 방화벽 |
| localhost 접근이 허용되지 않음 | `host.policy`의 주 호스트·별칭·`allowAnyHost` |
| 큰 요청이 처리되지 않음 | 요청 줄·헤더 개수·헤더 총량·본문의 각 상한 |
| 연결 또는 응답이 중간에 끝남 | 서버/클라이언트 타임아웃, 연결당 트랜잭션 제한 |
| WebSocket 협상/메시지 처리 실패 | 활성화, 하위 프로토콜 정책, 프레임·메시지 상한, pong/idle 제한 |
| 콘솔에 로그가 없음 | 현재 `consoleOut=false`; `debugOut`의 실제 출력 대상은 별도 확인 |
| 설정 수보다 동시 연결이 적음 | 라이선스 채널 수와 실행 로그; `maxConnections`만으로 판단하지 않기 |

상한 초과 시 HTTP 상태 코드, WebSocket close 코드, 연결 종료 시점, 누락/잘못된 값의 대체값 및 허용 범위는 구현이나 실행 로그로 추가 확인해야 합니다. 이 설명서는 원본 INI를 변경하지 않고 작성했습니다.
