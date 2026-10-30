[English](README.MD) | [한글](README.ko.md)

[VBA 프로젝트](https://github.com/uznet/vtnode/tree/main/was/vba) | [현재 폴더](https://github.com/uznet/vtnode/tree/main/was/vba/xlsx/vtnode_was)

> **[필독 — 다운로드한 ZIP 파일은 압축을 해제하기 전에 차단을 해제하세요](../README.ko.md)**

# VTNode WAS for XLSX (VBA)

## 공통 VBA 직접 Import 및 빌드

확정된 구조는 `was/vba/common/src`(Office 공통 VBA), `was/vba/common/xlsx`(Excel 전용), `was/vba/common/pptx`(PowerPoint 전용), `was/vba/xlsx/`(Excel 앱/예제)입니다. 앱 폴더의 `src/`는 앱 전용 코드이며 공통 소스 전체를 대신하지 않습니다.

1. 대상 파일을 백업합니다. Excel UI 원본 `.xlsx`는 `.xlsm`으로 저장하고, PowerPoint는 `.pptm`, Word는 `.docm` 등 VBA를 보존하는 형식을 사용합니다.
2. 대상 파일에서 VBA 편집기를 열고 **파일 > 파일 가져오기**로 [common/src](../../common/src/)의 필요한 `.bas`·`.cls`와 의존 모듈을 해당 VBA 프로젝트에 직접 Import합니다.
3. Excel에는 [common/xlsx](../../common/xlsx/), PowerPoint에는 [common/pptx](../../common/pptx/)의 해당 호스트 전용 소스를 추가합니다. Excel 전용 코드를 다른 Office 프로젝트에 Import하지 않습니다. Word는 Office 공통 소스를 사용하고 필요한 호스트 연결 코드를 별도로 구현·검증합니다.
4. 앱의 `src/`에서 앱 전용 소스를 추가합니다. 같은 이름의 모듈은 백업 후 교체하여 중복을 피합니다. UserForm은 `.frm`과 `.frx`를 같은 폴더에 두고 `.frm`을 Import합니다. `ThisWorkbook`·시트·문서 이벤트 코드는 대상 문서 모듈에 반영합니다.
5. 외부 런타임 Workbook을 열거나 VBA 프로젝트 참조로 연결하지 않습니다. 기존의 외부 런타임 Workbook 참조는 **도구 > 참조**에서 제거하고 필요한 라이브러리 참조를 확인합니다. UI·매크로 연결도 대상 파일 내부 프로시저를 사용하도록 확인합니다.
6. **디버그 > VBAProject 컴파일** 후 저장합니다. Office와 일치하는 비트 수의 네이티브 VTNode Runtime을 준비하고, `www`·`audio`·설정 파일의 상대 배치를 유지한 상태에서 시작·종료, HTTP/WebSocket, Runtime UI를 확인합니다. Office 공통 소스라는 분류만으로 모든 Office에서의 실행이 검증된 것은 아닙니다.


> **Enable web service with VBA.**

`vtnode_was.xlsm`은 **VTNode WAS for XLSX (VBA)**의 Main Demo Workbook입니다.

HTTP / WebSocket Server를 실행하고, 여러 Application을 실행하면서  
**HTTP Request / Response와 WebSocket Connection Session을 Excel 화면에서 직접 확인할 수 있습니다.**

### [Download Runtime](https://vtnode.io/downloads/) | [Quick Start](https://github.com/uznet/vtnode/blob/main/was/vba/documents/QUICK_START.ko.md) | [Build My App](../vtnode_build_myapp/README.ko.md) | [User Manual](https://vtnode.io/docs/docs_login)

---

# What is `vtnode_was.xlsm`?

`vtnode_hello`는 가장 작은 Quick Start Example입니다.

```text
Excel
  ↓
START
  ↓
Browser
  ↓
Hello, VTNode!
```

`vtnode_was.xlsm`은 그 다음 단계입니다.

단순히 Web Page를 실행하는 것을 넘어:

```text
                    vtnode_was.xlsm
                           │
             ┌─────────────┴─────────────┐
             │                           │
             ▼                           ▼
        APPLICATIONS                  MONITOR
             │                           │
      ┌──────┼──────┐           ┌────────┴────────┐
      │      │      │           │                 │
    Intro   WS    Table       HTTP             WebSocket
           Demo   Order    Request/Response     Sessions
```

Application을 실행하는 동시에  
**Network Activity를 Excel에서 직접 확인할 수 있습니다.**

---

# Main Screen

`vtnode_was.xlsm`의 Main Screen은 크게 네 영역으로 구성됩니다.

```text
┌───────────────────────────────────────────────────────────┐
│                 VTNode WAS for Excel · MVP                │
├──────────────────────────────┬────────────────────────────┤
│ SERVER / RUNTIME             │ APPLICATIONS               │
│                              │                            │
│ START SERVER   STOP SERVER   │ WHAT IS VTNODE WAS?       │
│ SHOW RUNTIME   STATUS        │ WEBSOCKET DEMO            │
│ ACCESS URL                  │ TABLE ORDER               │
├──────────────────────────────┴────────────────────────────┤
│ REQUEST → RESPONSE ACTIVITY                              │
│                                                         │
│ TIME | METHOD | REQUEST/URI | RESPONSE/STATUS | ...     │
├───────────────────────────────────────────────────────────┤
│ WEBSOCKET CONNECTION SESSIONS                           │
│                                                         │
│ ID | STATUS | CONNECTED AT | RECEIVED | SEND | ...     │
└───────────────────────────────────────────────────────────┘
```

이 화면 하나에서:

- Server를 시작하고 중지할 수 있습니다.
- VTNode Runtime 상태를 확인할 수 있습니다.
- ACCESS URL을 확인할 수 있습니다.
- Example Application을 실행할 수 있습니다.
- HTTP Request / Response Activity를 확인할 수 있습니다.
- WebSocket Connection Session을 확인할 수 있습니다.

---

# Server / Runtime

왼쪽 상단의 **SERVER / RUNTIME** 영역은  
VTNode WAS Server의 실행 상태를 제어하고 확인합니다.

```text
[ START SERVER ]     [ STOP SERVER ]

[ SHOW RUNTIME ]     SERVER STATUS
                     STOPPED / RUNNING

ACCESS URL
http://192.168.x.x:12345/

NETWORK
HTTP OFF / HTTPS OFF
```

### START SERVER

VTNode WAS HTTP / WebSocket Server를 시작합니다.

### STOP SERVER

실행 중인 Server를 중지합니다.

### SHOW RUNTIME

VTNode Runtime Console을 표시합니다.

### SERVER STATUS

현재 Server의 실행 상태를 표시합니다.

### ACCESS URL

Browser 또는 같은 Network의 다른 Device에서  
접속할 수 있는 URL을 표시합니다.

---

# Applications

오른쪽 상단의 **APPLICATIONS** 영역에서는  
하나의 Excel Workbook에서 여러 Application을 실행할 수 있습니다.

```text
APPLICATIONS

[ WHAT IS VTNODE WAS? / INTRO ]

[ WEBSOCKET DEMO ]

[ TABLE ORDER ]

[ + YOUR APP ]
```

각 Application은 자신의 URI Route를 가질 수 있습니다.

```text
Excel Apps
     │
     ├── Intro
     ├── WebSocket Demo
     ├── Table Order
     └── Your App
              │
              ▼
        URI Routing
```

즉, 하나의 VTNode WAS에서 여러 Excel/Web Application으로  
확장할 수 있다는 개념을 보여줍니다.

---

# Request → Response Activity

`vtnode_hello`와 비교했을 때  
`vtnode_was.xlsm`의 중요한 차이 중 하나입니다.

HTTP Request가 들어오고 Response가 반환되는 과정을  
Excel 화면에서 실시간으로 확인할 수 있습니다.

```text
Browser
   │
   │ HTTP Request
   ▼
VTNode WAS
   │
   ▼
Excel / VBA
   │
   ▼
VTNode WAS
   │
   │ HTTP Response
   ▼
Browser

        +
        │
        ▼

REQUEST → RESPONSE ACTIVITY
```

Log에는 다음과 같은 정보가 표시됩니다.

```text
TIME
METHOD
REQUEST / URI
RESPONSE / STATUS
DURATION
DETAIL
```

예를 들어 Browser에서 Page를 요청하면:

```text
GET
   ↓
/index.html
   ↓
Excel / VBA
   ↓
200 OK
```

와 같은 Request / Response 흐름을 Excel에서 확인할 수 있습니다.

---

# See HTTP Working

VTNode WAS의 HTTP 처리는 화면 뒤에서만 일어나지 않습니다.

Browser에서 Application을 사용하면서  
Excel 화면의 **REQUEST → RESPONSE ACTIVITY**를 보면:

```text
Browser Action
      │
      ▼
HTTP Request
      │
      ▼
Excel Log
      │
      ▼
VBA Execute
      │
      ▼
HTTP Response
      │
      ▼
Excel Log
```

의 흐름을 직접 관찰할 수 있습니다.

따라서 `vtnode_was.xlsm`은 Application Demo인 동시에  
VTNode WAS의 동작을 이해하기 위한 **Live Monitor** 역할도 합니다.

---

# WebSocket Connection Sessions

두 번째 Monitoring 영역은:

**WEBSOCKET CONNECTION SESSIONS**

입니다.

WebSocket Client가 연결되면 각 Connection을 Session 단위로  
Excel 화면에서 확인할 수 있습니다.

```text
Browser
   │
   │ WebSocket Connect
   ▼
VTNode WAS
   │
   ▼
Excel / VBA

        +
        │
        ▼

WEBSOCKET CONNECTION SESSIONS
```

Session 영역에서는 다음 정보를 확인할 수 있습니다.

```text
ID
STATUS
CONNECTED AT
RECEIVED MESSAGE
SEND MESSAGE
SEND
LAST ACTIVITY
```

---

# WebSocket Sessions Stay Visible

WebSocket은 HTTP Request / Response와 동작 방식이 다릅니다.

HTTP는:

```text
Request
   ↓
Response
   ↓
Completed
```

형태이지만 WebSocket은:

```text
Connect
   ↓
OPEN
   ↓
Receive / Send
   ↓
Receive / Send
   ↓
...
   ↓
CLOSED
```

처럼 Connection을 유지합니다.

따라서 `vtnode_was.xlsm`에서는 WebSocket Connection을  
**Session 단위로 관리하고 관찰할 수 있습니다.**

```text
Session #1    OPEN
Session #2    OPEN
Session #3    CLOSED
```

Connection이 종료된 후에도 해당 Row를 유지하여  
마지막 상태와 Activity를 확인할 수 있습니다.

완료된 Session은 **CLEAR CLOSED**를 사용하여 정리할 수 있습니다.

---

# HTTP and WebSocket at a Glance

Main Screen에서 두 Network Model을 동시에 비교할 수 있습니다.

```text
HTTP
────────────────────────────────

Browser
   │
   │ Request
   ▼
Excel / VBA
   │
   │ Response
   ▼
Browser

REQUEST → RESPONSE ACTIVITY


WebSocket
────────────────────────────────

Browser
   │
   ⇅
Persistent Connection
   ⇅
Excel / VBA

WEBSOCKET CONNECTION SESSIONS
```

이 두 영역은 VTNode WAS for XLSX의 Connectivity를  
눈으로 확인하기 위한 핵심 Monitoring UI입니다.

---

# WebSocket Demo

**WEBSOCKET DEMO**는 WebSocket Connection과  
Message 송수신을 직접 확인하기 위한 Example입니다.

```text
Browser
       ⇅
   WebSocket
       ⇅
VTNode WAS
       ⇅
Excel / VBA
```

Browser에서 Connection을 생성하면  
Excel의 **WEBSOCKET CONNECTION SESSIONS**에 Session이 나타납니다.

Message를 주고받으면:

```text
RECEIVED MESSAGE

SEND MESSAGE

LAST ACTIVITY
```

등을 통해 현재 Session의 동작을 확인할 수 있습니다.

즉, WebSocket 기능을 단순히 설명하는 것이 아니라  
**실제 Connection을 Excel에서 눈으로 볼 수 있습니다.**

---

# Table Order

**TABLE ORDER**는 VTNode WAS를 실제 Application 형태로  
확장하는 Example입니다.

```text
Smartphone Browser
        │
        │ HTTP / Web
        ▼
    VTNode WAS
        │
        ▼
    Excel / VBA
        │
        ▼
    Excel Data
```

사용자는 Smartphone Browser에서 Application을 사용하고,  
Excel에서는 Application Logic과 Data를 처리합니다.

동시에 HTTP Activity가 발생하면 Main Screen의:

```text
REQUEST → RESPONSE ACTIVITY
```

에서 확인할 수 있습니다.

따라서 Table Order는 단순한 UI Demo가 아니라:

```text
Web UI
   +
HTTP
   +
Excel/VBA
   +
Excel Data
```

가 하나의 Application으로 연결되는 것을 보여줍니다.

---

# One Workbook, Multiple Apps

`vtnode_was.xlsm`의 중요한 개념 중 하나는  
**하나의 Excel Workbook에서 여러 Application을 운영할 수 있다는 것**입니다.

```text
                    Browser
                       │
                       ▼
                  VTNode WAS
                       │
                 URI Routing
                       │
          ┌────────────┼────────────┐
          │            │            │
          ▼            ▼            ▼
        Intro       WebSocket     Table Order
                                     │
                                     ▼
                                  Excel Data
```

각 Application은 URI에 따라 분기할 수 있습니다.

Main Screen의:

```text
+ YOUR APP
ADD ANY URI ROUTE
```

는 이 구조가 새로운 Application으로 계속 확장될 수 있다는 것을 의미합니다.

---

# Repository Structure

```text
vtnode_was/
│
├─ vtnode_was.xlsm
│
├─ src/
│  ├─ *.bas
│  ├─ *.cls
│  └─ *.frm
│
├─ audio/
│  ├─ audio_tick.wav
│  └─ audio_was_started.wav
│
└─ www/
   ├─ default.html
   ├─ what-is-vtnode-was.html
   │
   ├─ images/
   │  └─ ...
   │
   └─ intro/
      └─ index.html
```

### `vtnode_was.xlsm`

Main Demo / Application Portal / Network Activity Monitor.

### `src/`

Workbook을 구성하는 VBA Module, Class 및 UserForm Source.

### `audio/`

Application과 Runtime에서 사용하는 Audio Resources.

### `www/`

Browser에 서비스되는 HTML / CSS / JavaScript 및 Media Contents.

---

# What Makes `vtnode_was` Different?

VTNode WAS for XLSX의 Example들은 각각 목적이 다릅니다.

```text
vtnode_hello
────────────────────────
Run It

가장 작은 Example을
바로 실행해 본다.


vtnode_build_myapp
────────────────────────
Build It

UI와 제공된 VBA Source를
직접 조립해 본다.


vtnode_was
────────────────────────
Explore It

여러 Application을 실행하고

HTTP Request / Response

WebSocket Sessions

을 실시간으로 관찰한다.
```

`vtnode_was.xlsm`에서는 단순히 결과만 보는 것이 아니라  
**Browser와 Excel 사이에서 실제로 무엇이 일어나고 있는지 볼 수 있습니다.**

---

# Excel Becomes a Network Node

```text
                        Browser
                           │
                  HTTP / WebSocket
                           │
                           ▼
                    VTNode WAS
                           │
                           ▼
                         Excel
                           │
             ┌─────────────┼─────────────┐
             │             │             │
             ▼             ▼             ▼
           Apps           Data          Logs
                                         │
                              ┌──────────┴──────────┐
                              │                     │
                              ▼                     ▼
                         HTTP Activity      WebSocket Sessions
```

Excel은 Application Logic과 Data를 가지고 있으면서  
동시에 Network Activity를 직접 확인할 수 있는 실행 환경이 됩니다.

**Excel becomes a Network Node.**

---

# No Apache. No Node.js. No Python.

**No Python**은 이 Demo를 실행하는 데 Python 설치나 Python 지식이 필요 없다는 뜻입니다.
Python 연동을 금지한다는 의미가 아니며, 필요한 경우 Python과 연동할 수 있습니다.

별도의 Apache나 Node.js Application Server를 구성하지 않고  
VTNode Runtime과 Excel/VBA를 이용합니다.

```text
Microsoft Excel
       +
      VBA
       +
 VTNode Runtime
       ↓
HTTP / WebSocket
       ↓
Browser / Smartphone
```

---

# Where Should I Start?

처음 사용하는 경우 `vtnode_was.xlsm`보다  
먼저 `vtnode_hello`부터 시작하는 것을 권장합니다.

```text
01. vtnode_hello
        │
        │ Run It
        ▼
02. vtnode_build_myapp
        │
        │ Build It
        ▼
03. vtnode_was
        │
        │ Explore It
        ▼
04. Your Application
        │
        │ Change It
        ▼
    Your Excel + Web
```

### [Quick Start](https://github.com/uznet/vtnode/blob/main/was/vba/documents/QUICK_START.ko.md)

### [Build My App](../vtnode_build_myapp/README.ko.md)

---

# What's Next?

VTNode WAS의 Connectivity는 Web에서 더 확장될 수 있습니다.

```text
Hello
  ↓
Build My App
  ↓
Table Order
  ↓
WebSocket
  ↓
AI Connectivity
  │
  ├─ MCP Server
  │
  └─ MCP + SSE
  ↓
Voice / Phone
```

HTTP / WebSocket에서 시작하여  
Local Excel Data와 Cloud AI를 연결하는 Connectivity로 확장할 수 있습니다.

다음에 살펴볼 관련 주제로 **VTNode XEvent for XLSX (VBA)**도 있습니다.

---

# Documentation

VTNode WAS for XLSX를 직접 구성하는 단계별 과정은  
공식 Documentation에서 제공합니다.

### [Download User Manual](https://vtnode.io/docs/docs_login)

```text
vtnode.io
    ↓
VTNode WAS for XLSX
    ↓
Documentation
    ↓
User Manual
```

---

# Start Here

### [Download Runtime](https://vtnode.io/downloads/) | [Quick Start](https://github.com/uznet/vtnode/blob/main/was/vba/documents/QUICK_START.ko.md) | [Build My App](../vtnode_build_myapp/README.ko.md) | [User Manual](https://vtnode.io/docs/docs_login)

---

# VTNode WAS for XLSX (VBA)

> **Enable web service with VBA.**

**Excel becomes a Network Node.**

**VTNode is an AI Glue Gun.**

[VTNode Website](https://vtnode.io)

Copyright © 2026 VTNode
