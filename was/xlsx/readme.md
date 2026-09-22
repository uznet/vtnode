# VTNode WAS for XLSX (VBA)

> **Enable web service with VBA.**

VTNode WAS for XLSX (VBA)는 Microsoft Excel/VBA를  
HTTP / WebSocket 기반 Web Application과 연결하기 위한  
VTNode Runtime 환경입니다.

별도의 Apache, Node.js 또는 Cloud Application Server 없이  
Excel/VBA에서 Web Service를 실행할 수 있습니다.

### [Download Runtime](https://vtnode.io/download/runtime) | [Quick Start](./vtnode_hello/) | [Build My App](./vtnode_build_myapp/) | [Main Demo](./vtnode_was/) | [User Manual](https://vtnode.io/docs/was-xlsx/manual)

---

# Start Here

처음 사용하는 경우 가장 작은 Example부터 시작하세요.

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
    Excel + Web
```

처음부터 Runtime API나 내부 구조를 이해할 필요는 없습니다.

**먼저 실행하고, 직접 만들어 보고, 그 다음 확장해 보세요.**

---

# What is VTNode WAS for XLSX?

VTNode WAS for XLSX의 기본 구조는 단순합니다.

```text
Browser
   │
   │ HTTP / WebSocket
   ▼
VTNode WAS Engine
   │
   ▼
Excel / VBA
   │
   ▼
VTNode WAS Engine
   │
   │ HTTP Response / WebSocket
   ▼
Browser
```

Browser에서 들어온 Request 또는 Event를  
Excel/VBA Application에서 처리할 수 있습니다.

Excel/VBA 관점에서는 다음과 같이 생각할 수 있습니다.

```text
Request Event
     │
     ▼
Excel / VBA Execute
     │
     ▼
Response Event
```

Excel 사용자는 Web Server 내부 구현보다:

```text
Request가 왔을 때
        ↓
URI / Content 확인
        ↓
Excel / VBA에서 처리
        ↓
Response
```

에 집중할 수 있습니다.

---

# Repository Structure

```text
vtnode/
└─ was/
   └─ xlsx/
      │
      ├─ vtnode_hello/
      │  ├─ vtnode_hello.xlsm
      │  ├─ src/
      │  ├─ audio/
      │  └─ www/
      │
      ├─ vtnode_build_myapp/
      │  ├─ vtnode_ui_only.xlsx
      │  ├─ src/
      │  ├─ audio/
      │  └─ www/
      │
      ├─ vtnode_was/
      │  ├─ vtnode_was.xlsm
      │  ├─ src/
      │  ├─ audio/
      │  └─ www/
      │
      └─ vtnode_vba_rt64/
         ├─ vtnode_vba_rt64.xlsm
         └─ src/
```

각 Folder는 서로 다른 목적을 가지고 있습니다.

| Folder | Purpose | Start With |
|---|---|---|
| [`vtnode_hello`](./vtnode_hello/) | 가장 작은 Quick Start Example | `vtnode_hello.xlsm` |
| [`vtnode_build_myapp`](./vtnode_build_myapp/) | 제공된 구성요소를 직접 조립하는 Hands-on Project | `vtnode_ui_only.xlsx` |
| [`vtnode_was`](./vtnode_was/) | Applications + Runtime + Live Monitor | `vtnode_was.xlsm` |
| [`vtnode_vba_rt64`](./vtnode_vba_rt64/) | 공통 64-bit VBA Runtime | `vtnode_vba_rt64.xlsm` |

---

# 1. VTNode Hello — Run It

[`vtnode_hello`](./vtnode_hello/)는  
VTNode WAS for XLSX를 처음 사용하는 사용자를 위한 가장 작은 Example입니다.

목표는 하나입니다.

```text
Excel
  ↓
START
  ↓
Browser
  ↓
Hello, VTNode!
```

먼저 VTNode Runtime을 설치한 다음:

```text
vtnode_hello.xlsm
```

을 실행합니다.

START를 클릭하고 ACCESS URL을 Browser에서 열었을 때:

```text
Hello, VTNode!

This page is served directly from Excel/VBA.

VTNode WAS for XLSX (VBA)

● Server is running
```

이 표시되면 성공입니다.

### [Open vtnode_hello](./vtnode_hello/)

---

# 2. Build My App — Build It

완성된 `vtnode_hello.xlsm`을 실행해 보았다면  
다음에는 직접 만들어 볼 수 있습니다.

[`vtnode_build_myapp`](./vtnode_build_myapp/)은  
VTNode WAS Application의 구성요소를 직접 조립해 보는 Hands-on Project입니다.

시작 파일은:

```text
vtnode_ui_only.xlsx
```

입니다.

이 Workbook의 **MyApp** Sheet는 `vtnode_hello`를 기본으로 한 UI를 가지고 있지만  
VTNode WAS 기능은 아직 연결되어 있지 않습니다.

```text
vtnode_ui_only.xlsx
        │
        │
        │  Provided VBA Components
        │  *.bas / *.cls / *.frm
        │
        ▼
vtnode_myapp.xlsm
```

이 과정의 목적은 Runtime API를 공부하는 것이 아닙니다.

준비된 UI, VBA Source, Web Contents가  
어떻게 하나의 VTNode WAS Application으로 연결되는지를 직접 경험하는 것입니다.

```text
UI
 +
VBA
 +
Web Contents
 +
VTNode Runtime
        ↓
VTNode WAS Application
```

### [Open vtnode_build_myapp](./vtnode_build_myapp/)

상세한 제작 과정은 공식 Build Manual에서 제공합니다.

### [Download Build Manual](https://vtnode.io/docs/was-xlsx/build-myapp)

---

# 3. VTNode WAS — Explore It

[`vtnode_was`](./vtnode_was/)는  
VTNode WAS for XLSX의 Main Demo Workbook입니다.

```text
vtnode_was.xlsm
```

에서는 여러 Application을 실행하면서  
HTTP / WebSocket Network Activity를 Excel에서 직접 확인할 수 있습니다.

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

Main Application에는 다음 Example이 포함됩니다.

```text
WHAT IS VTNODE WAS?

WEBSOCKET DEMO

TABLE ORDER

+ YOUR APP
```

---

## HTTP Request / Response Monitor

`vtnode_was.xlsm`에서는 Browser에서 발생하는 HTTP Activity를  
Excel 화면에서 실시간으로 확인할 수 있습니다.

```text
REQUEST → RESPONSE ACTIVITY

TIME
METHOD
REQUEST / URI
RESPONSE / STATUS
DURATION
DETAIL
```

따라서:

```text
Browser Action
      ↓
HTTP Request
      ↓
Excel / VBA
      ↓
HTTP Response
      ↓
Browser
```

의 흐름을 직접 관찰할 수 있습니다.

---

## WebSocket Session Monitor

WebSocket Connection도 Session 단위로 확인할 수 있습니다.

```text
WEBSOCKET CONNECTION SESSIONS

ID
STATUS
CONNECTED AT
RECEIVED MESSAGE
SEND MESSAGE
LAST ACTIVITY
```

HTTP와 WebSocket의 차이를 같은 Excel Workbook에서 직접 확인할 수 있습니다.

```text
HTTP

Request
   ↓
Response
   ↓
Completed


WebSocket

Connect
   ↓
OPEN
   ↓
Receive / Send
   ↓
Receive / Send
   ↓
CLOSED
```

`vtnode_was.xlsm`은 단순한 Demo Launcher가 아니라  
VTNode WAS의 Network Activity를 볼 수 있는 **Live Monitor** 역할도 합니다.

### [Open vtnode_was](./vtnode_was/)

---

# 4. VTNode VBA Runtime 64

[`vtnode_vba_rt64`](./vtnode_vba_rt64/)는  
VTNode WAS for XLSX Application에서 공통으로 사용하는  
64-bit Excel/VBA Runtime입니다.

Application Example이 아니라  
Excel Application과 Native VTNode Runtime 사이의 공통 Runtime Layer입니다.

```text
Excel Application
        │
        ▼
vtnode_vba_rt64
        │
        ▼
VTNode Runtime
        │
        ▼
HTTP / WebSocket
```

여러 Application에서 공통으로 사용할 수 있습니다.

```text
                   vtnode_vba_rt64
                          │
             ┌────────────┼────────────┐
             │            │            │
             ▼            ▼            ▼
      vtnode_hello   vtnode_myapp   vtnode_was
```

처음 사용하는 경우 `vtnode_vba_rt64`부터 분석할 필요는 없습니다.

`vtnode_hello` → `vtnode_build_myapp` → `vtnode_was`를 먼저 경험한 뒤  
Runtime 구조가 필요할 때 살펴보는 것을 권장합니다.

### [Open vtnode_vba_rt64](./vtnode_vba_rt64/)

---

# Application + Runtime

VTNode WAS for XLSX에서는 Application과 Runtime의 역할을 분리합니다.

```text
APPLICATION
─────────────────────

Excel UI
Excel Data
VBA Logic
URI Routes
HTML / CSS / JS


          │
          ▼


VBA RUNTIME
─────────────────────

vtnode_vba_rt64


          │
          ▼


VTNODE RUNTIME
─────────────────────

HTTP
WebSocket
Connectivity
```

Application은:

> **What should my application do?**

에 집중하고,

Runtime은 Application을 Network와 연결합니다.

---

# HTTP / WebSocket

VTNode WAS for XLSX는 HTTP와 WebSocket을 사용할 수 있습니다.

```text
                    VTNode WAS
                        │
              ┌─────────┴─────────┐
              │                   │
              ▼                   ▼
            HTTP              WebSocket
              │                   │
       Request/Response      Persistent
                             Connection
```

HTTP는 일반적인 Web Request / Response에 사용할 수 있고,  
WebSocket은 실시간 양방향 Communication에 사용할 수 있습니다.

---

# Excel UI + Web UI

VTNode WAS Application은 Excel UI와 Web UI를 함께 사용할 수 있습니다.

```text
                 Application
                      │
            ┌─────────┴─────────┐
            │                   │
            ▼                   ▼
        Excel UI             Web UI
            │                   │
     Worksheet/UserForm    HTML/CSS/JS
            │                   │
            └─────────┬─────────┘
                      │
                      ▼
                   VBA Logic
```

Excel에서 이미 사용하고 있는 UI와 Data를 유지하면서  
필요한 기능을 Browser로 확장할 수 있습니다.

---

# Excel Becomes a Network Node

일반적인 Excel Workbook은 Local Application입니다.

```text
Excel
  │
  ├─ Worksheet
  ├─ VBA
  ├─ Formula
  └─ Local Data
```

VTNode WAS를 사용하면 여기에 Connectivity가 추가됩니다.

```text
                  HTTP / WebSocket
                         │
                         ▼
Browser ◄────────── VTNode WAS
                         │
                         ▼
                       Excel
                         │
             ┌───────────┼───────────┐
             │           │           │
             ▼           ▼           ▼
         Worksheet      VBA      Local Data
```

기존 Excel을 없애는 것이 아니라  
**Excel이 이미 가지고 있는 UI, Logic, Data를 Network와 연결합니다.**

**Excel becomes a Network Node.**

---

# No Apache. No Node.js.

VTNode WAS for XLSX를 사용하기 위해  
별도의 Apache나 Node.js Application Server를 구성할 필요는 없습니다.

```text
No Apache
No Node.js
No separate Cloud Application Server
```

기본 구조는:

```text
Microsoft Excel
       +
      VBA
       +
 VTNode Runtime
       ↓
HTTP / WebSocket
```

입니다.

---

# Learning Path

각 Folder는 단계적으로 다른 경험을 제공합니다.

```text
┌───────────────────────────────┐
│ 01. vtnode_hello              │
│                               │
│ RUN IT                        │
│                               │
│ 먼저 실행해 본다.             │
└───────────────┬───────────────┘
                │
                ▼
┌───────────────────────────────┐
│ 02. vtnode_build_myapp        │
│                               │
│ BUILD IT                      │
│                               │
│ 직접 조립해 본다.             │
└───────────────┬───────────────┘
                │
                ▼
┌───────────────────────────────┐
│ 03. vtnode_was                │
│                               │
│ EXPLORE IT                    │
│                               │
│ Apps와 Network Activity를     │
│ 직접 관찰한다.                │
└───────────────┬───────────────┘
                │
                ▼
┌───────────────────────────────┐
│ 04. Your Application          │
│                               │
│ CHANGE IT                     │
│                               │
│ 자신의 Application으로       │
│ 확장한다.                     │
└───────────────────────────────┘
```

---

# What's Next?

Web Connectivity는 시작점입니다.

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

향후 VTNode Connectivity는  
Local Excel Data와 Cloud AI를 연결하는 방향으로 확장될 수 있습니다.

```text
              Cloud AI
                  │
                  │ MCP
                  ▼
              VTNode
                  │
                  ▼
           Excel / VBA
                  │
                  ▼
             Local Data
```

---

# Documentation

GitHub Repository에서는 Example과 Source Code를 제공합니다.

사용자가 직접 Application을 구성하는 상세 과정은  
VTNode 공식 Website의 User Manual에서 제공합니다.

### [VTNode WAS for XLSX User Manual](https://vtnode.io/docs/was-xlsx/manual)

### [Build My App Manual](https://vtnode.io/docs/was-xlsx/build-myapp)

```text
GitHub
   │
   ├─ Examples
   └─ Source Code


vtnode.io
   │
   ├─ Runtime Download
   ├─ Documentation
   └─ Build Manual
```

---

# Quick Links

### [Download Runtime](https://vtnode.io/download/runtime) | [Hello](./vtnode_hello/) | [Build My App](./vtnode_build_myapp/) | [Main Demo](./vtnode_was/) | [VBA Runtime](./vtnode_vba_rt64/) | [User Manual](https://vtnode.io/docs/was-xlsx/manual)

---

# VTNode WAS for XLSX (VBA)

> **Enable web service with VBA.**

```text
Excel / VBA
     │
     ▼
VTNode WAS
     │
     ▼
HTTP / WebSocket
     │
     ▼
Browser / Application
```

**Excel becomes a Network Node.**

**VTNode is an AI Glue Gun.**

> Humans should not have to learn the language of AI.  
> AI should adapt to the tools humans already use.

[VTNode Website](https://vtnode.io)

Copyright © 2026 VTNode