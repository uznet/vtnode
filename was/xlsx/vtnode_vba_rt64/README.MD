# VTNode VBA Runtime 64

> **Common VBA Runtime for VTNode WAS for XLSX (VBA)**

`vtnode_vba_rt64.xlsm`은 **VTNode WAS for XLSX (VBA)** Application에서
공통으로 사용하는 64-bit Excel/VBA Runtime입니다.

Application Workbook과 Native VTNode Runtime 사이에서
공통 Runtime 기능과 VBA 연결 계층을 제공합니다.

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

---

# What is `vtnode_vba_rt64.xlsm`?

`vtnode_vba_rt64.xlsm`은 일반적인 Excel Application Example이 아닙니다.

다음과 같은 VTNode WAS Application에서 공통으로 사용할 수 있는
**VBA Runtime Workbook**입니다.

```text
vtnode_hello.xlsm
        │
        │
vtnode_myapp.xlsm
        │
        │
vtnode_was.xlsm
        │
        ▼
vtnode_vba_rt64.xlsm
        │
        ▼
VTNode Runtime
```

Application은 자신의 UI, Data, VBA Logic에 집중하고,

공통적인 VTNode Runtime 연결과 실행 기능은
`vtnode_vba_rt64`에서 담당합니다.

---

# Why a Separate VBA Runtime?

VTNode WAS Application마다 동일한 Runtime 기능을
반복해서 구성할 필요는 없습니다.

Application과 Runtime을 분리하면 역할이 명확해집니다.

```text
APPLICATION
────────────────────────

Worksheet
UserForm
Application UI
Business Logic
Excel Data
Application Routes


RUNTIME
────────────────────────

VTNode Initialization
Runtime State
HTTP / WebSocket Connection
Request / Response Events
Runtime Console
Common VBA Components
```

전체 구조는 다음과 같습니다.

```text
┌──────────────────────────────┐
│      Excel Application       │
│                              │
│   UI / Data / VBA Logic      │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│      vtnode_vba_rt64         │
│                              │
│      Common VBA Runtime      │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐
│        VTNode Runtime        │
│                              │
│     HTTP / WebSocket Engine  │
└──────────────────────────────┘
```

---

# Application vs Runtime

VTNode WAS for XLSX에서는
**Application과 Runtime을 분리해서 생각하는 것이 중요합니다.**

Application은:

```text
"What should my application do?"
```

를 담당합니다.

Runtime은:

```text
"How does my application connect to VTNode?"
```

를 담당합니다.

따라서:

```text
Application Logic
       │
       ▼
Excel / VBA
       │
       ▼
VBA Runtime
       │
       ▼
VTNode Runtime
       │
       ▼
Network
```

이라는 구조가 만들어집니다.

---

# Runtime Console

`vtnode_vba_rt64.xlsm`에는 VTNode Runtime을 확인하고
관리하기 위한 Runtime Console이 포함됩니다.

Application에서는:

```text
[ SHOW RUNTIME ]
```

을 통해 Runtime Console을 표시할 수 있습니다.

```text
Excel Application
       │
       │ SHOW RUNTIME
       ▼
┌─────────────────────────┐
│ VTNode Runtime Console  │
│                         │
│ Runtime Status          │
│ Server Status           │
│ Runtime Information     │
│ Runtime Log             │
└─────────────────────────┘
```

Application UI와 Runtime 관리 UI를 분리함으로써
사용자는 자신의 Excel Application 화면에 집중할 수 있습니다.

---

# Used by VTNode WAS Applications

`vtnode_vba_rt64`은 여러 VTNode WAS for XLSX Application에서
공통 Runtime으로 사용할 수 있습니다.

```text
                   vtnode_vba_rt64
                          │
            ┌─────────────┼─────────────┐
            │             │             │
            ▼             ▼             ▼
     vtnode_hello     vtnode_myapp   vtnode_was
```

각 Application의 목적은 서로 다릅니다.

| Application | Purpose |
|---|---|
| `vtnode_hello` | 가장 작은 Quick Start Example |
| `vtnode_myapp` | 사용자가 직접 구성한 Application |
| `vtnode_was` | Main Demo / Apps + Runtime + Live Monitor |

하지만 VTNode Runtime과 연결하는 공통 계층은
`vtnode_vba_rt64`을 사용할 수 있습니다.

---

# Repository Structure

```text
vtnode_vba_rt64/
│
├─ vtnode_vba_rt64.xlsm
│
└─ src/
   ├─ *.bas
   ├─ *.cls
   └─ *.frm
```

### `vtnode_vba_rt64.xlsm`

사용할 수 있도록 구성된 64-bit VBA Runtime Workbook입니다.

### `src/`

Runtime Workbook을 구성하는 VBA Source Code입니다.

```text
*.bas    VBA Modules
*.cls    VBA Classes
*.frm    VBA UserForms
```

Source Code를 함께 제공하므로
필요한 경우 VBA Editor에서 Runtime의 구성을 직접 확인할 수 있습니다.

---

# Ready-to-Use and Source

두 가지 관점에서 사용할 수 있습니다.

```text
Ready-to-Use
     │
     ▼
vtnode_vba_rt64.xlsm
```

또는:

```text
Source
   │
   ▼
src/
├─ *.bas
├─ *.cls
└─ *.frm
```

처음 사용하는 경우 Runtime Source를 먼저 분석할 필요는 없습니다.

완성된 `vtnode_vba_rt64.xlsm`을 통해
Application과 Runtime의 관계를 먼저 이해하는 것을 권장합니다.

---

# 64-bit Excel

`vtnode_vba_rt64`의 `rt64`는
64-bit Runtime 환경을 의미합니다.

```text
vtnode_vba_rt64
           │
           └── rt64
                 │
                 ▼
           64-bit Excel/VBA
```

이 Repository는 **64-bit Microsoft Excel/VBA 환경**을 대상으로 합니다.

---

# VTNode Runtime Required

`vtnode_vba_rt64.xlsm` 자체가 Native HTTP/WebSocket Engine은 아닙니다.

VTNode Runtime이 설치되어 있어야 합니다.

### [Download VTNode Runtime](https://vtnode.io/download/runtime)

전체 관계는 다음과 같습니다.

```text
Microsoft Excel
       │
       ▼
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
       │
       ▼
Browser / Client
```

---

# Where Should I Start?

VTNode WAS for XLSX를 처음 사용하는 경우
`vtnode_vba_rt64`부터 시작하지 않는 것을 권장합니다.

먼저 가장 작은 Example을 실행해 보세요.

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

04. vtnode_vba_rt64
        │
        │ Understand the Runtime
        ▼

05. Your Application
```

### [Quick Start](https://github.com/vtnode/vtnode_hello)

---

# Build My App

VTNode WAS Application이 어떻게 구성되는지
직접 경험하고 싶다면 `vtnode_build_myapp`을 사용하세요.

### [Build My App](https://github.com/vtnode/vtnode_build_myapp)

여기서는 UI만 준비된:

```text
vtnode_ui_only.xlsx
```

에서 시작합니다.

제공되는 VBA 구성요소를 조립하여 최종적으로:

```text
vtnode_myapp.xlsm
```

을 완성합니다.

```text
vtnode_ui_only.xlsx
        │
        │ Assemble
        ▼
VBA Components
        │
        ▼
vtnode_myapp.xlsm
```

이 과정의 목적은 Native Runtime API를 공부하는 것이 아니라,

```text
Excel Application
        ↓
VBA Runtime
        ↓
VTNode Runtime
```

의 관계를 직접 경험하는 것입니다.

---

# Main Demo

HTTP Request / Response와 WebSocket Session을
실제로 관찰하려면 `vtnode_was.xlsm`을 사용할 수 있습니다.

### [VTNode WAS Main Demo](https://github.com/vtnode/vtnode_was)

```text
vtnode_was.xlsm

Apps
 +
Runtime
 +
Live Monitor
```

Main Demo에서는:

```text
REQUEST → RESPONSE ACTIVITY

WEBSOCKET CONNECTION SESSIONS
```

을 통해 Browser와 Excel 사이의 Network Activity를
실시간으로 확인할 수 있습니다.

---

# Do I Need to Learn the Runtime API?

처음에는 필요하지 않습니다.

VTNode WAS for XLSX의 기본 사용 흐름은:

```text
Run It
   ↓
Build It
   ↓
Explore It
   ↓
Change It
```

입니다.

`vtnode_vba_rt64`의 목적은 사용자가 Native Runtime API를
처음부터 학습하게 만드는 것이 아니라,

**Excel/VBA Application이 VTNode Runtime을 사용할 수 있도록
공통 Runtime 환경을 제공하는 것**입니다.

필요한 경우 이후에 더 상세한 Runtime Documentation과
Source Code를 참고할 수 있습니다.

---

# Documentation

VTNode WAS for XLSX의 사용자용 Documentation은
공식 Website에서 제공합니다.

### [VTNode WAS for XLSX Documentation](https://vtnode.io/docs/was-xlsx)

직접 Application을 구성하는 단계별 과정은:

### [Download Build Manual](https://vtnode.io/docs/was-xlsx/build-myapp)

에서 확인할 수 있습니다.

---

# Related Projects

### [VTNode Hello](https://github.com/vtnode/vtnode_hello)

가장 작은 Quick Start Example.

### [VTNode Build My App](https://github.com/vtnode/vtnode_build_myapp)

제공된 UI와 VBA Source를 직접 조립하는 Hands-on Project.

### [VTNode WAS](https://github.com/vtnode/vtnode_was)

HTTP Request / Response와 WebSocket Session을
실시간으로 확인할 수 있는 Main Demo Workbook.

---

# Start Here

### [Download Runtime](https://vtnode.io/download/runtime) | [Quick Start](https://github.com/vtnode/vtnode_hello) | [Build My App](https://github.com/vtnode/vtnode_build_myapp) | [Main Demo](https://github.com/vtnode/vtnode_was)

---

# VTNode WAS for XLSX (VBA)

> **Enable web service with VBA.**

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

**Excel becomes a Network Node.**

[VTNode Website](https://vtnode.io)

Copyright © 2026 VTNode