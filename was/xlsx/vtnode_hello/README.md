# VTNode Hello

> **Your first VTNode WAS for XLSX (VBA) example.**

`vtnode_hello`는 **VTNode WAS for XLSX (VBA)**를 처음 사용하는 사용자를 위한
가장 간단한 Quick Start Example입니다.

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

Excel에서 Web Service가 실제로 실행되는 것을 먼저 경험해 보세요.

---

## Quick Start

### 1. Install VTNode Runtime

`vtnode_hello.xlsm`을 실행하기 전에 **VTNode Runtime**이 설치되어 있어야 합니다.

### [Download VTNode Runtime](https://vtnode.io/download/runtime)

```text
Download VTNode Runtime
        ↓
Install
        ↓
Ready
```

---

### 2. Download `vtnode_hello`

이 Repository를 다운로드하거나 Clone합니다.

```text
vtnode_hello/
│
├─ vtnode_hello.xlsm
│
├─ src/
│  ├─ *.bas
│  ├─ *.cls
│  └─ *.frm
│
├─ audio/
│
└─ www/
   └─ default.html
```

처음에는 `src/`의 VBA Source Code를 이해할 필요가 없습니다.

먼저 완성된:

```text
vtnode_hello.xlsm
```

을 실행해 보세요.

---

### 3. Open `vtnode_hello.xlsm`

Microsoft Excel에서:

```text
vtnode_hello.xlsm
```

을 엽니다.

필요한 경우 Excel에서 VBA Macro 실행을 허용합니다.

기본 화면에서는 다음 기능을 확인할 수 있습니다.

```text
VTNode WAS for XLSX (VBA)
Created with AI.

[ START ]   [ STOP ]   [ SHOW RUNTIME ]

STATUS
Stopped

ACCESS URL
http://192.168.x.x:12345/

QR CODE
```

---

### 4. Click START

**START** 버튼을 클릭합니다.

VTNode WAS가 정상적으로 시작되면:

```text
STATUS
Running
```

으로 변경됩니다.

동시에 현재 PC에서 접속할 수 있는 **ACCESS URL**과
Smartphone에서 사용할 수 있는 **QR Code**를 확인할 수 있습니다.

Example:

```text
http://192.168.x.x:12345/
```

---

### 5. Open in Browser

Desktop에서는 **ACCESS URL**을 클릭합니다.

같은 Local Network에 연결된 Smartphone에서는
**QR Code**를 Scan할 수 있습니다.

```text
                         ┌──► Desktop Browser
                         │
Excel ──► VTNode WAS ────┤
                         │
                         └──► Smartphone
                              QR Code
```

Browser에 다음 페이지가 표시되면 성공입니다.

```text
Hello, VTNode!

This page is served directly from Excel/VBA.

──────────────────────────

VTNode WAS for XLSX (VBA)

● Server is running
```

**That's it.**

Excel에서 Web Service가 실행되고 있습니다.

---

# What Just Happened?

START 버튼을 클릭했을 때 기본적으로 다음 흐름이 만들어졌습니다.

```text
Browser
   │
   │ HTTP Request
   ▼
VTNode WAS Engine
   │
   ▼
Excel / VBA
   │
   ▼
VTNode WAS Engine
   │
   │ HTTP Response
   ▼
Browser
```

Browser는 HTTP Request를 보내고,

VTNode WAS Engine은 Request를 Excel/VBA Application으로 전달합니다.

Excel/VBA에서 처리된 결과는 다시 VTNode WAS Engine을 통해
HTTP Response로 Browser에 전달됩니다.

---

# Excel Becomes a Network Node

일반적인 Excel Workbook은 PC 내부에서 동작합니다.

```text
Excel
  │
  ├─ Worksheet
  ├─ VBA
  └─ Local Data
```

VTNode WAS를 사용하면 여기에 Network Connectivity가 추가됩니다.

```text
              HTTP / WebSocket
                     │
                     ▼
Browser ◄────── VTNode WAS
                     │
                     ▼
                   Excel
```

따라서 Excel/VBA Application을 Browser,
Smartphone 및 다른 Network Application과 연결할 수 있습니다.

**Excel becomes a Network Node.**

---

# What's Inside?

`vtnode_hello`에는 실행 파일뿐만 아니라
Application을 구성하는 Source와 Web Contents도 함께 포함되어 있습니다.

```text
vtnode_hello/
│
├─ vtnode_hello.xlsm       ← Ready-to-run Example
│
├─ src/                    ← VBA Source
│  ├─ *.bas
│  ├─ *.cls
│  └─ *.frm
│
├─ audio/                  ← Audio Resources
│
└─ www/                    ← Web Contents
   └─ default.html
```

`www/default.html`이 Browser에서 보이는
**Hello, VTNode!** 페이지입니다.

HTML / CSS / JavaScript를 사용하여 Web Contents를 구성할 수 있습니다.

---

# No Apache. No Node.js.

이 Example을 실행하기 위해 별도의 Apache나 Node.js
Application Server를 구성할 필요가 없습니다.

```text
No Apache
No Node.js
No separate Cloud Application Server
```

필요한 것은:

```text
Microsoft Excel
      +
VTNode Runtime
```

입니다.

---

# Don't Study It Yet

처음부터 VBA Source나 Runtime 내부 구조를
모두 이해하려고 할 필요는 없습니다.

`vtnode_hello`의 첫 번째 목적은 단순합니다.

```text
Open
  ↓
START
  ↓
Browser
  ↓
Hello, VTNode!
```

**먼저 실행해 보세요.**

동작하는 것을 확인한 다음 직접 만들어 보는 단계로 넘어갈 수 있습니다.

---

# Next: Build It Yourself

`vtnode_hello.xlsm`이 어떻게 만들어졌는지 궁금하다면
다음 단계는 **`vtnode_build_myapp`**입니다.

### [Build My App](https://github.com/vtnode/vtnode_build_myapp)

`vtnode_build_myapp`에서는 완성된 `.xlsm`에서 시작하지 않습니다.

`vtnode_hello`와 동일한 기본 UI를 가진:

```text
vtnode_ui_only.xlsx
```

에서 시작합니다.

```text
vtnode_ui_only.xlsx
        │
        │ Assemble
        ▼
vtnode_myapp.xlsm
```

제공되는 VBA Module, Class, UserForm 등의 구성요소를 이용하여
직접 동작하는 VTNode WAS Application으로 만들어 봅니다.

```text
vtnode_hello
    │
    │ Run It
    ▼
vtnode_build_myapp
    │
    │ Build It
    ▼
vtnode_myapp.xlsm
    │
    │ Change It
    ▼
Your Application
```

---

# Build Manual

직접 `.xlsm`을 구성하는 단계별 방법은
GitHub README가 아닌 별도의 **Build Manual**에서 제공합니다.

### [Download Build Manual](https://vtnode.io/docs/was-xlsx/build-myapp)

```text
vtnode.io
    ↓
VTNode WAS for XLSX
    ↓
Build My App
    ↓
Build Manual
```

---

# Where to Go Next

```text
Hello
  │
  ▼
Build My App
  │
  ▼
Table Order
  │
  ▼
WebSocket
  │
  ▼
AI Connectivity
  │
  ├─ MCP Server
  │
  └─ MCP + SSE
  │
  ▼
Voice / Phone
```

`vtnode_hello`는 그 시작점입니다.

---

## Start Here

### [Download Runtime](https://vtnode.io/download/runtime) | [Build My App](https://github.com/vtnode/vtnode_build_myapp) | [Build Manual](https://vtnode.io/docs/was-xlsx/build-myapp)

---

# VTNode WAS for XLSX (VBA)

> **Enable web service with VBA.**

**VTNode is an AI Glue Gun.**

**Excel becomes a Network Node.**

[VTNode Website](https://vtnode.io)