# VTNode WAS for XLSX (VBA)

> **Enable web service with VBA.**

VTNode WAS for XLSX (VBA)는 Microsoft Excel/VBA에서  
HTTP / WebSocket 기반 Web Service를 실행할 수 있도록 하는 경량 Web Application Server입니다.

별도의 Apache, Node.js 또는 Cloud Application Server 없이  
Excel/VBA를 Web Application과 직접 연결할 수 있습니다.

### [Download Runtime](https://vtnode.io/download/runtime) | [Quick Start](https://github.com/vtnode/vtnode_hello) | [User Manual](https://vtnode.io/docs/was-xlsx/manual)

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

**Excel becomes a Network Node.**

---

# Quick Start

처음 사용하는 경우 `vtnode_hello` Example부터 시작하세요.

```text
Download VTNode Runtime
        ↓
Install
        ↓
Download vtnode_hello
        ↓
Open vtnode_hello.xlsm
        ↓
START
        ↓
Open ACCESS URL
        ↓
Hello, VTNode!
```

---

## Step 1. Download & Install VTNode Runtime

먼저 **VTNode Runtime**을 설치합니다.

VTNode Runtime에는 Excel/VBA에서 사용하는  
HTTP / WebSocket Engine과 실행 환경이 포함되어 있습니다.

### [Download VTNode Runtime](https://vtnode.io/download/runtime)

```text
vtnode.io
    ↓
Download VTNode Runtime
    ↓
Run Installer
    ↓
VTNode Runtime Installed
```

설치가 완료되면 VTNode WAS for XLSX Example을 실행할 준비가 끝납니다.

---

## Step 2. Download `vtnode_hello`

GitHub에서 `vtnode_hello`를 다운로드합니다.

### [Download / View vtnode_hello](https://github.com/vtnode/vtnode_hello)

`vtnode_hello`는 VTNode WAS for XLSX의 가장 작은  
**Hello / Quick Start Example**입니다.

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
│  ├─ audio_tick.wav
│  └─ audio_was_started.wav
│
└─ www/
   └─ default.html
```

처음에는 Source Code를 이해할 필요가 없습니다.

먼저 다음 파일을 실행해 보세요.

```text
vtnode_hello.xlsm
```

---

## Step 3. Open `vtnode_hello.xlsm`

Microsoft Excel에서:

```text
vtnode_hello.xlsm
```

을 실행합니다.

필요한 경우 Excel에서 VBA Macro 실행을 허용합니다.

화면에는 다음과 같은 기본 기능이 표시됩니다.

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

## Step 4. START

**START** 버튼을 클릭합니다.

VTNode WAS가 정상적으로 시작되면:

```text
STATUS
Running
```

으로 변경되고 **ACCESS URL**과 **QR Code**를 사용할 수 있습니다.

Example:

```text
http://192.168.x.x:12345/
```

---

## Step 5. Open in Browser

Desktop Browser에서는 **ACCESS URL**을 클릭합니다.

같은 Network에 연결된 Smartphone에서는  
화면의 **QR Code**를 Scan할 수 있습니다.

```text
                         ┌──► Desktop Browser
                         │
Excel ──► VTNode WAS ────┤
                         │
                         └──► Smartphone
                              QR Code
```

Browser에서 다음 페이지가 표시되면 성공입니다.

```text
Hello, VTNode!

This page is served directly from Excel/VBA.

──────────────────────────

VTNode WAS for XLSX (VBA)

● Server is running
```

**Excel에서 Web Service가 실행되고 있습니다.**

---

# How It Works

VTNode WAS for XLSX의 기본 구조는 단순합니다.

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
   │ Execute
   ▼
VTNode WAS Engine
   │
   │ HTTP Response
   ▼
Browser
```

HTTP Request를 Excel/VBA가 처리하고  
VTNode WAS Engine을 통해 HTTP Response를 Browser로 전달합니다.

VBA Application 관점에서는 더 단순하게 볼 수 있습니다.

```text
Request Event
     │
     ▼
Excel / VBA Execute
     │
     ▼
Response Event
```

Excel 사용자는 Web Server 내부 구현보다,

```text
Request가 왔을 때
        ↓
URI / Content 확인
        ↓
Excel / VBA Execute
        ↓
Response 생성
```

에 집중할 수 있습니다.

---

# HTTP / WebSocket

VTNode WAS for XLSX는 HTTP뿐만 아니라  
WebSocket 기반 Application도 구성할 수 있습니다.

```text
                 ┌── HTTP ────────► Request / Response
Browser ─────────┤
                 └── WebSocket ───► Real-time Communication
                         │
                         ▼
                    VTNode WAS
                         │
                         ▼
                    Excel / VBA
```

이를 이용하여 Excel을 단순한 문서가 아니라  
Network Application의 실행 지점으로 사용할 수 있습니다.

---

# Build Your Own XLSM

`vtnode_hello.xlsm`을 그대로 실행할 수도 있고,  
사용자가 새로운 Excel Workbook에서 직접  
VTNode WAS Application을 만들 수도 있습니다.

```text
New Excel Workbook
        ↓
Save as *.xlsm
        ↓
VTNode VBA Runtime
        ↓
Your VBA Code
        ↓
VTNode WAS Engine
        ↓
HTTP / WebSocket
        ↓
Your Web Application
```

직접 `.xlsm`을 만드는 상세 과정은  
별도의 **VTNode WAS for XLSX User Manual**에서 설명합니다.

### [Download VTNode WAS for XLSX User Manual](https://vtnode.io/docs/was-xlsx/manual)

User Manual에서는 다음 과정을 단계별로 설명합니다.

- VTNode Runtime 설치
- 새로운 Excel Workbook 생성
- `.xlsm`으로 저장
- VTNode VBA Runtime 연결
- VBA Module / Class / UserForm 구성
- `www/` Directory 구성
- `default.html` 작성
- HTTP Request / Response 처리
- WebSocket 사용
- Server START / STOP
- Desktop Browser Test
- Smartphone / QR Code Test
- Troubleshooting

User Manual은 **vtnode.io**에서 제공합니다.

```text
vtnode.io
    ↓
VTNode WAS for XLSX
    ↓
User Manual
    ↓
Email Verification
    ↓
Download
```

> GitHub Repository, Example 및 Source Code는  
> Email Verification 없이 사용할 수 있습니다.

---

# Repository Structure

VTNode WAS for XLSX 관련 Repository는 역할에 따라 분리되어 있습니다.

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

| Repository | Role |
|---|---|
| [`vtnode_hello`](https://github.com/vtnode/vtnode_hello) | Hello / Quick Start |
| [`vtnode_was`](https://github.com/vtnode/vtnode_was) | Main Demo / Examples |
| [`vtnode_vba_rt64`](https://github.com/vtnode/vtnode_vba_rt64) | Common VBA Runtime |

---

# `vtnode_hello`

가장 작은 VTNode WAS for XLSX Example입니다.

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

VTNode WAS를 처음 사용하는 경우 여기에서 시작하세요.

### [Start with vtnode_hello](https://github.com/vtnode/vtnode_hello)

---

# `vtnode_was`

VTNode WAS for XLSX의 주요 기능과 Application Example을 제공합니다.

Example은 단계적으로 확장됩니다.

```text
Hello
  ↓
HTTP
  ↓
WebSocket
  ↓
Table Order
  ↓
Your Application
```

단순한 Web Page부터 Excel 데이터를 사용하는  
실제 Web Application까지 확장할 수 있습니다.

### [View vtnode_was](https://github.com/vtnode/vtnode_was)

---

# `vtnode_vba_rt64`

Excel/VBA Application과 VTNode Runtime을 연결하는  
공통 VBA Runtime입니다.

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
HTTP / WebSocket Engine
```

Source Code도 함께 제공하여 필요한 경우  
VBA 환경에서 직접 확인하고 구성할 수 있습니다.

### [View vtnode_vba_rt64](https://github.com/vtnode/vtnode_vba_rt64)

---

# Web Contents

VTNode WAS는 Web Contents를 `www/` Directory에서 서비스할 수 있습니다.

```text
vtnode_hello/
│
├─ vtnode_hello.xlsm
└─ www/
   └─ default.html
```

Example:

```html
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <title>Hello, VTNode!</title>
</head>

<body>

    <h1>Hello, VTNode!</h1>

    <p>
        This page is served directly from Excel/VBA.
    </p>

</body>

</html>
```

HTML / CSS / JavaScript를 사용하여  
자유롭게 Web UI를 구성할 수 있습니다.

---

# From Excel to Web

기존 Excel Application은 일반적으로 PC 내부에서 실행됩니다.

```text
Excel
  │
  ├─ Worksheet
  ├─ VBA
  ├─ Formula
  └─ Local Data
```

VTNode WAS를 사용하면 여기에 Network Interface가 추가됩니다.

```text
              HTTP / WebSocket
                     │
                     ▼
Browser ◄────── VTNode WAS
                     │
                     ▼
                   Excel
                     │
          ┌──────────┼──────────┐
          │          │          │
      Worksheet     VBA       Local Data
```

기존 Excel Application을 모두 버리고  
새로운 Web System으로 다시 만드는 것이 아니라,

**Excel이 이미 가지고 있는 기능과 데이터를 Web에 연결합니다.**

---

# No Apache. No Node.js.

VTNode WAS for XLSX의 목적은  
Excel 사용자가 Web Server 기술 전체를 먼저 배워야만  
Web Application을 만들 수 있게 하는 것이 아닙니다.

```text
No Apache
No Node.js
No separate Cloud Application Server
```

VTNode Runtime과 Excel/VBA를 이용하여  
Excel에서 직접 Web Service를 실행할 수 있습니다.

---

# Created with AI

VTNode WAS for XLSX Example과 UI는  
AI를 개발 도구로 활용하여 만들어지고 있습니다.

하지만 Runtime, HTTP, WebSocket, Networking과 같은  
기반 기술은 VTNode Engine에서 처리합니다.

```text
AI
 │
 ├─ HTML / CSS / JavaScript
 ├─ VBA
 ├─ UI
 └─ Application Logic
        │
        ▼
   VTNode Runtime
        │
        ▼
   Connectivity
```

AI가 Application 개발을 돕고,  
VTNode는 그 Application을 실제 System과 연결합니다.

---

# Requirements

- Microsoft Windows
- Microsoft Excel with VBA support
- VTNode Runtime
- Web Browser

Smartphone에서 테스트하려면 PC와 Smartphone이  
동일한 Local Network에 연결되어 있어야 합니다.

---

# Start Here

처음 방문했다면 아래 순서만 따라 하면 됩니다.

```text
1. Download VTNode Runtime
           ↓
2. Install
           ↓
3. Download vtnode_hello
           ↓
4. Open vtnode_hello.xlsm
           ↓
5. Click START
           ↓
6. Open ACCESS URL
           ↓
7. Hello, VTNode!
```

### [Download Runtime](https://vtnode.io/download/runtime) | [Quick Start](https://github.com/vtnode/vtnode_hello) | [User Manual](https://vtnode.io/docs/was-xlsx/manual)

---

# VTNode

**VTNode is an AI Glue Gun.**

> Humans should not have to learn the language of AI.  
> AI should adapt to the tools humans already use.

**Excel becomes a Network Node.**

---

Website: [vtnode.io](https://vtnode.io)

Copyright © 2026 VTNode