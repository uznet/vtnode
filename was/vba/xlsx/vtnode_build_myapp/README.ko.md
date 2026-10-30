[English](README.md) | [한글](README.ko.md)

# VTNode Build My App

> **Run it first. Then build it yourself.**

`vtnode_build_myapp`은 VTNode WAS for XLSX (VBA)가  
Excel Workbook 안에서 어떻게 구성되고 동작하는지를 직접 경험하기 위한 프로젝트입니다.

이미 완성된 `vtnode_hello.xlsm`을 실행해 보았다면,  
이번에는 동일한 Application을 직접 만들어 볼 차례입니다.

---

## From Hello to My App

`vtnode_hello`는 이미 완성된 Example입니다.

```text
vtnode_hello.xlsm
        ↓
      START
        ↓
Hello, VTNode!
```

`vtnode_build_myapp`은 다릅니다.

처음에는 UI만 준비된 Excel Workbook에서 시작합니다.

```text
vtnode_ui_only.xlsx
        │
        │  Build
        ▼
vtnode_myapp.xlsm
```

`vtnode_ui_only.xlsx`의 **MyApp** Sheet는  
`vtnode_hello`와 동일한 기본 UI를 사용합니다.

하지만 VTNode WAS 기능은 아직 연결되어 있지 않습니다.

제공되는 VBA Source와 구성요소를 직접 연결하면서  
하나의 동작하는 VTNode WAS Application을 완성하게 됩니다.

---

# What Will You Learn?

이 프로젝트의 목적은 VTNode Runtime API를 공부하는 것이 아닙니다.

또한 처음부터 Web Server를 프로그래밍하는 것도 아닙니다.

목적은 **Excel Workbook과 VTNode WAS가 어떻게 연결되는지를 직접 조립하면서 이해하는 것**입니다.

완료하고 나면 다음 구조가 자연스럽게 보이게 됩니다.

```text
                Browser
                   │
            HTTP / WebSocket
                   │
                   ▼
           VTNode WAS Engine
                   │
                   ▼
          vtnode_myapp.xlsm
                   │
          ┌────────┼────────┐
          │        │        │
         UI       VBA      www
```

---

## 1. An Excel Workbook Can Become a Web Application

Excel Workbook은 Worksheet와 VBA만 사용하는  
Local Application에 머물 필요가 없습니다.

VTNode WAS를 연결하면:

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
Browser
```

형태로 Browser와 통신할 수 있습니다.

**Excel becomes a Network Node.**

---

## 2. You Don't Have to Build Everything from Scratch

`vtnode_build_myapp`에는 필요한 구성요소가 준비되어 있습니다.

```text
vtnode_build_myapp/
│
├─ vtnode_ui_only.xlsx
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

사용자는 Web Server Engine이나 Runtime Library를  
처음부터 개발하지 않습니다.

준비된 구성요소가 어떤 역할을 하고  
어떻게 하나의 Excel Application으로 연결되는지를 경험합니다.

---

## 3. UI and Functionality Are Different Layers

처음 제공되는:

```text
vtnode_ui_only.xlsx
```

에는 UI가 이미 존재합니다.

Sheet Name:

```text
MyApp
```

화면은 `vtnode_hello`를 기본으로 하지만  
아직 실제 기능은 연결되어 있지 않습니다.

이것을 통해 Excel Application에서도:

```text
UI
 │
 ├── START
 ├── STOP
 ├── SHOW RUNTIME
 ├── STATUS
 ├── ACCESS URL
 └── QR CODE

        +

VBA Logic

        +

VTNode Runtime
```

처럼 **UI와 실행 기능을 분리해서 생각할 수 있다는 것**을 알 수 있습니다.

---

## 4. VBA Connects Excel to VTNode

제공되는 `src/`에는 VTNode WAS Application을 구성하기 위한  
VBA Module, Class, UserForm 등이 포함되어 있습니다.

```text
src/
│
├─ *.bas
├─ *.cls
└─ *.frm
```

이들을 Workbook에 구성하면서:

```text
Excel UI
    │
    ▼
VBA
    │
    ▼
VTNode Runtime
```

이라는 연결 관계를 이해할 수 있습니다.

중요한 것은 Runtime API 자체를 암기하는 것이 아니라  
**각 구성요소가 Application 안에서 어떤 역할을 하는지를 이해하는 것**입니다.

---

## 5. Web Contents Are Just Web Contents

Web Page는 `www/`에서 관리할 수 있습니다.

```text
www/
└─ default.html
```

HTML / CSS / JavaScript로 만든 Web UI를  
Excel/VBA Application과 연결할 수 있습니다.

즉:

```text
Excel / VBA                 Browser
     │                         ▲
     │                         │
     └──── VTNode WAS ─────────┘
```

구조가 됩니다.

기존 Web 기술과 Excel을 서로 대체하는 것이 아니라  
**연결해서 사용하는 것**이 핵심입니다.

---

# What You Will Build

시작할 때는:

```text
vtnode_ui_only.xlsx
```

입니다.

완료하면:

```text
vtnode_myapp.xlsm
```

이 됩니다.

그리고 Browser에서는:

```text
Hello, VTNode!

This page is served directly from Excel/VBA.

VTNode WAS for XLSX (VBA)

● Server is running
```

을 확인할 수 있습니다.

결과 화면은 `vtnode_hello`와 비슷합니다.

하지만 중요한 차이가 있습니다.

**이번에는 완성된 Example을 실행한 것이 아니라  
직접 구성해서 동작시킨 것입니다.**

---

# Run It → Build It → Change It

VTNode WAS for XLSX를 이해하는 가장 간단한 흐름입니다.

```text
vtnode_hello
    │
    │ Run It
    ▼
"How does it work?"
    │
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

`vtnode_build_myapp`의 최종 목적은  
`vtnode_myapp.xlsm` 자체가 아닙니다.

한 번 직접 만들어 본 후:

```text
Change the UI
Change the HTML
Use Excel Data
Add VBA Logic
Add HTTP Requests
Add WebSocket
```

처럼 자신의 Application으로 확장할 수 있다는 것을  
이해하는 것이 더 중요합니다.

---

# No Python Required

**No Python**은 이 Example을 구성하고 실행하는 데 Python 설치나 Python 지식이 필요 없다는 뜻입니다.
Python 연동을 금지한다는 의미가 아니며, 필요한 경우 Python과 연동할 수 있습니다.

### [Download VTNode Runtime](https://vtnode.io/downloads/)

---

# Build Manual

이 README는 상세한 제작 절차를 설명하지 않습니다.

`vtnode_ui_only.xlsx`에서 시작하여  
`vtnode_myapp.xlsm`을 완성하는 단계별 과정은  
별도의 **VTNode WAS for XLSX Build Manual**에서 제공할 예정입니다.

### [Download Build Manual](https://github.com/uznet/vtnode/blob/main/docs/quick-start/MYAPP_BUILD_GUIDE.ko.md)

[한글](https://github.com/uznet/vtnode/blob/main/docs/quick-start/MYAPP_BUILD_GUIDE.ko.md) | [English](https://github.com/uznet/vtnode/blob/main/docs/quick-start/MYAPP_BUILD_GUID.en.md)

```text
github.com/uznet/vtnode
    ↓
docs/quick-start/
    ↓
MYAPP_BUILD_GUIDE.ko.md
```

Manual은 현재 **준비중 ...**이며, 실제 Excel 화면과 VBA Editor를 따라가면서  
하나씩 구성하는 과정을 설명할 예정입니다.

---

# What's Next?

[VTNode WAS Main Demo](../vtnode_was/README.MD)에서
여러 Application과 실시간 HTTP / WebSocket 모니터링을 살펴볼 수 있습니다.

다음에 살펴볼 관련 주제로 **VTNode XEvent for XLSX (VBA)**도 있습니다.

---

# Start

### [Quick Start](https://github.com/uznet/vtnode/blob/main/docs/quick-start/QUICK_START.md) | [VTNode Hello](../vtnode_hello/README.md) | [Build My App](../vtnode_build_myapp/README.md) | [Documentation](https://vtnode.io/docs/docs_login)

먼저 `vtnode_hello.xlsm`을 실행해 보세요.

그리고 여기로 돌아오세요.

```text
Run It.
   ↓
Build It.
   ↓
Understand It.
   ↓
Change It.
```

**Your Excel. Your VBA. Your Web Application.**

---

**VTNode WAS for XLSX (VBA)**  
*Enable web service with VBA.*

