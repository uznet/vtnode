[English](README.md) | [한글](README.ko.md)

# VTNode VBA Runtime 64

> **Common VBA Runtime for VTNode WAS for XLSX (VBA)**

`vtnode_vba_rt64.xlsm` is a shared 64-bit Excel/VBA runtime for
**VTNode WAS for XLSX (VBA)** applications.

It provides common runtime features and a VBA connection layer
between application workbooks and the native VTNode Runtime.

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

`vtnode_vba_rt64.xlsm` is not a typical Excel application example.

It is a **VBA runtime workbook** that can be shared by
VTNode WAS applications such as the following:

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

Each application focuses on its own UI, data, and VBA logic.

`vtnode_vba_rt64` handles the common features for connecting to
and working with VTNode Runtime.

---

# Why a Separate VBA Runtime?

You do not need to implement the same runtime features
in every VTNode WAS application.

Separating the application from the runtime makes their roles clear.

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

The overall structure is as follows:

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

In VTNode WAS for XLSX,
**it is important to distinguish the application from the runtime.**

The application answers this question:

```text
"What should my application do?"
```

The runtime answers this question:

```text
"How does my application connect to VTNode?"
```

This produces the following structure:

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

---

# Runtime Console

`vtnode_vba_rt64.xlsm` includes a Runtime Console for viewing
and managing VTNode Runtime.

An application can display the Runtime Console through:

```text
[ SHOW RUNTIME ]
```

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

Separating the application UI from the runtime management UI
lets users focus on their Excel application.

---

# Used by VTNode WAS Applications

`vtnode_vba_rt64` can serve as a shared runtime for multiple
VTNode WAS for XLSX applications.

```text
                   vtnode_vba_rt64
                          │
            ┌─────────────┼─────────────┐
            │             │             │
            ▼             ▼             ▼
     vtnode_hello     vtnode_myapp   vtnode_was
```

Each application has a different purpose.

| Application | Purpose |
|---|---|
| `vtnode_hello` | The smallest Quick Start example |
| `vtnode_myapp` | An application you assemble yourself |
| `vtnode_was` | Main Demo / Apps + Runtime + Live Monitor |

They can all use `vtnode_vba_rt64` as their common layer
for connecting to VTNode Runtime.

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

A ready-to-use 64-bit VBA runtime workbook.

### `src/`

The VBA source code that makes up the runtime workbook.

```text
*.bas    VBA Modules
*.cls    VBA Classes
*.frm    VBA UserForms
```

The source code is included so you can inspect the runtime's
components in the VBA Editor when needed.

---

# Ready-to-Use and Source

You can work with the runtime in two ways:

```text
Ready-to-Use
     │
     ▼
vtnode_vba_rt64.xlsm
```

Or explore its source:

```text
Source
   │
   ▼
src/
├─ *.bas
├─ *.cls
└─ *.frm
```

If you are getting started, you do not need to analyze the runtime source first.

Start with the ready-to-use `vtnode_vba_rt64.xlsm` to understand
the relationship between the application and the runtime.

---

# 64-bit Excel

The `rt64` in `vtnode_vba_rt64` refers to
a 64-bit runtime environment.

```text
vtnode_vba_rt64
           │
           └── rt64
                 │
                 ▼
           64-bit Excel/VBA
```

This repository targets **64-bit Microsoft Excel/VBA environments**.

---

# VTNode Runtime Required

`vtnode_vba_rt64.xlsm` is not itself a native HTTP/WebSocket engine.

VTNode Runtime must be installed.

### [Download VTNode Runtime](https://vtnode.io/downloads/)

**No Python** means that you do not need to install Python or have any Python knowledge
to use this VBA runtime. It does not prohibit integration with Python;
you can integrate with Python when needed.

The overall relationship is as follows:

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

If you are new to VTNode WAS for XLSX,
we recommend exploring `vtnode_vba_rt64` later.

First, run the smallest example.

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

### [Quick Start](https://github.com/uznet/vtnode/blob/main/docs/quick-start/QUICK_START.md)

---

# Build My App

Use `vtnode_build_myapp` to get hands-on experience
assembling a VTNode WAS application.

### [Build My App](https://github.com/uznet/vtnode/tree/main/was/xlsx/vtnode_build_myapp)

Start with a workbook that contains only the UI:

```text
vtnode_ui_only.xlsx
```

Assemble the provided VBA components to create:

```text
vtnode_myapp.xlsm
```

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

The goal is to experience the following relationship firsthand,
without having to study the native runtime API:

```text
Excel Application
        ↓
VBA Runtime
        ↓
VTNode Runtime
```

---

# Main Demo

Use `vtnode_was.xlsm` to observe HTTP requests and responses
and WebSocket sessions in action.

### [VTNode WAS Main Demo](https://github.com/uznet/vtnode/tree/main/was/xlsx/vtnode_was)

```text
vtnode_was.xlsm

Apps
 +
Runtime
 +
Live Monitor
```

In the Main Demo, the following views show network activity
between the browser and Excel in real time:

```text
REQUEST → RESPONSE ACTIVITY

WEBSOCKET CONNECTION SESSIONS
```

---

# Do I Need to Learn the Runtime API?

You do not need to learn it to get started.

The basic workflow for VTNode WAS for XLSX is:

```text
Run It
   ↓
Build It
   ↓
Explore It
   ↓
Change It
```

The purpose of `vtnode_vba_rt64` is to **provide a shared runtime environment
that lets Excel/VBA applications use VTNode Runtime**, without requiring
users to learn the native runtime API from the start.

You can refer to the detailed runtime documentation and source code
later, when needed.

---

# Documentation

User documentation for VTNode WAS for XLSX
is available on the official website.

### [VTNode WAS for XLSX Documentation](https://vtnode.io/docs/docs_login)

Step-by-step instructions for assembling your own application
are available in the Build Manual:

### [Download Build Manual](https://github.com/uznet/vtnode/blob/main/docs/quick-start/BUILD_MYAPP.md)

---

# Related Projects

### [VTNode Hello](https://github.com/uznet/vtnode/tree/main/was/xlsx/vtnode_hello)

The smallest Quick Start example.

### [VTNode Build My App](https://github.com/uznet/vtnode/tree/main/was/xlsx/vtnode_build_myapp)

A hands-on project where you assemble the provided UI and VBA source code.

### [VTNode WAS](https://github.com/uznet/vtnode/tree/main/was/xlsx/vtnode_was)

The Main Demo workbook for observing HTTP requests and responses
and WebSocket sessions in real time.

---

# Start Here

### [Download Runtime](https://vtnode.io/downloads/) | [Quick Start](https://github.com/uznet/vtnode/blob/main/docs/quick-start/QUICK_START.md) | [Build My App](https://github.com/uznet/vtnode/tree/main/was/xlsx/vtnode_build_myapp) | [Main Demo](https://github.com/uznet/vtnode/tree/main/was/xlsx/vtnode_was)

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
