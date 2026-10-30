[English](README.md) | [한글](README.ko.md)

[VBA project](https://github.com/uznet/vtnode/tree/main/was/vba) | [Current directory](https://github.com/uznet/vtnode/tree/main/was/vba/xlsx/vtnode_hello)

# VTNode Hello

## Direct VBA import and build

The agreed layout is `was/vba/common/src` (Office shared VBA), `was/vba/common/xlsx` (Excel only), `was/vba/common/pptx` (PowerPoint only), and `was/vba/xlsx/` (Excel apps/examples). An application's `src/` holds app-specific code; it does not replace the shared sources.

1. Back up the target file. Save the Excel UI `.xlsx` as `.xlsm`; use `.pptm` for PowerPoint or `.docm` for Word to retain VBA.
2. Open the target file's VBA editor. Use **File > Import File** to import the required `.bas` and `.cls` files and their dependencies from [common/src](../../common/src/) directly into that file's VBA project.
3. Add host-specific sources from [common/xlsx](../../common/xlsx/) for Excel or [common/pptx](../../common/pptx/) for PowerPoint. Keep Excel-only code out of other Office projects. For Word, use Office shared sources and implement and verify the required host integration separately.
4. Import app-specific sources from the application's `src/`. Back up and replace components with matching names to avoid duplicates. Keep each UserForm's `.frm` and `.frx` together and import the `.frm`. Apply `ThisWorkbook`, worksheet, or document event code to the appropriate document modules.
5. Do not open or reference an external runtime workbook. Remove any existing external runtime workbook project reference under **Tools > References**, then check required library references. Check that UI and macro bindings call procedures inside the target file.
6. Run **Debug > Compile VBAProject** and save. Prepare the native VTNode Runtime with the same bitness as Office. Preserve the relative placement of `www`, `audio`, and configuration files, then verify start/stop, HTTP/WebSocket, and Runtime UI behavior. Classification as Office shared source does not establish tested compatibility with every Office host.


> **Your first VTNode WAS for XLSX (VBA) example.**

`vtnode_hello` is the simplest Quick Start example for users getting started
with **VTNode WAS for XLSX (VBA)**.

It has one goal:

```text
Excel
  ↓
START
  ↓
Browser
  ↓
Hello, VTNode!
```

Start by seeing a web service running directly from Excel.

---

## Quick Start

### 1. Install VTNode Runtime

You must install **VTNode Runtime** before running `vtnode_hello.xlsm`.

### [Download VTNode Runtime](https://vtnode.io/downloads/)

```text
Download VTNode Runtime
        ↓
Install
        ↓
Ready
```

---

### 2. Download `vtnode_hello`

Download or clone this repository.

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

You do not need to understand the VBA source code in `src/` to get started.

First, run the ready-to-use workbook:

```text
vtnode_hello.xlsm
```

---

### 3. Open `vtnode_hello.xlsm`

Open the following file in Microsoft Excel:

```text
vtnode_hello.xlsm
```

If prompted, enable VBA macros in Excel.

The main screen includes the following controls and information:

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

Click the **START** button.

When VTNode WAS starts successfully, the status changes to:

```text
STATUS
Running
```

You will also see an **ACCESS URL** for accessing the service from your PC
and a **QR Code** for accessing it from a smartphone.

Example:

```text
http://192.168.x.x:12345/
```

---

### 5. Open in Browser

On your desktop, click the **ACCESS URL**.

On a smartphone connected to the same local network,
scan the **QR Code**.

```text
                         ┌──► Desktop Browser
                         │
Excel ──► VTNode WAS ────┤
                         │
                         └──► Smartphone
                              QR Code
```

If the following page appears in your browser, it worked:

```text
Hello, VTNode!

This page is served directly from Excel/VBA.

──────────────────────────

VTNode WAS for XLSX (VBA)

● Server is running
```

**That's it.**

A web service is now running from Excel.

---

# What Just Happened?

Clicking START set up the following basic flow:

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

The browser sends an HTTP request.

The VTNode WAS Engine forwards the request to the Excel/VBA application.

The result produced by Excel/VBA passes back through the VTNode WAS Engine
and is returned to the browser as an HTTP response.

---

# Excel Becomes a Network Node

A typical Excel workbook runs locally on your PC.

```text
Excel
  │
  ├─ Worksheet
  ├─ VBA
  └─ Local Data
```

VTNode WAS adds network connectivity.

```text
              HTTP / WebSocket
                     │
                     ▼
Browser ◄────── VTNode WAS
                     │
                     ▼
                   Excel
```

This lets you connect your Excel/VBA application to browsers,
smartphones, and other network applications.

**Excel becomes a Network Node.**

---

# What's Inside?

`vtnode_hello` includes the ready-to-run workbook, along with
the source code and web content that make up the application.

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

`www/default.html` is the **Hello, VTNode!** page
you see in the browser.

You can build web content using HTML, CSS, and JavaScript.

---

# No Apache. No Node.js. No Python.

You do not need to set up a separate Apache or Node.js
application server to run this example.

**No Python** means that you do not need to install Python or have any Python knowledge
to run this example. It does not prohibit integration with Python;
you can integrate with Python when needed.

```text
No Apache
No Node.js
No Python installation or knowledge required
No separate Cloud Application Server
```

All you need is:

```text
Microsoft Excel
      +
VTNode Runtime
```

---

# Don't Study It Yet

You do not need to understand all the VBA source code
or the runtime's internal structure from the start.

The first goal of `vtnode_hello` is simple:

```text
Open
  ↓
START
  ↓
Browser
  ↓
Hello, VTNode!
```

**Run it first.**

Once you have seen it working, you can move on to building it yourself.

---

# Next: Build It Yourself

If you want to see how `vtnode_hello.xlsm` was built,
the next step is **`vtnode_build_myapp`**.

### [Build My App](../vtnode_build_myapp/README.md)

In `vtnode_build_myapp`, you do not start with a finished `.xlsm` file.

Instead, you start with the following file, which has the same basic UI
as `vtnode_hello`:

```text
vtnode_ui_only.xlsx
```

```text
vtnode_ui_only.xlsx
        │
        │ Assemble
        ▼
vtnode_myapp.xlsm
```

Use the provided VBA modules, classes, UserForms, and other components
to turn it into a working VTNode WAS application yourself.

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

Step-by-step instructions for assembling the `.xlsm` file yourself
are provided in a separate **Build Manual**, rather than in the GitHub README.

### [Download Build Manual](https://github.com/uznet/vtnode/blob/main/was/vba/documents/MYAPP_BUILD_GUIDE.en.md)

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

`vtnode_hello` is the starting point.

---

## Start Here

### [Download Runtime](https://vtnode.io/downloads/) | [Build My App](../vtnode_build_myapp/README.md) | [Build Manual](https://github.com/uznet/vtnode/blob/main/was/vba/documents/MYAPP_BUILD_GUIDE.en.md)

---

# VTNode WAS for XLSX (VBA)

> **Enable web service with VBA.**

**VTNode is an AI Glue Gun.**

**Excel becomes a Network Node.**

[VTNode Website](https://vtnode.io)
