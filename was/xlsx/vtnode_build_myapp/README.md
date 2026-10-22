[English](README.md) | [한글](README.ko.md)

# VTNode Build My App

> **Run it first. Then build it yourself.**

`vtnode_build_myapp` is a hands-on project for learning how VTNode WAS for XLSX (VBA)
is assembled and runs inside an Excel workbook.

If you have already run the ready-to-use `vtnode_hello.xlsm`,
it is time to build the same application yourself.

---

## From Hello to My App

`vtnode_hello` is a ready-to-run example.

```text
vtnode_hello.xlsm
        ↓
      START
        ↓
Hello, VTNode!
```

In `vtnode_build_myapp`, you assemble the application yourself.

You start with an Excel workbook that contains only the UI.

```text
vtnode_ui_only.xlsx
        │
        │  Build
        ▼
vtnode_myapp.xlsm
```

The **MyApp** sheet in `vtnode_ui_only.xlsx` uses
the same basic UI as `vtnode_hello`.

The VTNode WAS functionality has not been connected yet.

Connect the provided VBA source code and components
to create a working VTNode WAS application.

---

# What Will You Learn?

You do not need to study the VTNode Runtime API for this project.

You do not need to program a web server from scratch either.

The goal is to **understand how an Excel workbook connects to VTNode WAS by assembling it yourself**.

By the end, you will understand the following structure:

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

An Excel workbook can do more than run locally
with worksheets and VBA.

Connect it to VTNode WAS:

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

This lets it communicate with a browser.

**Excel becomes a Network Node.**

---

## 2. You Don't Have to Build Everything from Scratch

`vtnode_build_myapp` includes the components you need.

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

You do not have to develop a web server engine or runtime library
from scratch.

You will learn what the provided components do
and how they fit together in a single Excel application.

---

## 3. UI and Functionality Are Different Layers

The starting workbook is:

```text
vtnode_ui_only.xlsx
```

Its UI is already in place.

Sheet Name:

```text
MyApp
```

The screen is based on `vtnode_hello`,
but its functionality has not been connected yet.

This illustrates the following layers in an Excel application:

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

**You can think of the UI and executable functionality as separate layers.**

---

## 4. VBA Connects Excel to VTNode

The provided `src/` directory contains VBA modules, classes,
UserForms, and other components for building a VTNode WAS application.

```text
src/
│
├─ *.bas
├─ *.cls
└─ *.frm
```

Assemble them in the workbook:

```text
Excel UI
    │
    ▼
VBA
    │
    ▼
VTNode Runtime
```

This helps you understand how these layers connect.

The goal is to **understand each component's role in the application**,
without having to memorize the runtime API.

---

## 5. Web Contents Are Just Web Contents

You can manage web pages in `www/`.

```text
www/
└─ default.html
```

Connect a web UI built with HTML, CSS, and JavaScript
to your Excel/VBA application.

The resulting structure is:

```text
Excel / VBA                 Browser
     │                         ▲
     │                         │
     └──── VTNode WAS ─────────┘
```

The browser and Excel/VBA communicate through VTNode WAS.

The key is to **connect existing web technologies with Excel**
so you can use them together.

---

# What You Will Build

You start with:

```text
vtnode_ui_only.xlsx
```

This is the workbook you will build on.

By the end, you will have:

```text
vtnode_myapp.xlsm
```

This is your working application workbook.

In your browser, you will see:

```text
Hello, VTNode!

This page is served directly from Excel/VBA.

VTNode WAS for XLSX (VBA)

● Server is running
```

This confirms that the application is running.

The result looks similar to `vtnode_hello`.

There is one key difference:

**This time, you assembled the components yourself
and got the application running.**

---

# Run It → Build It → Change It

This is a simple path to understanding VTNode WAS for XLSX.

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

The goal of `vtnode_build_myapp` extends beyond
creating `vtnode_myapp.xlsm`.

Once you have built it yourself, you can:

```text
Change the UI
Change the HTML
Use Excel Data
Add VBA Logic
Add HTTP Requests
Add WebSocket
```

The aim is to understand how to extend it
into your own application.

---

# No Python Required

**No Python** means that you do not need to install Python or have any Python knowledge
to build and run this example. It does not prohibit integration with Python;
you can integrate with Python when needed.

### [Download VTNode Runtime](https://vtnode.io/downloads/)

---

# Build Manual

This README provides an overview of the project.

Step-by-step instructions for going from `vtnode_ui_only.xlsx`
to `vtnode_myapp.xlsm` are planned for the separate
**VTNode WAS for XLSX Build Manual**.

### [Download Build Manual](https://github.com/uznet/vtnode/blob/main/docs/quick-start/MYAPP_BUILD_GUID.en.md)

[한글](https://github.com/uznet/vtnode/blob/main/docs/quick-start/MYAPP_BUILD_GUIDE.ko.md) | [English](https://github.com/uznet/vtnode/blob/main/docs/quick-start/MYAPP_BUILD_GUID.en.md)

```text
github.com/uznet/vtnode
    ↓
docs/quick-start/
    ↓
MYAPP_BUILD_GUID.en.md
```

The manual is being prepared. It will walk you through the assembly process
using Excel screens and the VBA Editor.

---

# What's Next?

Explore multiple applications and live HTTP / WebSocket monitoring in the
[VTNode WAS Main Demo](https://github.com/uznet/vtnode/tree/main/was/xlsx/vtnode_was).

**VTNode XEvent for XLSX (VBA)** is another related topic to explore next.

---

# Start

### [Quick Start](https://github.com/uznet/vtnode/blob/main/docs/quick-start/QUICK_START.md) | [VTNode Hello](https://github.com/uznet/vtnode/tree/main/was/xlsx/vtnode_hello) | [Build My App](https://github.com/uznet/vtnode/tree/main/was/xlsx/vtnode_build_myapp) | [Documentation](https://vtnode.io/docs/docs_login)

Run `vtnode_hello.xlsm` first.

Then return here.

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

