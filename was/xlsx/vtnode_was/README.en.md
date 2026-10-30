[English](README.MD) | [한글](README.ko.md)

# VTNode WAS for XLSX (VBA)

## Direct VBA import and build

The agreed layout is `was/vba/common/src` (Office shared VBA), `was/vba/common/xlsx` (Excel only), `was/vba/common/pptx` (PowerPoint only), and `was/vba/xlsx/` (Excel apps/examples). An application's `src/` holds app-specific code; it does not replace the shared sources.

1. Back up the target file. Save the Excel UI `.xlsx` as `.xlsm`; use `.pptm` for PowerPoint or `.docm` for Word to retain VBA.
2. Open the target file's VBA editor. Use **File > Import File** to import the required `.bas` and `.cls` files and their dependencies from [common/src](https://github.com/uznet/vtnode/tree/main/was/vba/common/src/) directly into that file's VBA project.
3. Add host-specific sources from [common/xlsx](https://github.com/uznet/vtnode/tree/main/was/vba/common/xlsx/) for Excel or [common/pptx](https://github.com/uznet/vtnode/tree/main/was/vba/common/pptx/) for PowerPoint. Keep Excel-only code out of other Office projects. For Word, use Office shared sources and implement and verify the required host integration separately.
4. Import app-specific sources from the application's `src/`. Back up and replace components with matching names to avoid duplicates. Keep each UserForm's `.frm` and `.frx` together and import the `.frm`. Apply `ThisWorkbook`, worksheet, or document event code to the appropriate document modules.
5. Do not open or reference an external runtime workbook. Remove any existing external runtime workbook project reference under **Tools > References**, then check required library references. Check that UI and macro bindings call procedures inside the target file.
6. Run **Debug > Compile VBAProject** and save. Prepare the native VTNode Runtime with the same bitness as Office. Preserve the relative placement of `www`, `audio`, and configuration files, then verify start/stop, HTTP/WebSocket, and Runtime UI behavior. Classification as Office shared source does not establish tested compatibility with every Office host.


> **Enable web service with VBA.**

`vtnode_was.xlsm` is the Main Demo workbook for **VTNode WAS for XLSX (VBA)**.

Run an HTTP / WebSocket server and multiple applications while
**viewing HTTP requests and responses and WebSocket connection sessions directly in Excel.**

### [Download Runtime](https://vtnode.io/downloads/) | [Quick Start](https://github.com/uznet/vtnode/blob/main/docs/quick-start/QUICK_START.md) | [Build My App](https://github.com/uznet/vtnode/tree/main/was/vba/xlsx/vtnode_build_myapp) | [User Manual](https://vtnode.io/docs/docs_login)

---

# What is `vtnode_was.xlsm`?

`vtnode_hello` is the smallest Quick Start example.

```text
Excel
  ↓
START
  ↓
Browser
  ↓
Hello, VTNode!
```

`vtnode_was.xlsm` is the next step.

It goes beyond serving a web page:

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

Run applications while
**viewing network activity directly in Excel.**

---

# Main Screen

The main screen of `vtnode_was.xlsm` has four main areas.

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

From this single screen, you can:

- Start and stop the server.
- Check the VTNode Runtime status.
- View the ACCESS URL.
- Run example applications.
- View HTTP request and response activity.
- View WebSocket connection sessions.

---

# Server / Runtime

The **SERVER / RUNTIME** area at the top left lets you
control and check the VTNode WAS server's status.

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

Starts the VTNode WAS HTTP / WebSocket server.

### STOP SERVER

Stops the running server.

### SHOW RUNTIME

Displays the VTNode Runtime Console.

### SERVER STATUS

Displays the server's current status.

### ACCESS URL

Displays the URL you can open in a browser
or on another device on the same network.

---

# Applications

The **APPLICATIONS** area at the top right lets you
run multiple applications from a single Excel workbook.

```text
APPLICATIONS

[ WHAT IS VTNODE WAS? / INTRO ]

[ WEBSOCKET DEMO ]

[ TABLE ORDER ]

[ + YOUR APP ]
```

Each application can have its own URI route.

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

This shows how a single VTNode WAS instance can support
multiple Excel/web applications.

---

# Request → Response Activity

This is one of the key differences between
`vtnode_was.xlsm` and `vtnode_hello`.

You can watch HTTP requests arrive and responses return
in real time on the Excel screen.

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

The log displays information such as:

```text
TIME
METHOD
REQUEST / URI
RESPONSE / STATUS
DURATION
DETAIL
```

For example, when a browser requests a page:

```text
GET
   ↓
/index.html
   ↓
Excel / VBA
   ↓
200 OK
```

You can follow this request/response flow in Excel.

---

# See HTTP Working

VTNode WAS makes HTTP processing visible.

While using an application in your browser, watch
**REQUEST → RESPONSE ACTIVITY** in Excel:

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

You can observe this flow directly.

Alongside its application demos, `vtnode_was.xlsm` serves as a
**Live Monitor** that helps you understand how VTNode WAS works.

---

# WebSocket Connection Sessions

The second monitoring area is:

**WEBSOCKET CONNECTION SESSIONS**



When a WebSocket client connects, you can view each connection
as an individual session in Excel.

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

The session area displays the following information:

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

WebSocket connections work differently from HTTP requests and responses.

An HTTP exchange follows this pattern:

```text
Request
   ↓
Response
   ↓
Completed
```

A WebSocket connection follows this pattern instead:

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

The connection stays open for ongoing communication.

In `vtnode_was.xlsm`, you can
**manage and observe WebSocket connections as individual sessions.**

```text
Session #1    OPEN
Session #2    OPEN
Session #3    CLOSED
```

Each row remains after its connection closes,
so you can review the final status and activity.

Use **CLEAR CLOSED** to clear completed sessions.

---

# HTTP and WebSocket at a Glance

The main screen lets you compare both network models side by side.

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

These two areas are the main monitoring views for seeing
VTNode WAS for XLSX connectivity in action.

---

# WebSocket Demo

**WEBSOCKET DEMO** is an example for observing WebSocket connections
and sending and receiving messages.

```text
Browser
       ⇅
   WebSocket
       ⇅
VTNode WAS
       ⇅
Excel / VBA
```

When you establish a connection in the browser,
a session appears under **WEBSOCKET CONNECTION SESSIONS** in Excel.

As messages are exchanged, check these fields:

```text
RECEIVED MESSAGE

SEND MESSAGE

LAST ACTIVITY
```

These fields show activity in the current session.

**You can see an actual WebSocket connection in Excel**
and observe how it behaves.

---

# Table Order

**TABLE ORDER** demonstrates how to build a practical application
with VTNode WAS.

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

Users interact with the application in a smartphone browser,
while Excel handles the application logic and data.

You can also view HTTP activity on the main screen under:

```text
REQUEST → RESPONSE ACTIVITY
```

This view updates as HTTP activity occurs.

Table Order shows how the following pieces work together:

```text
Web UI
   +
HTTP
   +
Excel/VBA
   +
Excel Data
```

Together, they form a single application.

---

# One Workbook, Multiple Apps

A key concept in `vtnode_was.xlsm` is that
**a single Excel workbook can host multiple applications.**

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

Requests can be directed to different applications based on their URI.

The main screen includes:

```text
+ YOUR APP
ADD ANY URI ROUTE
```

These labels indicate that you can extend this structure with new applications.

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

The VBA module, class, and UserForm source code that makes up the workbook.

### `audio/`

Audio resources used by the applications and runtime.

### `www/`

HTML, CSS, JavaScript, and media content served to browsers.

---

# What Makes `vtnode_was` Different?

Each VTNode WAS for XLSX example has a different purpose.

```text
vtnode_hello
────────────────────────
Run It

Run the smallest example
right away.


vtnode_build_myapp
────────────────────────
Build It

Assemble the UI and
provided VBA source code yourself.


vtnode_was
────────────────────────
Explore It

Run multiple applications and observe

HTTP Request / Response

WebSocket Sessions

in real time.
```

With `vtnode_was.xlsm`,
**you can see what actually happens between the browser and Excel.**

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

Excel becomes an environment that runs application logic, manages data,
and lets you observe network activity directly.

**Excel becomes a Network Node.**

---

# No Apache. No Node.js. No Python.

**No Python** means that you do not need to install Python or have any Python knowledge
to run this demo. It does not prohibit integration with Python;
you can integrate with Python when needed.

Use VTNode Runtime and Excel/VBA without setting up
a separate Apache or Node.js application server.

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

If you are getting started, we recommend running
`vtnode_hello` before exploring `vtnode_was.xlsm`.

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

### [Quick Start](https://github.com/uznet/vtnode/blob/main/docs/quick-start/QUICK_START.md)

### [Build My App](https://github.com/uznet/vtnode/tree/main/was/vba/xlsx/vtnode_build_myapp)

---

# What's Next?

VTNode WAS connectivity can extend beyond the web.

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

Start with HTTP / WebSocket, then expand to connectivity
between local Excel data and cloud AI.

**VTNode XEvent for XLSX (VBA)** is another related topic to explore next.

---

# Documentation

The official documentation provides step-by-step instructions
for setting up VTNode WAS for XLSX yourself.

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

### [Download Runtime](https://vtnode.io/downloads/) | [Quick Start](https://github.com/uznet/vtnode/blob/main/docs/quick-start/QUICK_START.md) | [Build My App](https://github.com/uznet/vtnode/tree/main/was/vba/xlsx/vtnode_build_myapp) | [User Manual](https://vtnode.io/docs/docs_login)

---

# VTNode WAS for XLSX (VBA)

> **Enable web service with VBA.**

**Excel becomes a Network Node.**

**VTNode is an AI Glue Gun.**

[VTNode Website](https://vtnode.io)

Copyright © 2026 VTNode