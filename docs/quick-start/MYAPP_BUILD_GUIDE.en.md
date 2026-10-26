[한글](MYAPP_BUILD_GUIDE.ko.md) | [English](MYAPP_BUILD_GUIDE.en.md)

# VTNode WAS for XLSX — Build MyApp Yourself

Build a web service by connecting VBA source code and VTNode Runtime to an Excel workbook that contains only the UI. This guide presents the manually tested build process step by step.

> **The `vtnode_myapp.xlsm` you build in this guide provides the same functionality and behavior as the completed `vtnode_hello.xlsm` example.**
>
> You will assemble the same application yourself, starting with a UI-only workbook. The output has a different filename, but this exercise does not add new features to the Hello example.

## Preparation

### Download and extract the examples

This guide assumes that you have **downloaded the VTNode examples from GitHub and fully extracted the archive into a folder on your PC**. Extract all files before opening any Excel workbook; do not open the workbooks directly from inside the archive.

Find the `was/xlsx` folder in the extracted repository. The example folders are arranged as follows:

```text
<extracted repository folder>/
└─ was/
   └─ xlsx/
      ├─ vtnode_build_myapp/   ← Working folder for this exercise
      ├─ vtnode_hello/         ← Completed Hello example
      ├─ vtnode_vba_rt64/      ← Shared VBA runtime
      ├─ vtnode_was/           ← WAS demo
      └─ readme.md
```

You may extract the repository to any drive or folder. Keep the relative folder structure intact. **Perform this exercise in `was/xlsx/vtnode_build_myapp`.** This is the “working folder” throughout this guide.

### Check the required files

Confirm that the following files and folders are available in the working folder:

| Item in the working folder | Purpose |
| --- | --- |
| `vtnode_ui_only.xlsx` | Starting workbook containing only the UI |
| `.bas` and `.cls` files in `src/` | Application VBA source code |
| `vtnode_httpa_params.ini` | Web service configuration |
| `www/` | Web content served to the browser |
| `audio/` | Sound effect resources |

The shared VBA runtime source files are in **`vtnode_vba_rt64/src`**, in the sibling runtime folder. Import its `.bas`, `.cls`, and `.frm` files directly into the MyApp project first.

```text
Path relative to the working folder:
..\vtnode_vba_rt64\src\
```

If any files or folders are missing, check the downloaded examples and make sure the archive was fully extracted. VTNode Runtime must also be installed, and Excel must be ready to run macros.

## 1. Create the MyApp workbook

1. Open `vtnode_ui_only.xlsx` from the working folder in Excel.
2. Select **File → Save As**.
3. Save it as `vtnode_myapp.xlsm` in the same folder.
4. Select **Excel Macro-Enabled Workbook (*.xlsm)** as the file type.

## 2. Import the shared VBA runtime source files first

1. Select **Developer → Visual Basic**, or press `Alt+F11`, to open the VBA editor.
2. Select **VBAProject (vtnode_myapp.xlsm)** in Project Explorer.
3. Select **File → Import File**, navigate up from the working folder to `xlsx`, and open **`vtnode_vba_rt64/src`**.
4. Import **all `.bas`, `.cls`, and `.frm` files** from this folder into the MyApp project. Avoid importing the same module more than once.
5. When importing `VtRuntimeConsole.frm`, keep **`VtRuntimeConsole.frx` in the same folder**. The `.frx` file contains form resources and should not be imported separately.

### Add the Microsoft Forms 2.0 Object Library reference

1. With **VBAProject (vtnode_myapp.xlsm)** selected, open **Tools → References**.
2. Find and select **Microsoft Forms 2.0 Object Library**. If it is already selected, leave it selected.
3. Click **OK** to apply the reference.

This reference is required for the runtime UserForm. After importing the `.frm` file, confirm that the reference is selected.

This includes the shared runtime source code directly in MyApp. You do not need to add a reference to `vtnode_vba_rt64.xlsm` through **Tools → References**.

### Add the reference for the source export module

Compiling `Dim comp As VBComponent` and the `vbext_` constants in `ExportProjectSources.bas` also requires this reference:

1. Open **Tools → References**.
2. Select **Microsoft Visual Basic for Applications Extensibility 5.3**, then click **OK**.
3. Select **Debug → Compile VBAProject** to check compilation.

If `VBComponent` produces a “User-defined type not defined” error, check this reference. It is separate from `Microsoft Forms 2.0 Object Library`.

`ExportProjectSources` exports VBA source code to files. You may omit this module if you do not need source export. Running the export macro also requires **Trust access to the VBA project object model** in Excel's Trust Center settings; changing that setting alone does not resolve the compilation error above.

## 3. Import the application VBA source files

1. Select the same **VBAProject (vtnode_myapp.xlsm)**.
2. Select **File → Import File** and open **`src/` inside the working folder, `vtnode_build_myapp`**.
3. Import all `.cls` and `.bas` files from this folder. Avoid importing the same module more than once.

Import the source files in this order: **shared runtime source from `vtnode_vba_rt64/src` → application source from `vtnode_build_myapp/src`**. Add both sets of source files to the same MyApp project.

## 4. Add the workbook open and close code

In Project Explorer, double-click **Microsoft Excel Objects → ThisWorkbook** and add the code below. If `Workbook_Open` or `Workbook_BeforeClose` already exists, do not create a duplicate event procedure.

- On opening: connect the buttons, initialize VTNode, and register the current workbook with the runtime manager.
- On closing: stop the HTTP engine, uninitialize VTNode, and detach the workbook from the runtime manager.

The code below retains the tested logic, with comments and the user-facing message translated into English.

```vb
Option Explicit

Private Sub Workbook_Open()

    On Error GoTo OpenError

    Debug.Print "Workbook_Open()"
    Application.EnableEvents = False

    HelloUI.UI_InitButtons
    WasMain.WasInit

    Call VtRuntimeManager.VtRuntimeAttach(Me.Name)

OpenExit:
    Application.EnableEvents = True
    Exit Sub

OpenError:
    MsgBox _
        "An error occurred while initializing the VTNode UI." & vbCrLf & _
        Err.Description, _
        vbExclamation, _
        "VTNode"

    Resume OpenExit
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error GoTo CloseError

    Application.EnableEvents = False

    Debug.Print "Workbook_BeforeClose: BEGIN"

    ' ---------------------------------------------------------
    ' 1. Stop the native HTTP engine.
    '
    ' If the HTTP engine is running, stop receiving and
    ' processing requests before releasing the native runtime.
    ' ---------------------------------------------------------
    If VtHttpaEngine.VtHttpaIsRun Then

        Debug.Print "Workbook_BeforeClose: VtHttpaStop"
        VtHttpaEngine.VtHttpaStop

    End If

    Debug.Print "Workbook_BeforeClose: VtNodeUnInit"
    VtNode.VtNodeUnInit

    ' ---------------------------------------------------------
    ' 2. Detach the current workbook from the runtime manager.
    '
    ' After native resources have been cleaned up, remove
    ' the current workbook's runtime registration.
    ' ---------------------------------------------------------
    Debug.Print "Workbook_BeforeClose: VtRuntimeDetach"
    VtRuntimeManager.VtRuntimeDetach Me.Name

    Debug.Print "Workbook_BeforeClose: END"

CloseExit:
    Application.EnableEvents = True
    Exit Sub

CloseError:
    Debug.Print _
        "Workbook_BeforeClose Error " & _
        Err.Number & ": " & Err.Description

    Resume CloseExit
End Sub
```

## 5. Save and reopen the workbook

1. Save and close `vtnode_myapp.xlsm`.
2. Open the workbook again.
3. Allow macros to run for this trusted file.

Reopening the workbook runs the `Workbook_Open` code you added, initializing the buttons and runtime.

## 6. Start and verify the web service

1. Click **START** in the Excel UI.
2. Check the status display to confirm that the web service is running.
3. Click **SHOW RUNTIME** to inspect the runtime status.
4. Open the displayed **ACCESS URL** in a browser and confirm that the web page loads.

At this point, you have assembled a web application with the same functionality and behavior as `vtnode_hello.xlsm`.

## 7. Stop the service

1. When you have finished, click **STOP**.
2. Confirm that the service has stopped.
3. Close the workbook.

Although workbook close code is included, finish the exercise in this order: **STOP → confirm the service has stopped → close the workbook**.

## Completion checklist

- [ ] Saved the workbook as `vtnode_myapp.xlsm`.
- [ ] Imported the `.bas`, `.cls`, and `.frm` files from `vtnode_vba_rt64/src` first.
- [ ] Added the `Microsoft Forms 2.0 Object Library` reference.
- [ ] Added the `Microsoft Visual Basic for Applications Extensibility 5.3` reference if including `ExportProjectSources`.
- [ ] Imported the `.cls` and `.bas` files from `vtnode_build_myapp/src` into the same project.
- [ ] Added the open and close event code to `ThisWorkbook`.
- [ ] Saved and reopened the workbook, then started the service with START.
- [ ] Checked SHOW RUNTIME and browser access.
- [ ] Stopped the service with STOP before closing the workbook.

Prepared: 2026-10-22




