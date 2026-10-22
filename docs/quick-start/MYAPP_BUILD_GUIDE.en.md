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

The shared VBA runtime is in the **sibling `vtnode_vba_rt64` folder**, outside the working folder.

```text
Path relative to the working folder:
..\vtnode_vba_rt64\vtnode_vba_rt64.xlsm
```

If any files or folders are missing, check the downloaded examples and make sure the archive was fully extracted. VTNode Runtime must also be installed, and Excel must be ready to run macros.

## 1. Create the MyApp workbook

1. Open `vtnode_ui_only.xlsx` from the working folder in Excel.
2. Select **File → Save As**.
3. Save it as `vtnode_myapp.xlsm` in the same folder.
4. Select **Excel Macro-Enabled Workbook (*.xlsm)** as the file type.

## 2. Import the VBA source code

1. Select **Developer → Visual Basic**, or press `Alt+F11`, to open the VBA editor.
2. Select **VBAProject (vtnode_myapp.xlsm)** in Project Explorer.
3. Select **File → Import File** and import all `.cls` and `.bas` files from the working folder’s `src/` folder. Avoid importing the same module more than once.

## 3. Add the VTNodeRuntime reference

1. With the MyApp project selected, open **Tools → References**.
2. Click **Browse**, navigate up to the `xlsx` folder, and select `vtnode_vba_rt64.xlsm` inside the `vtnode_vba_rt64` folder. Change the file type filter if the workbook is not visible.
3. Confirm that the `VTNodeRuntime` reference is selected, then click **OK**.

This reference allows MyApp’s VBA code to use the shared runtime functions.

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
- [ ] Imported the `.cls` and `.bas` files from `src/`.
- [ ] Added the `VTNodeRuntime` reference.
- [ ] Added the open and close event code to `ThisWorkbook`.
- [ ] Saved and reopened the workbook, then started the service with START.
- [ ] Checked SHOW RUNTIME and browser access.
- [ ] Stopped the service with STOP before closing the workbook.

Prepared: 2026-10-22

