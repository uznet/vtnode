[한글](./README.ko.md) | [English](./README.md)

[VBA project](https://github.com/uznet/vtnode/tree/main/was/vba) | [Current directory](https://github.com/uznet/vtnode/tree/main/was/vba/xlsx)

# ⚠️ STOP — UNBLOCK THE DOWNLOADED ZIP BEFORE EXTRACTING!

## Direct VBA import and build

The agreed layout is `was/vba/common/src` (Office shared VBA), `was/vba/common/xlsx` (Excel only), `was/vba/common/pptx` (PowerPoint only), and `was/vba/xlsx/` (Excel apps/examples). An application's `src/` holds app-specific code; it does not replace the shared sources.

1. Back up the target file. Save the Excel UI `.xlsx` as `.xlsm`; use `.pptm` for PowerPoint or `.docm` for Word to retain VBA.
2. Open the target file's VBA editor. Use **File > Import File** to import the required `.bas` and `.cls` files and their dependencies from [common/src](../common/src/) directly into that file's VBA project.
3. Add host-specific sources from [common/xlsx](../common/xlsx/) for Excel or [common/pptx](../common/pptx/) for PowerPoint. Keep Excel-only code out of other Office projects. For Word, use Office shared sources and implement and verify the required host integration separately.
4. Import app-specific sources from the application's `src/`. Back up and replace components with matching names to avoid duplicates. Keep each UserForm's `.frm` and `.frx` together and import the `.frm`. Apply `ThisWorkbook`, worksheet, or document event code to the appropriate document modules.
5. Do not open or reference an external runtime workbook. Remove any existing external runtime workbook project reference under **Tools > References**, then check required library references. Check that UI and macro bindings call procedures inside the target file.
6. Run **Debug > Compile VBAProject** and save. Prepare the native VTNode Runtime with the same bitness as Office. Preserve the relative placement of `www`, `audio`, and configuration files, then verify start/stop, HTTP/WebSocket, and Runtime UI behavior. Classification as Office shared source does not establish tested compatibility with every Office host.


> [!WARNING]
> **GitHub ZIP download → Right-click the ZIP → Properties → General → Unblock → Apply → OK → Extract All.**
>
> **Do this BEFORE extracting the ZIP to avoid Excel VBA macros being blocked because the files came from the Internet.** Windows may pass the ZIP's Internet security mark to extracted `.xlsm` files. Unblocking the ZIP afterward does not unblock files already extracted.
>
> **SECURITY: Only unblock files from a source you trust, after checking the repository and download. VBA macros can execute harmful code. Downloading from GitHub alone does not establish trust.**

---

# VTNode WAS for Excel

VTNode WAS connects Excel/VBA applications with web applications using HTTP and WebSocket. See the [existing XLSX/VBA guide](./vtnode_hello/README.md) for project details and examples.

## Windows: before extracting a GitHub ZIP

1. Download the repository ZIP from GitHub and locate it in File Explorer.
2. Confirm that you trust the repository and the downloaded file.
3. Right-click the **ZIP file**, then select **Properties**.
4. On the **General** tab, find **Unblock** near the bottom under Security and check it.
5. Click **Apply**, then **OK**.
6. Right-click the ZIP and choose **Extract All**. Open the required workbook from the extracted folder.

## Already extracted? Unblock the XLSM file itself

1. Close the workbook in Excel.
2. In File Explorer, right-click the actual **`.xlsm` file** you intend to use and select **Properties**.
3. Under **General → Security**, check **Unblock**, click **Apply**, then **OK**.
4. Reopen the workbook. If Excel requests permission to enable macros, allow them only after verifying that you trust the workbook and its code.
5. Repeat for each trusted XLSM workbook you need. Alternatively, unblock the original ZIP and extract a fresh copy into a new folder; preserve any work in the existing copy.

## If macros are still blocked

The Unblock option may be absent if no Internet security mark is present. Unblocking removes that mark; Excel settings or organization policies can still prevent macros from running. Follow your organization's rules and contact your administrator if necessary. Do not globally enable all macros or disable security protections to resolve this issue.

[Microsoft: Macros from the Internet are blocked by default in Office](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked)
