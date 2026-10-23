[한글](./README.ko.md) | [English](./README.md)

# ⚠️ STOP — UNBLOCK THE DOWNLOADED ZIP BEFORE EXTRACTING!

> [!WARNING]
> **GitHub ZIP download → Right-click the ZIP → Properties → General → Unblock → Apply → OK → Extract All.**
>
> **Do this BEFORE extracting the ZIP to avoid Excel VBA macros being blocked because the files came from the Internet.** Windows may pass the ZIP's Internet security mark to extracted `.xlsm` files. Unblocking the ZIP afterward does not unblock files already extracted.
>
> **SECURITY: Only unblock files from a source you trust, after checking the repository and download. VBA macros can execute harmful code. Downloading from GitHub alone does not establish trust.**

---

# VTNode WAS for Excel

VTNode WAS connects Excel/VBA applications with web applications using HTTP and WebSocket. See the [existing XLSX/VBA guide](./xlsx/readme.md) for project details and examples.

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
