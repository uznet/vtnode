# VTNode Common VBA Sources

Import shared VBA directly into each application file. The native VTNode Runtime provides the HTTP/WebSocket engine and is installed separately.

```text
was/vba/
├─ common/
│  ├─ src/    # Office shared VBA
│  ├─ xlsx/   # Excel only
│  └─ pptx/   # PowerPoint only
└─ xlsx/      # Excel apps and examples
```

## Direct VBA import and build

The agreed layout is `was/vba/common/src` (Office shared VBA), `was/vba/common/xlsx` (Excel only), `was/vba/common/pptx` (PowerPoint only), and `was/vba/xlsx/` (Excel apps/examples). An application's `src/` holds app-specific code; it does not replace the shared sources.

1. Back up the target file. Save the Excel UI `.xlsx` as `.xlsm`; use `.pptm` for PowerPoint or `.docm` for Word to retain VBA.
2. Open the target file's VBA editor. Use **File > Import File** to import the required `.bas` and `.cls` files and their dependencies from [common/src](https://github.com/uznet/vtnode/tree/main/was/vba/common/src/) directly into that file's VBA project.
3. Add host-specific sources from [common/xlsx](https://github.com/uznet/vtnode/tree/main/was/vba/common/xlsx/) for Excel or [common/pptx](https://github.com/uznet/vtnode/tree/main/was/vba/common/pptx/) for PowerPoint. Keep Excel-only code out of other Office projects. For Word, use Office shared sources and implement and verify the required host integration separately.
4. Import app-specific sources from the application's `src/`. Back up and replace components with matching names to avoid duplicates. Keep each UserForm's `.frm` and `.frx` together and import the `.frm`. Apply `ThisWorkbook`, worksheet, or document event code to the appropriate document modules.
5. Do not open or reference an external runtime workbook. Remove any existing external runtime workbook project reference under **Tools > References**, then check required library references. Check that UI and macro bindings call procedures inside the target file.
6. Run **Debug > Compile VBAProject** and save. Prepare the native VTNode Runtime with the same bitness as Office. Preserve the relative placement of `www`, `audio`, and configuration files, then verify start/stop, HTTP/WebSocket, and Runtime UI behavior. Classification as Office shared source does not establish tested compatibility with every Office host.

[VBA project](https://github.com/uznet/vtnode/tree/main/was/vba) | [Excel apps](https://github.com/uznet/vtnode/tree/main/was/vba/xlsx)
