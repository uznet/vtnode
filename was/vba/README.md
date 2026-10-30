[English](README.md) | [한국어](README.ko.md)

# VTNode WAS VBA

This guide reflects the local directory layout checked on 2026-10-30.

```text
was/vba/
├─ common/
│  ├─ vtnode_vba_rt64.xlsm
│  └─ src/                  Shared VBA runtime sources
│     ├─ *.bas, *.cls, *.frm, *.frx
│     ├─ xlsx/              Empty; reserved for Excel-specific sources
│     └─ pptx/              Empty; reserved for PowerPoint-specific sources
└─ xlsx/                    Excel applications and examples
   ├─ vtnode_hello/
   ├─ vtnode_build_myapp/
   ├─ vtnode_was/
   ├─ examples/
   └─ wrapper/
```

The shared source root is `common/src/`. Application source folders such as `xlsx/vtnode_hello/src/` remain separate. The Office-specific subfolders are placeholders; they do not establish working PowerPoint or Word support. The runtime workbook has not been renamed.

## Start here

- [Excel setup and ZIP unblocking](xlsx/README.md)
- [Hello example](xlsx/vtnode_hello/README.md)
- [Build My App](xlsx/vtnode_build_myapp/README.md)
- [Main demo](xlsx/vtnode_was/README.MD)
- [Common runtime](common/README.MD)

Keep each application's workbook beside its existing `www/`, `audio/`, and configuration files. Existing workbook-relative source export paths continue to use that workbook's `src/` folder. Moving an application folder does not require changing those paths.

## Planned source separation

Use `common/src/` for shared modules and its `xlsx/` and `pptx/` subfolders for host-specific modules when those modules are actually separated. Before moving modules or renaming workbooks, update host dependencies, import/export routines, and workbook references together, then compile and run each supported Office application.
