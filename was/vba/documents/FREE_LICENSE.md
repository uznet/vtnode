# Get a Free VTNode License

**English** | [한국어](./FREE_LICENSE.ko.md)

This guide explains how to issue and install a free VTNode license after installing the VTNode Runtime.

> You can run VTNode with one channel without a license, time limit, or feature restrictions. Follow this guide only when you need the additional channels provided by a free license.

## Overview

```text
01 Install the VTNode Runtime
        ↓
02 Create a Reg Code on the target PC
        ↓
03 Upload the Reg Code and download a Meta Code
        ↓
04 Activate the Meta Code on the same PC
        ↓
05 Configure the issued License ID
```

## Before you begin

- Install the VTNode Runtime on the Windows PC where the license will be used.
- Confirm that the `VTNODE_HOME` environment variable is configured.
- The web application can be completed on another PC, but the Reg Code must be created and the Meta Code must be activated on the PC that will use the license.
- The target PC must be connected to the internet during Meta Code activation.

## 1. Install the VTNode Runtime

Download and install the latest Windows x64 Runtime from the [VTNode Downloads page](https://vtnode.io/en/downloads/).

The installer configures `VTNODE_HOME` automatically. Run `VTCheck64.exe` after installation and confirm that the final status is `OK`.

## 2. Create a Reg Code

Run `VTMakeRegCode64.exe` on the PC where the license will be used.

When the operation succeeds, the following Reg Code file is created:

```text
%VTNODE_HOME%\licenses\vtnode_regcode.vtreg
```

The file contains hashes derived from the target PC's device identifiers. Do not modify its contents.

> Creating the file does not issue or install a license. Continue with the web issuance step.

## 3. Issue and download a Meta Code

Open the [VTNode Free License page](https://vtnode.io/license/en/free-trial) in a web browser.

1. Select the product for which you want to issue a license.
2. Select the `vtnode_regcode.vtreg` file created in the previous step.
3. Review and accept the license terms and device-hash processing notice.
4. Upload the Reg Code file and complete the issuance process.
5. Download the issued Meta Code (`.vtmeta`) file.

No account registration or email verification is required. If you use a different PC for the web step, copy the downloaded Meta Code file to the target PC.

> A Meta Code is not an executable license file. You must complete activation on the target PC.

## 4. Activate the Meta Code

Connect the target PC to the internet, open Command Prompt or PowerShell, and run:

```powershell
VTActivateMetaCode64.exe "<path-to-meta-code-file>.vtmeta"
```

Example:

```powershell
VTActivateMetaCode64.exe "C:\Users\User\Downloads\<License-ID>.vtmeta"
```

### Why activation is required

The downloaded Meta Code is an issuance token, not a license file that the Runtime can use directly. `VTActivateMetaCode64.exe` performs the following tasks:

- Verifies that the Meta Code is valid and has not been modified
- Confirms that the current PC matches the device for which the Reg Code was created
- Registers the first successful activation with the VTNode license server
- Downloads and stores a signed local license file (`.vtlic`)

This process binds the issued rights to the intended PC and prevents a downloaded Meta Code or license file from being copied and used freely on another device.

### When the license period begins

The free-license period begins when `VTActivateMetaCode64.exe` completes the **first successful online activation**. It does not begin when the Reg Code is created, when the Meta Code is issued, or when the Meta Code is downloaded.

```text
Meta Code issued or downloaded  → License period has not started
First successful activation     → License start date is recorded
```

The VTNode license server records the activation time and calculates the license period according to the applicable free-license policy. Running the activation tool again does not reset or extend the original license period.

The activation tool verifies the Meta Code and the current PC, then requests online activation from the VTNode license server.

When activation succeeds, a signed license file is created in:

```text
%VTNODE_HOME%\licenses\<License-ID>.vtlic
```

Keep the generated License ID. It is the `.vtlic` filename without the extension.

## 5. Configure the License ID

Open `vtnode_httpa_params.ini` used by your VTNode WAS application and configure the issued License ID:

```ini
[license]
enabled = true
id = <License-ID>
```

Example:

```ini
[license]
enabled = true
id = bc0f9373-782d-46da-ab5c-4725b668d152
```

Restart the VTNode WAS application after saving the file. The Runtime finds the matching license file under `%VTNODE_HOME%\licenses` and validates it when the application starts.

## Final checklist

- `VTCheck64.exe` reports `Status : OK`.
- The Reg Code was created on the PC that will use the license.
- The Meta Code was activated on the same PC.
- Online activation completed successfully.
- The license start date was recorded at the first successful activation.
- `%VTNODE_HOME%\licenses\<License-ID>.vtlic` exists.
- The same License ID is configured in `vtnode_httpa_params.ini`.

## Troubleshooting

### The Reg Code file is not created

- Confirm that `VTNODE_HOME` is configured correctly.
- Run `VTCheck64.exe` and check the Runtime installation.
- Confirm that the `licenses` folder is writable.

### Meta Code activation fails

- Confirm that the target PC is connected to the internet.
- Use the Meta Code issued from the Reg Code created on that PC.
- Verify the Meta Code path and enclose a path containing spaces in quotation marks.
- Do not edit the `.vtreg` or `.vtmeta` file.

### The application cannot find the license

- Confirm that the `.vtlic` file exists in `%VTNODE_HOME%\licenses`.
- Enter the License ID without the `.vtlic` extension in `vtnode_httpa_params.ini`.
- Restart Excel and the VTNode WAS application after changing the configuration.

## Related links

- [VTNode Quick Start](./QUICK_START.md)
- [VTNode Downloads](https://vtnode.io/en/downloads/)
- [Free License](https://vtnode.io/license/en/free-trial)
- [VTNode Documentation](https://vtnode.io/docs/en/)

