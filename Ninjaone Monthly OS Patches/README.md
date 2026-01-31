# NinjaOne Windows Patch Report (Org KB Upload)

This repository contains a PowerShell script that generates a **monthly Windows patch installation report** for a specific **NinjaOne Organization (OrgId)**.

The script:
- Connects to **NinjaOne** using API credentials stored in Ninja Script properties
- Retrieves the organization’s **Windows devices**
- Queries **OS patch installs** per device for a selected month
- Generates an **XLSX report** with 2 worksheets:
  - **Summary** (patch count + last patch date per device)
  - **PatchInstalls** (detailed installed patches per device)
- Uploads the XLSX into the **Organization Knowledge Base root** (and optionally tries to delete previous same-name articles to simulate overwrite)

---

## Features

- ✅ Works with **NinjaOne EU / APP / OC** instances (auto-resolves base URL)
- ✅ Uses **Unix timestamps** for `installedAfter` / `installedBefore` filters
- ✅ Queries per-device endpoint:  
  `GET /v2/device/{deviceId}/os-patch-installs`
- ✅ Filters for installed patches only (`status=INSTALLED`)
- ✅ Optional exclusion for Defender Intelligence updates
- ✅ Creates a **valid XLSX** without Excel/COM dependencies
- ✅ Highlights in Summary devices with **0 patches** (light red fill)
- ✅ Uploads report to **Org KB root** using multipart upload

---

## Output

The script generates an Excel file like:

**`Windows Patch Report <Month> - <Organization>.xlsx`**

Saved locally to:

`C:\ProgramData\NinjaRMMAgent\scripting\Reports`

### Worksheet: Summary
Columns:
- `Device`
- `PatchCount`
- `LastPatchAt`

> Devices with `PatchCount = 0` are highlighted.

### Worksheet: PatchInstalls
Columns:
- `PatchName`
- `KBNumber`
- `Device` (device name, not ID)
- `InstalledAt`

---

## Requirements

### Runtime
- PowerShell **7+**
  - If executed with Windows PowerShell 5.1, the script attempts to re-run itself in pwsh automatically.

### NinjaOne
- NinjaOne API credentials must be configured inside Ninja Script custom properties:
  - `ninjaoneInstance` (e.g. `eu`, `app`, `oc` or full URL)
  - `ninjaoneClientId`
  - `ninjaoneClientSecret`

### Module
- The script uses the `NinjaOneDocs` module:
  - If missing, it attempts to install it automatically.

---

## Parameters

### `-Customer`
Organization ID in NinjaOne (OrgId).

Default: `1`  
Can be overridden by Ninja Script form variable `customer`.

### `-ReportMonth`
Month to report in format: `YYYYMM` (example: `202512`).

If empty, the script automatically uses the **previous month**.  
Can be overridden by Ninja Script form variable `reportMonth`.

Examples:
- `202512` → December 2025
- empty → previous month

---

## Usage

### Run locally (manual test)
```powershell
pwsh .\WindowsPatchReport.ps1 -Customer 133 -ReportMonth 202512