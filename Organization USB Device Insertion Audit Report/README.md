# Removable Media Activity Report (NinjaOne) - Organization USB Insert Events

This PowerShell script collects **USB / removable media insertion events** (status: `DISK_DRIVE_ADDED`) across all devices in a specific **NinjaOne organization (customer / organizationId)**, generates an **XLSX report** for the **previous month**, and uploads it to the **Organization Knowledge Base root** (overwrite behavior included).

---

## What it does

- Detects and runs under **PowerShell 7+** (auto-bootstrap from Windows PowerShell 5.1 if needed)
- Connects to NinjaOne via the `NinjaOneDocs` module (using Ninja Script properties)
- Loads all devices in the selected organization
- Fetches device activities filtered by `DISK_DRIVE_ADDED`
- Filters events to the **previous month** time range
- Builds a table with:
  - EventTime
  - Device
  - DiskName
  - DiskModel / Message
  - DiskInterface
  - MediaType
- Exports a lightweight **XLSX** file (no Excel dependency)
- Uploads the XLSX into the **Organization KB root**
- "Overwrite" logic: deletes existing KB items with the same title prefix (including duplicates like `(1)`, `(2)`, etc.) before uploading

---

## Requirements

- NinjaOne environment with API access
- PowerShell 7+ installed on the endpoint  
  (script will fail if PowerShell 7 is missing)
- NinjaOne module:
  - `NinjaOneDocs`

The script expects the following Ninja Script custom properties (retrieved via `Ninja-Property-Get`):
- `ninjaoneInstance` (e.g. `eu`, `app`, `oc`, or a full tenant domain)
- `ninjaoneClientId`
- `ninjaoneClientSecret`

---

## Parameters

| Parameter | Type | Default | Description |
|----------|------|---------|-------------|
| `-Customer` | int | `1` | NinjaOne OrganizationId to report on |

> In Ninja Scripts, the parameter can be overridden by the environment variable `customer`.

---

## Output

- XLSX is created here:
  - `C:\ProgramData\NinjaRMMAgent\scripting\Reports`
- File name format:
  - `Removable Media Activity Report <Month Year> - <Organization Name>.xlsx`

---

## How overwrite works (KB cleanup)

Before uploading the new report, the script:
1. Lists existing Organization KB articles (`organizationId=<Customer>`)
2. Finds entries matching the same report title prefix (including duplicates)
3. Deletes those entries
4. Uploads the new XLSX to the Organization KB root

---

## Notes / Limitations

- The activity fetch is **single page per device** (`pageSize = 250`).
  - If a device returns exactly `pageSize` items, it may indicate truncated results.
- Parsing of the removable media fields depends on the activity payload.
  - If `data.message.params` exists, those values are used.
  - Otherwise, the script attempts to parse the human message text.
- If your NinjaOne activity message language is not English, update the regex patterns in `Parse-FromMessage` accordingly (or support multiple languages).

---

## Customization

You can change these variables inside the script:

- `$StatusFilter` (default `DISK_DRIVE_ADDED`)
- `$PageSize` (default `250`)
- `$MaxRowsToPublish` (default `800`)
- `$MessageMaxLen` (default `600`)
- `$ContinueOnDeviceError` (default `$true`)
- `$DocBaseTitle` (default `"Removable Media Activity Report"`)
- `$outDir` (default `C:\ProgramData\NinjaRMMAgent\scripting\Reports`)

---

## Troubleshooting

### PowerShell 7 not installed
The script checks:
`$env:SystemDrive\Program Files\PowerShell\7`

If missing, install PowerShell 7 and retry.

### KB upload fails (404 / endpoint differences)
Different tenants may expose API paths under `/v2` or `/api/v2`.  
The script tries both automatically.

### No data in report
- Check that activities exist for `DISK_DRIVE_ADDED`
- Verify the previous month date range logic
- Confirm that your tenant uses the same activity fields / message format

---

## Suggested script file name

- `Export-OrgUsbInsertionEvents.ps1`