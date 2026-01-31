# Windows Patch Report (NinjaOne) → XLSX + Upload to Organization KB

This PowerShell script generates a monthly report of **installed** Windows patches for all Windows devices in a NinjaOne Organization, exports a **.xlsx** file with **4 worksheets in a fixed order**, and uploads it to the **Organization Knowledge Base (root)**, attempting to **overwrite** any existing document with the same name.

---

## What it does

1. Reads input:
   - `Customer` (OrganizationId)
   - `ReportMonth` (optional `YYYYMM`)

2. Retrieves the organization devices and keeps only:
   - `WINDOWS_WORKSTATION`
   - `WINDOWS_SERVER`

3. For each device, collects:
   - OS patch installs (`os-patch-installs`)
   - Software patch installs (`software-patch-installs`)

4. Generates an XLSX with 4 worksheets (fixed order):
   1. **OS Summary**
   2. **OS PatchInstalls**
   3. **SW Summary**
   4. **SW PatchInstalls**

5. Highlights (light red fill) rows in `*Summary` sheets where `PatchCount = 0`.

6. Uploads the file to **Organization KB (root)**:
   - best-effort: looks for same-name KB articles and deletes them
   - uploads using multipart to:
     - `/v2/knowledgebase/articles/upload`
     - fallback: `/api/v2/knowledgebase/articles/upload`

---

## Requirements

- **PowerShell 7+**
  - The script auto-relaunches itself under `pwsh` if it is started with Windows PowerShell 5.1.
- NinjaOne API credentials with enough permissions to:
  - read organizations/devices
  - read patch installs
  - upload to knowledge base
- **NinjaOneDocs** PowerShell module
  - The script installs/imports it automatically (PowerShell Gallery access required).

---

## Inputs (Parameters)

| Parameter     | Type   | Description |
|--------------|--------|-------------|
| `Customer`   | int    | NinjaOne OrganizationId |
| `ReportMonth`| string | Month to report in `YYYYMM`. If empty, the script uses the **previous month** |

### Ninja Script Form Variables (Overrides)

If you run it from Ninja as a Script, these environment variables override parameters:

- `customer`
- `reportMonth`

---

## Output

### Local file
The XLSX is written to:

`C:\ProgramData\NinjaRMMAgent\scripting\Reports`

Filename pattern:

`Windows Patch Report <Month Label> - <Organization Name>.xlsx`

> `<Month Label>` uses an Italian month label (e.g. `Gennaio 2026`) because the script formats it using `it-IT`.

### Worksheets

- **OS Summary**
  - `Device`
  - `PatchCount`
  - `LastPatchAt`
- **OS PatchInstalls**
  - `PatchName`
  - `KBNumber`
  - `Device`
  - `InstalledAt`
- **SW Summary**
  - `Device`
  - `PatchCount`
  - `LastPatchAt`
- **SW PatchInstalls**
  - `PatchName`
  - `KBNumber` (may be empty for software patches)
  - `Device`
  - `InstalledAt`

> `DeviceId` is used internally for calculations but is **not exported** to the XLSX.

---

## Sheet order consistency (important)

The XLSX writer uses:

- `[ordered]@{ ... }` when building the sheet map
- `IDictionary` + `.GetEnumerator()` in `New-MinimalXlsxFromTables`

This ensures the **worksheet order in the file matches the insertion order**.

---

## Patch filtering

- Only `status = INSTALLED` is included (`$PatchStatus = "INSTALLED"`).
- Optional noise filter:
  - Excludes “Security Intelligence Update for Microsoft Defender Antivirus” when `$ExcludeDefenderIntel = $true`.

---

## How overwrite works (KB)

1. Builds the KB article name from the XLSX filename (without `.xlsx`).
2. Tries to find existing KB articles with the same name using multiple endpoints.
3. Tries to delete them (best-effort; some tenants may not allow delete).
4. Uploads the new XLSX to the Organization KB root.

> If DELETE is not available in your tenant, the upload may create duplicates. In that case, keep deletion disabled or adjust the upload strategy.

---

## Troubleshooting

### “Worksheet order is different in the file”
Make sure:
- You are passing an **ordered** hashtable (`[ordered]@{...}`).
- The XLSX function parameter type is `IDictionary` and it enumerates via `GetEnumerator()`.

### “Endpoint not reachable: /device/<id>/...”
Your tenant might route endpoints differently. The script tries multiple path variants:
- `v2/device/...`
- `device/...`
- `api/v2/device/...`

If all fail, verify API base URL / instance and required permissions.

### Upload fails (HTTP 401/403)
- Check `ninjaoneClientId` / `ninjaoneClientSecret`
- Verify the OAuth scope/permissions in NinjaOne
- Confirm the tenant base URL (EU/App/OC) mapping is correct

---

## License / Notes

Internal automation utility for NinjaOne reporting. Adapt as needed for your environment.