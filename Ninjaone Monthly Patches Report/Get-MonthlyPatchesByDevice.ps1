[CmdletBinding()]
param(
    [Parameter()]
    [int]$Customer = 1,

    # Format: YYYYMM (e.g. 202512). If empty, previous month is used.
    [Parameter()]
    [string]$ReportMonth = ""
)

begin {
    # Ninja Script Form Variable overrides
    if ($env:customer -and $env:customer -notlike "null") { $Customer = [int]$env:customer }
    if ($env:reportMonth -and $env:reportMonth -notlike "null") { $ReportMonth = [string]$env:reportMonth }

    if (-not $Customer -or $Customer -lt 1) {
        Write-Host -Object "[Error] Customer (OrganizationId) is required and must be >= 1"
        exit 1
    }
}

process {

# ==================== CONFIG ====================
$PageSize              = 1000
$LogEveryNDevices       = 10
$ContinueOnDeviceError  = $true

# Patch query filters
$PatchStatus            = "INSTALLED"
$ExcludeDefenderIntel   = $true  # Exclude Microsoft Defender Intelligence Update noise (optional)

# Output settings
$DocBaseTitle           = "Windows Patch Report"
$outDir                 = "C:\ProgramData\NinjaRMMAgent\scripting\Reports"
New-Item -Path $outDir -ItemType Directory -Force | Out-Null
# ===============================================


# -------------------- PowerShell 7+ bootstrap --------------------
if ($PSVersionTable.PSVersion.Major -lt 7) {
    try {
        if (!(Test-Path "$env:SystemDrive\Program Files\PowerShell\7")) {
            Write-Output 'PowerShell 7 is not installed.'
            exit 1
        }
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + `
                    [System.Environment]::GetEnvironmentVariable('Path','User')
        pwsh -File "`"$PSCommandPath`"" @PSBoundParameters
    } catch {
        Write-Output 'PowerShell 7 was not installed. Update PowerShell and try again.'
        throw
    } finally { exit $LASTEXITCODE }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO"
    )
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Output "[$ts][$Level] $Message"
}

function Has-Prop($obj, [string]$name) {
    return ($null -ne $obj -and $obj.PSObject -and $obj.PSObject.Properties.Match($name).Count -gt 0)
}
function Get-Prop($obj, [string]$name) {
    if (Has-Prop $obj $name) { return $obj.$name }
    return $null
}

function Normalize-ToArray {
    <#
      Normalizes various API payload shapes into a plain array.
      Supports common property containers: items, data, results, etc.
    #>
    param([object]$obj)

    if ($null -eq $obj) { return @() }
    if ($obj -is [string]) { return @($obj) }
    if ($obj -is [System.Array]) { return $obj }

    foreach ($k in @('items','data','results','documents','activities','organizations','devices','folders','patchInstalls')) {
        if ($obj.PSObject -and $obj.PSObject.Properties.Match($k).Count -gt 0 -and $obj.$k) {
            return (Normalize-ToArray $obj.$k)
        }
    }
    return @($obj)
}

function Resolve-NinjaBaseUrl([string]$InstanceValue) {
    # Maps instance values to base URLs
    if ([string]::IsNullOrWhiteSpace($InstanceValue)) { return "https://eu.ninjarmm.com" }
    $v = $InstanceValue.Trim()
    switch -Regex ($v.ToLowerInvariant()) {
        '^eu$'  { return "https://eu.ninjarmm.com" }
        '^app$' { return "https://app.ninjarmm.com" }
        '^oc$'  { return "https://oc.ninjarmm.com" }
    }
    if ($v -match '^https?://') { return $v.TrimEnd('/') }
    return ("https://{0}" -f $v).TrimEnd('/')
}

function Get-OAuthToken {
    <#
      Retrieves a raw OAuth token for direct upload calls (multipart).
      This is independent from the NinjaOneDocs module session.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$BaseUrl,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][string]$ClientSecret
    )

    $tokenUrl = "$($BaseUrl.TrimEnd('/'))/ws/oauth/token"
    $resp = Invoke-RestMethod -Method Post `
        -Uri $tokenUrl `
        -ContentType "application/x-www-form-urlencoded" `
        -Body @{
            grant_type    = "client_credentials"
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "monitoring management"
        }

    if (-not $resp -or -not $resp.access_token) { throw "OAuth token missing from response." }
    return $resp.access_token
}

function Convert-EpochToLocalStringSafe([object]$ep) {
    # Converts Unix epoch seconds to local datetime string (dd/MM/yyyy HH:mm:ss)
    if ($null -eq $ep) { return "" }
    try {
        if ($ep -is [double]) { $ep = [int64][math]::Floor($ep) }
        elseif ($ep -is [int]) { $ep = [int64]$ep }
        elseif ($ep -is [string]) {
            if ($ep -match '^\d+$') { $ep = [int64]$ep } else { return [string]$ep }
        }
        return ([DateTimeOffset]::FromUnixTimeSeconds([int64]$ep).ToLocalTime().DateTime).ToString("dd/MM/yyyy HH:mm:ss")
    } catch {
        return [string]$ep
    }
}

function Get-ReportMonthInfo {
    <#
      Determines the report month boundaries.
      If ReportMonth is empty, uses the previous month.
    #>
    param([string]$Yyyymm)

    if ($Yyyymm -and $Yyyymm -match '^\d{6}$') {
        $y = [int]$Yyyymm.Substring(0,4)
        $m = [int]$Yyyymm.Substring(4,2)
        $dt = Get-Date -Year $y -Month $m -Day 1 -Hour 0 -Minute 0 -Second 0
    } else {
        $now = Get-Date
        $firstThisMonth = Get-Date -Year $now.Year -Month $now.Month -Day 1 -Hour 0 -Minute 0 -Second 0
        $dt = $firstThisMonth.AddMonths(-1)
    }

    $start = $dt
    $end   = $dt.AddMonths(1).AddSeconds(-1)

    # Use Italian month label (keeps your previous behavior)
    $it  = [System.Globalization.CultureInfo]::GetCultureInfo("it-IT")
    $rawLabel = $dt.ToString("MMMM yyyy", $it)
    $label = ($it.TextInfo.ToTitleCase($rawLabel))

    [pscustomobject]@{
        Start   = $start
        End     = $end
        Yyyymm  = $dt.ToString("yyyyMM")
        LabelIt = $label
    }
}

# ---------- KB helpers (overwrite behavior) ----------
function Get-OrgKbArticlesByName {
    <#
      Tries a few variants because KB endpoints are not always consistent between tenants.
    #>
    param(
        [Parameter(Mandatory)][int]$OrganizationId,
        [Parameter(Mandatory)][string]$ArticleName
    )

    $tries = @(
        @{ Path="knowledgebase/organization/articles"; Q="organizationId=$OrganizationId&articleName=$([uri]::EscapeDataString($ArticleName))" },
        @{ Path="v2/knowledgebase/organization/articles"; Q="organizationId=$OrganizationId&articleName=$([uri]::EscapeDataString($ArticleName))" },
        @{ Path="api/v2/knowledgebase/organization/articles"; Q="organizationId=$OrganizationId&articleName=$([uri]::EscapeDataString($ArticleName))" }
    )

    foreach ($t in $tries) {
        try {
            $resp = Invoke-NinjaOneRequest -Method GET -Path $t.Path -QueryParams $t.Q
            $arr = @(Normalize-ToArray $resp)
            if ($arr.Count -gt 0) { return $arr }
        } catch { }
    }
    return @()
}

function Remove-KbArticleById {
    param([Parameter(Mandatory)][int]$Id)

    $pathsToTry = @(
        ("knowledgebase/articles/{0}" -f $Id),
        ("v2/knowledgebase/articles/{0}" -f $Id),
        ("api/v2/knowledgebase/articles/{0}" -f $Id)
    )

    foreach ($p in $pathsToTry) {
        try {
            $null = Invoke-NinjaOneRequest -Method DELETE -Path $p
            Write-Log ("Deleted KB article id={0} via '{1}'" -f $Id, $p) "INFO"
            return $true
        } catch { }
    }

    Write-Log ("Unable to delete KB article id={0} (DELETE not available?). Upload may duplicate." -f $Id) "WARN"
    return $false
}

function Upload-FileToOrgKbRoot {
    <#
      Uploads a file into the Organization Knowledge Base ROOT (no folderPath/folderId).
    #>
    param(
        [Parameter(Mandatory=$true)][string]$BaseUrl,
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][int]$OrganizationId,
        [Parameter(Mandatory=$true)][string]$FilePath
    )

    Add-Type -AssemblyName System.Net.Http

    $http = New-Object System.Net.Http.HttpClient
    $http.DefaultRequestHeaders.Authorization =
        New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $AccessToken)
    $http.DefaultRequestHeaders.Accept.Add(
        (New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"))
    )

    $multipart = New-Object System.Net.Http.MultipartFormDataContent
    $multipart.Add((New-Object System.Net.Http.StringContent($OrganizationId.ToString())), "organizationId")

    $stream = [System.IO.File]::OpenRead($FilePath)
    try {
        $fileContent = New-Object System.Net.Http.StreamContent($stream)
        $fileContent.Headers.ContentType =
            [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/octet-stream")
        $multipart.Add($fileContent, "files", [System.IO.Path]::GetFileName($FilePath))

        $base = $BaseUrl.TrimEnd('/')
        $uris = @(
            "$base/v2/knowledgebase/articles/upload",
            "$base/api/v2/knowledgebase/articles/upload"
        )

        $lastError = $null
        foreach ($u in $uris) {
            Write-Log ("Uploading to: {0}" -f $u) "INFO"
            $resp = $http.PostAsync($u, $multipart).Result
            $body = $resp.Content.ReadAsStringAsync().Result
            if ($resp.IsSuccessStatusCode) {
                Write-Log "Upload OK." "INFO"
                return $true
            }
            $lastError = "HTTP $([int]$resp.StatusCode) $($resp.ReasonPhrase) - $body"
        }

        Write-Log ("Upload failed: {0}" -f $lastError) "ERROR"
        return $false
    }
    finally {
        try { if ($stream) { $stream.Dispose() } } catch {}
        try { if ($multipart) { $multipart.Dispose() } } catch {}
        try { if ($http) { $http.Dispose() } } catch {}
    }
}

# ---------- XLSX writer (minimal OpenXML, no external modules required) ----------
function New-MinimalXlsxFromTables {
    <#
      Creates a minimal valid .xlsx with multiple worksheets and a tiny styles.xml.
      - Writes all values as inline strings.
      - Adds one highlight style (light red fill) applied to rows where PatchCount == 0
        on any sheet whose name ends with "Summary".
      IMPORTANT: uses IDictionary + GetEnumerator() to preserve insertion order.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [System.Collections.IDictionary]$Sheets,  # preserves [ordered] insertion order

        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function Escape-Xml([string]$s) {
        if ($null -eq $s) { return "" }
        return ($s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'","&apos;")
    }
    function ColName([int]$index) {
        $name = ""
        $i = $index
        while ($i -gt 0) {
            $i--
            $name = [char](65 + ($i % 26)) + $name
            $i = [int]($i / 26)
        }
        return $name
    }
    function CellInlineStr([string]$r, [string]$text, [int]$styleIndex = 0) {
        $t = Escape-Xml $text
        if ($styleIndex -gt 0) {
            return "<c r=`"$r`" s=`"$styleIndex`" t=`"inlineStr`"><is><t>$t</t></is></c>"
        }
        return "<c r=`"$r`" t=`"inlineStr`"><is><t>$t</t></is></c>"
    }

    # Styles: 0 = default, 1 = red fill
    $stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFFC7CE"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/>
  </cellXfs>
</styleSheet>
"@

    $sheetEntries = @()
    $sheetParts   = @()

    $sheetId = 0

    foreach ($entry in $Sheets.GetEnumerator()) {
        $sheetName = [string]$entry.Key
        $rows      = @($entry.Value)

        $sheetId++
        $rid = "rId$sheetId"
        $sheetFile = "xl/worksheets/sheet$sheetId.xml"
        $sheetEntries += [pscustomobject]@{ name=$sheetName; id=$sheetId; rid=$rid; file=$sheetFile }

        $headers = @()
        if ($rows.Count -gt 0) { $headers = $rows[0].PSObject.Properties.Name }

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        [void]$sb.AppendLine('<sheetData>')

        # Header row
        $rowIndex = 1
        [void]$sb.Append("<row r=`"$rowIndex`">")
        for ($c=0; $c -lt $headers.Count; $c++) {
            $col = ColName ($c+1)
            $ref = "$col$rowIndex"
            [void]$sb.Append((CellInlineStr $ref ([string]$headers[$c]) 0))
        }
        [void]$sb.AppendLine("</row>")

        # Data rows
        for ($i=0; $i -lt $rows.Count; $i++) {
            $rowIndex = $i + 2

            # Highlight rows with PatchCount == 0 on any *Summary sheet
            $applyStyle = 0
            if ($sheetName -like "*Summary") {
                try {
                    $pcProp = $rows[$i].PSObject.Properties["PatchCount"]
                    if ($pcProp -and [int]$pcProp.Value -eq 0) { $applyStyle = 1 }
                } catch {}
            }

            [void]$sb.Append("<row r=`"$rowIndex`">")
            for ($c=0; $c -lt $headers.Count; $c++) {
                $col = ColName ($c+1)
                $ref = "$col$rowIndex"
                $val = $null
                try { $val = $rows[$i].PSObject.Properties[$headers[$c]].Value } catch {}
                [void]$sb.Append((CellInlineStr $ref ([string]$val) $applyStyle))
            }
            [void]$sb.AppendLine("</row>")
        }

        [void]$sb.AppendLine('</sheetData>')
        [void]$sb.AppendLine('</worksheet>')

        $sheetParts += [pscustomobject]@{ path=$sheetFile; content=$sb.ToString() }
    }

    # workbook.xml
    $sheetsXml = New-Object System.Text.StringBuilder
    foreach ($e in $sheetEntries) {
        [void]$sheetsXml.AppendLine(("    <sheet name=`"{0}`" sheetId=`"{1}`" r:id=`"{2}`"/>" -f (Escape-Xml $e.name), $e.id, $e.rid))
    }

    $workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
$($sheetsXml.ToString().TrimEnd())
  </sheets>
</workbook>
"@

    # Root relationships
    $relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@

    # Workbook relationships (sheets + styles)
    $wbRels = New-Object System.Text.StringBuilder
    foreach ($e in $sheetEntries) {
        [void]$wbRels.AppendLine(("  <Relationship Id=`"{0}`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet{1}.xml`"/>" -f $e.rid, $e.id))
    }
    [void]$wbRels.AppendLine(("  <Relationship Id=`"rIdStyles`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>"))

    $workbookRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
$($wbRels.ToString().TrimEnd())
</Relationships>
"@

    # Content types
    $ct = New-Object System.Text.StringBuilder
    [void]$ct.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$ct.AppendLine('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
    [void]$ct.AppendLine('  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
    [void]$ct.AppendLine('  <Default Extension="xml" ContentType="application/xml"/>')
    [void]$ct.AppendLine('  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    [void]$ct.AppendLine('  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    foreach ($e in $sheetEntries) {
        [void]$ct.AppendLine(("  <Override PartName=`"/xl/worksheets/sheet{0}.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>" -f $e.id))
    }
    [void]$ct.AppendLine('</Types>')
    $contentTypesXml = $ct.ToString()

    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            function Add-ZipEntry([System.IO.Compression.ZipArchive]$z, [string]$entryName, [string]$content) {
                $e = $z.CreateEntry($entryName)
                $s = $e.Open()
                try {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
                    $s.Write($bytes, 0, $bytes.Length)
                } finally { $s.Dispose() }
            }

            Add-ZipEntry $zip "[Content_Types].xml" $contentTypesXml
            Add-ZipEntry $zip "_rels/.rels" $relsXml
            Add-ZipEntry $zip "xl/workbook.xml" $workbookXml
            Add-ZipEntry $zip "xl/_rels/workbook.xml.rels" $workbookRelsXml
            Add-ZipEntry $zip "xl/styles.xml" $stylesXml

            foreach ($sp in $sheetParts) {
                Add-ZipEntry $zip $sp.path $sp.content
            }
        }
        finally { $zip.Dispose() }
    }
    finally { $fs.Dispose() }
}

function Build-SummaryRows {
    <#
      Builds a per-device summary (PatchCount + LastPatchAt).
      DetailSorted can be empty (e.g. SW patching not enabled).
    #>
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[object]]$DeviceList,

        [AllowEmptyCollection()]
        [object[]]$DetailSorted = @()
    )

    $counts = @{}
    $latest = @{}
    foreach ($dev in $DeviceList) {
        $counts[[int]$dev.DeviceId] = 0
        $latest[[int]$dev.DeviceId] = ""
    }

    foreach ($r in $DetailSorted) {
        $did = [int]$r.DeviceId
        $counts[$did] = [int]$counts[$did] + 1

        try {
            $dt = [datetime]::ParseExact($r.InstalledAt, "dd/MM/yyyy HH:mm:ss", $null)
            if ($latest[$did] -eq "") {
                $latest[$did] = $r.InstalledAt
            } else {
                $cur = [datetime]::ParseExact([string]$latest[$did], "dd/MM/yyyy HH:mm:ss", $null)
                if ($dt -gt $cur) { $latest[$did] = $r.InstalledAt }
            }
        } catch { }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($dev in $DeviceList) {
        $did = [int]$dev.DeviceId
        $rows.Add([pscustomobject]@{
            Device      = [string]$dev.DeviceName
            PatchCount  = [int]$counts[$did]
            LastPatchAt = [string]$latest[$did]
        }) | Out-Null
    }

    return @(
        $rows |
        Sort-Object @{Expression={ $_.PatchCount }; Descending=$false },
                    @{Expression={ $_.Device }; Descending=$false }
    )
}

function Get-DevicePatchInstalls {
    <#
      Calls per-device endpoints:
        /v2/device/{id}/os-patch-installs
        /v2/device/{id}/software-patch-installs
      Tries multiple path prefixes to support different tenant routing.
    #>
    param(
        [Parameter(Mandatory)][int]$DeviceId,
        [Parameter(Mandatory)][string]$EndpointSuffix,  # "os-patch-installs" or "software-patch-installs"
        [Parameter(Mandatory)][long]$AfterEpoch,
        [Parameter(Mandatory)][long]$BeforeEpoch,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][int]$PageSize
    )

    $pathsToTry = @(
        ("v2/device/{0}/{1}" -f $DeviceId, $EndpointSuffix),
        ("device/{0}/{1}"    -f $DeviceId, $EndpointSuffix),
        ("api/v2/device/{0}/{1}" -f $DeviceId, $EndpointSuffix)
    )

    $qp = "installedAfter=$AfterEpoch&installedBefore=$BeforeEpoch&status=$Status&pageSize=$PageSize"

    foreach ($p in $pathsToTry) {
        try {
            return (Invoke-NinjaOneRequest -Method GET -Path $p -QueryParams $qp)
        } catch {
            # Try next path variant
        }
    }

    throw "Endpoint not reachable: /device/$DeviceId/$EndpointSuffix"
}

# -------------------- Module load + Connect (NinjaOneDocs) --------------------
try {
    $moduleName = "NinjaOneDocs"
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Install-Module -Name $moduleName -Force -AllowClobber
    }
    Import-Module $moduleName -ErrorAction Stop
} catch {
    Write-Error "Failed to import NinjaOneDocs module. Error: $_"
    exit 1
}

$NinjaOneInstance     = Ninja-Property-Get ninjaoneInstance
$NinjaOneClientId     = Ninja-Property-Get ninjaoneClientId
$NinjaOneClientSecret = Ninja-Property-Get ninjaoneClientSecret

if (!$NinjaOneClientId -or !$NinjaOneClientSecret) {
    Write-Output "Missing required API credentials (ninjaoneClientId / ninjaoneClientSecret)"
    exit 1
}

Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientId -NinjaOneClientSecret $NinjaOneClientSecret
Write-Log ("Connected (instance={0})." -f $NinjaOneInstance) "INFO"

# -------------------- MAIN --------------------
try {
    $mi = Get-ReportMonthInfo -Yyyymm $ReportMonth

    # Unix timestamps (seconds), UTC boundaries
    $afterEpoch  = [DateTimeOffset]::new($mi.Start.ToUniversalTime()).ToUnixTimeSeconds()
    $beforeEpoch = [DateTimeOffset]::new($mi.End.ToUniversalTime()).ToUnixTimeSeconds()

    Write-Log ("OrgId: {0}" -f $Customer) "INFO"
    Write-Log ("Month: {0} | Range: {1} -> {2}" -f $mi.LabelIt, $mi.Start, $mi.End) "INFO"

    # Resolve org name (best-effort)
    $orgName = ("OrgId {0}" -f $Customer)
    try {
        $orgs = Invoke-NinjaOneRequest -Method GET -Path "organizations"
        $o = (Normalize-ToArray $orgs) | Where-Object { $_ -and (Get-Prop $_ 'id') -and [int]$_.id -eq $Customer } | Select-Object -First 1
        if ($o -and (Get-Prop $o 'name')) { $orgName = [string]$o.name }
    } catch {}
    Write-Log ("Organization: {0}" -f $orgName) "INFO"

    # Load org devices and keep Windows only
    Write-Log "Fetching organization devices..." "INFO"
    $orgDevices = @(Normalize-ToArray (Invoke-NinjaOneRequest -Method GET -Path ("organization/{0}/devices" -f $Customer)))

    $deviceList = New-Object System.Collections.Generic.List[object]
    foreach ($d in $orgDevices) {
        $idVal = Get-Prop $d 'id'
        if (-not $idVal) { continue }
        $did = [int]$idVal

        $nodeClass = [string](Get-Prop $d 'nodeClass')
        if ($nodeClass -and ($nodeClass -notin @("WINDOWS_WORKSTATION","WINDOWS_SERVER"))) { continue }

        $nm = $null
        foreach ($k in @('systemName','hostname','displayName','name')) {
            $tmp = Get-Prop $d $k
            if ($tmp) { $nm = [string]$tmp; break }
        }
        if (-not $nm) { $nm = "DeviceId $did" }

        $deviceList.Add([pscustomobject]@{
            DeviceId   = $did
            DeviceName = $nm
            NodeClass  = $nodeClass
        }) | Out-Null
    }

    Write-Log ("Org devices loaded (Windows only): {0}" -f $deviceList.Count) "INFO"

    # Collect details (OS + Software) in separate lists
    $osDetailRows = New-Object System.Collections.Generic.List[object]
    $swDetailRows = New-Object System.Collections.Generic.List[object]

    $i = 0
    foreach ($dev in $deviceList) {
        $i++
        if (($i % $LogEveryNDevices) -eq 1) {
            Write-Log ("Processing devices: {0}/{1}" -f $i, $deviceList.Count) "INFO"
        }

        # ---------------- OS patch installs ----------------
        try {
            $osResp = Get-DevicePatchInstalls -DeviceId $dev.DeviceId -EndpointSuffix "os-patch-installs" `
                -AfterEpoch $afterEpoch -BeforeEpoch $beforeEpoch -Status $PatchStatus -PageSize $PageSize

            $osItems = @()
            if ($osResp -and (Has-Prop $osResp 'items')) { $osItems = @(Normalize-ToArray $osResp.items) }
            elseif ($osResp) { $osItems = @(Normalize-ToArray $osResp) }

            foreach ($pi in $osItems) {
                if (-not $pi) { continue }

                $pname = [string](Get-Prop $pi 'name')
                if ($ExcludeDefenderIntel -and $pname -like "*Security Intelligence Update for Microsoft Defender Antivirus*") { continue }

                $osDetailRows.Add([pscustomobject]@{
                    PatchName   = $pname
                    KBNumber    = [string](Get-Prop $pi 'kbNumber')
                    Device      = [string]$dev.DeviceName
                    InstalledAt = (Convert-EpochToLocalStringSafe (Get-Prop $pi 'installedAt'))
                    DeviceId    = [int]$dev.DeviceId  # internal only (not exported)
                }) | Out-Null
            }
        }
        catch {
            if ($ContinueOnDeviceError) {
                Write-Log ("WARN OS patches: deviceId={0} ({1}) - {2}" -f $dev.DeviceId, $dev.DeviceName, $_.Exception.Message) "WARN"
            } else {
                throw
            }
        }

        # -------------- Software patch installs --------------
        try {
            $swResp = Get-DevicePatchInstalls -DeviceId $dev.DeviceId -EndpointSuffix "software-patch-installs" `
                -AfterEpoch $afterEpoch -BeforeEpoch $beforeEpoch -Status $PatchStatus -PageSize $PageSize

            $swItems = @()
            if ($swResp -and (Has-Prop $swResp 'items')) { $swItems = @(Normalize-ToArray $swResp.items) }
            elseif ($swResp) { $swItems = @(Normalize-ToArray $swResp) }

            foreach ($pi in $swItems) {
                if (-not $pi) { continue }

                # IMPORTANT: SW patch payloads often use "title" (not "name")
                $title = $null
                foreach ($k in @('title','name','patchName','productName','displayName')) {
                    $tmp = Get-Prop $pi $k
                    if ($tmp) { $title = [string]$tmp; break }
                }
                if (-not $title) { $title = "(unknown)" }

                $swDetailRows.Add([pscustomobject]@{
                    Title             = $title
                    ProductIdentifier = [string](Get-Prop $pi 'productIdentifier')
                    PatchId           = [string](Get-Prop $pi 'id')
                    Impact            = [string](Get-Prop $pi 'impact')
                    Status            = [string](Get-Prop $pi 'status')
                    Device            = [string]$dev.DeviceName
                    InstalledAt       = (Convert-EpochToLocalStringSafe (Get-Prop $pi 'installedAt'))
                    DeviceId          = [int]$dev.DeviceId  # internal only (not exported)
                }) | Out-Null
            }
        }
        catch {
            if ($ContinueOnDeviceError) {
                Write-Log ("WARN SW patches: deviceId={0} ({1}) - {2}" -f $dev.DeviceId, $dev.DeviceName, $_.Exception.Message) "WARN"
            } else {
                throw
            }
        }
    }

    Write-Log ("OS patch installs collected: {0}" -f $osDetailRows.Count) "INFO"
    Write-Log ("SW patch installs collected: {0}" -f $swDetailRows.Count) "INFO"

    # Sort details (keep DeviceId internally for summary)
    $osDetailSorted = @(
        $osDetailRows |
        Sort-Object @{Expression={ $_.Device }; Descending=$false },
                    @{Expression={ $_.InstalledAt }; Descending=$true },
                    @{Expression={ $_.PatchName }; Descending=$false }
    )
    $swDetailSorted = @(
        $swDetailRows |
        Sort-Object @{Expression={ $_.Device }; Descending=$false },
                    @{Expression={ $_.InstalledAt }; Descending=$true },
                    @{Expression={ $_.Title }; Descending=$false }
    )

    # Build summaries (arrays can be empty)
    $osSummarySorted = Build-SummaryRows -DeviceList $deviceList -DetailSorted $osDetailSorted
    $swSummarySorted = Build-SummaryRows -DeviceList $deviceList -DetailSorted $swDetailSorted

    # Export sheets without DeviceId columns
    $osDetailForXlsx = @($osDetailSorted | Select-Object PatchName, KBNumber, Device, InstalledAt)

    # SW PatchInstalls -> renamed columns
    $swDetailForXlsx = @($swDetailSorted | Select-Object Title, ProductIdentifier, PatchId, Impact, Status, Device, InstalledAt)

    # ---- Create XLSX (4 sheets) ----
    $safeOrg  = ($orgName -replace '[\\/:*?"<>|]+', ' ').Trim()
    $fileName = ("{0} {1} - {2}.xlsx" -f $DocBaseTitle, $mi.LabelIt, $safeOrg)
    $xlsxPath = Join-Path $outDir $fileName

    Write-Log ("Creating XLSX: {0}" -f $xlsxPath) "INFO"
    New-MinimalXlsxFromTables -Sheets ([ordered]@{
        "OS Summary"       = $osSummarySorted
        "OS PatchInstalls" = $osDetailForXlsx
        "SW Summary"       = $swSummarySorted
        "SW PatchInstalls" = $swDetailForXlsx
    }) -Path $xlsxPath

    # ---- Overwrite KB: delete same-name articles, then upload ----
    $kbName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    Write-Log ("Ensuring overwrite in Org KB root. Target name: {0}" -f $kbName) "INFO"

    $existing = @(Get-OrgKbArticlesByName -OrganizationId $Customer -ArticleName $kbName)
    if ($existing.Count -gt 0) {
        Write-Log ("Found {0} existing KB article(s) with same name. Deleting..." -f $existing.Count) "INFO"
        foreach ($e in $existing) {
            $eid = $null
            if (Has-Prop $e 'id') { $eid = [int]$e.id }
            elseif (Has-Prop $e 'articleId') { $eid = [int]$e.articleId }
            if ($eid) { $null = Remove-KbArticleById -Id $eid }
        }
    } else {
        Write-Log "No existing KB article found with same name (or list endpoint not available)." "WARN"
    }

    $baseUrl = Resolve-NinjaBaseUrl $NinjaOneInstance
    $token   = Get-OAuthToken -BaseUrl $baseUrl -ClientId $NinjaOneClientId -ClientSecret $NinjaOneClientSecret

    Write-Log ("Uploading XLSX to Org KB root (orgId={0})" -f $Customer) "INFO"
    $ok = Upload-FileToOrgKbRoot -BaseUrl $baseUrl -AccessToken $token -OrganizationId $Customer -FilePath $xlsxPath

    if ($ok) {
        Write-Log "Done." "INFO"
        exit 0
    } else {
        Write-Log "Upload failed." "ERROR"
        exit 1
    }
}
catch {
    $ln = $null
    $line = $null
    try { $ln = $_.InvocationInfo.ScriptLineNumber; $line = $_.InvocationInfo.Line } catch {}

    if ($ln) { Write-Log ("FAILED at line {0}: {1}" -f $ln, $_.Exception.Message) "ERROR" }
    else     { Write-Log ("FAILED: {0}" -f $_.Exception.Message) "ERROR" }

    if ($line) { Write-Log ("Line: {0}" -f $line.Trim()) "ERROR" }
    exit 1
}

} # process