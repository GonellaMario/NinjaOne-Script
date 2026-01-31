[CmdletBinding()]
param(
    [Parameter()]
    [int]$Customer = 1
)

begin {
    # Ninja Script Form Variable override
    if ($env:customer -and $env:customer -notlike "null") { $Customer = [int]$env:customer }

    if (-not $Customer -or $Customer -lt 1) {
        Write-Host -Object "[Error] Customer (OrganizationId) is required and must be >= 1"
        exit 1
    }
}

process {

# ====== CONFIG ======
$StatusFilter          = "DISK_DRIVE_ADDED"
$PageSize              = 250
$MaxRowsToPublish      = 800
$MessageMaxLen         = 600
$LogEveryNDevices      = 10
$ContinueOnDeviceError = $true

$DocBaseTitle = "Removable Media Activity Report"
$outDir       = "C:\ProgramData\NinjaRMMAgent\scripting\Reports"
New-Item -Path $outDir -ItemType Directory -Force | Out-Null
# ====================

# -------------------- PS 7+ bootstrap --------------------
if ($PSVersionTable.PSVersion.Major -lt 7) {
    try {
        if (!(Test-Path "$env:SystemDrive\Program Files\PowerShell\7")) {
            Write-Output 'Does not appear PowerShell 7 is installed'
            exit 1
        }
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User')
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
    param([object]$obj)
    if ($null -eq $obj) { return @() }
    if ($obj -is [string]) { return @($obj) }
    if ($obj -is [System.Array]) { return $obj }

    foreach ($k in @('items','data','results','documents','activities','organizations','devices','folders','articles')) {
        if ($obj.PSObject -and $obj.PSObject.Properties.Match($k).Count -gt 0 -and $obj.$k) {
            return (Normalize-ToArray $obj.$k)
        }
    }
    return @($obj)
}

function Convert-ActivityTimeToEpoch([object]$t) {
    if ($null -eq $t) { return $null }
    if ($t -is [double]) { return [int64][math]::Floor($t) }
    if ($t -is [int64]) { return $t }
    if ($t -is [int]) { return [int64]$t }
    return $null
}
function Convert-EpochToLocalString([int64]$ep) {
    return ([DateTimeOffset]::FromUnixTimeSeconds($ep).ToLocalTime().DateTime).ToString("dd/MM/yyyy HH:mm:ss")
}

function Get-PreviousMonthRange {
    $now = Get-Date
    $firstThisMonth = Get-Date -Year $now.Year -Month $now.Month -Day 1 -Hour 0 -Minute 0 -Second 0
    $endPrevMonth = $firstThisMonth.AddSeconds(-1)
    $startPrevMonth = Get-Date -Year $endPrevMonth.Year -Month $endPrevMonth.Month -Day 1 -Hour 0 -Minute 0 -Second 0
    return [pscustomobject]@{
        Start  = $startPrevMonth
        End    = $endPrevMonth
        Yyyymm = $endPrevMonth.ToString("yyyyMM")
        Label  = $endPrevMonth.ToString("MMMM yyyy")
    }
}

function Try-ParseInstallDateUtc([object]$installDate) {
    if (-not $installDate) { return $null }
    try { return [DateTimeOffset]::Parse([string]$installDate).ToUniversalTime() } catch { return $null }
}

function Get-ParamsSafe($activity) {
    $data = Get-Prop $activity 'data'
    if (-not $data) { return $null }
    $msg  = Get-Prop $data 'message'
    if (-not $msg) { return $null }
    $prm  = Get-Prop $msg 'params'
    if (-not $prm) { return $null }
    return $prm
}

function Parse-FromMessage {
    param([string]$Message)

    $result = [ordered]@{
        disk_name       = $null
        capacity        = $null
        install_date    = $null
        disk_interface  = $null
        disk_model      = $null
        disk_media_type = $null
    }

    if (-not $Message) { return [pscustomobject]$result }

    # Disk name: Italian + English patterns
    if ($Message -match "(Unità disco aggiunta|Disk drive added):\s*'(?<disk>[^']+)'") {
        $result.disk_name = $Matches.disk
    }

    # Capacity: Italian + English patterns
    if ($Message -match "(capacit[aà]|capacity):\s*'(?<cap>\d+)'") {
        $result.capacity = $Matches.cap
    }

    # Install date: Italian + English patterns
    if ($Message -match "(data installazione|install date):\s*'(?<dt>[^']+)'") {
        $result.install_date = $Matches.dt
    }

    # Additional info: Italian + English patterns
    if ($Message -match "(informazioni aggiuntive|additional information):\s*\(\s*(?<iface>[^,]+)\s*,\s*(?<model>[^,]+)\s*,\s*(?<media>[^\)]+)\s*\)") {
        $result.disk_interface  = $Matches.iface.Trim()
        $result.disk_model      = $Matches.model.Trim()
        $result.disk_media_type = $Matches.media.Trim()
    }

    return [pscustomobject]$result
}

function Get-EventEpoch {
    param([object]$activity)

    $params = Get-ParamsSafe $activity
    if ($params) {
        $idate = Get-Prop $params 'install_date'
        $dto = Try-ParseInstallDateUtc $idate
        if ($dto) { return $dto.ToUnixTimeSeconds() }
    }

    $msgText = [string](Get-Prop $activity 'message')
    if ($msgText) {
        $parsed = Parse-FromMessage -Message $msgText
        $dto2 = Try-ParseInstallDateUtc $parsed.install_date
        if ($dto2) { return $dto2.ToUnixTimeSeconds() }
    }

    return (Convert-ActivityTimeToEpoch (Get-Prop $activity 'activityTime'))
}

function In-Range([int64]$epoch, [int64]$AfterEpoch, [int64]$BeforeEpoch) {
    if ($epoch -le 0) { return $false }
    return ($epoch -ge $AfterEpoch -and $epoch -le $BeforeEpoch)
}

function Get-DeviceActivitiesAddedSinglePage {
    param([Parameter(Mandatory)][int]$DeviceId)

    $path = ("device/{0}/activities?status={1}&pageSize={2}" -f $DeviceId, $StatusFilter, $PageSize)
    $resp = Invoke-NinjaOneRequest -Method GET -Path $path

    if (Has-Prop $resp 'activities') { return (Normalize-ToArray $resp.activities) }
    return (Normalize-ToArray $resp)
}

function New-MinimalXlsxFromObjects {
    param(
        [Parameter(Mandatory=$true)][object[]]$Rows,
        [Parameter(Mandatory=$true)][string]$Path,
        [string]$SheetName = "Report"
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $headers = @()
    if ($Rows.Count -gt 0) { $headers = $Rows[0].PSObject.Properties.Name }

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
    function CellInlineStr([string]$r, [string]$text) {
        $t = Escape-Xml $text
        return "<c r=`"$r`" t=`"inlineStr`"><is><t>$t</t></is></c>"
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.AppendLine('<sheetData>')

    $rowIndex = 1
    [void]$sb.Append("<row r=`"$rowIndex`">")
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $col = ColName ($c + 1)
        $ref = "$col$rowIndex"
        [void]$sb.Append( (CellInlineStr $ref ([string]$headers[$c])) )
    }
    [void]$sb.AppendLine("</row>")

    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $rowIndex = $i + 2
        [void]$sb.Append("<row r=`"$rowIndex`">")
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $col = ColName ($c + 1)
            $ref = "$col$rowIndex"
            $val = $null
            try { $val = $Rows[$i].PSObject.Properties[$headers[$c]].Value } catch {}
            [void]$sb.Append( (CellInlineStr $ref ([string]$val)) )
        }
        [void]$sb.AppendLine("</row>")
    }

    [void]$sb.AppendLine('</sheetData>')
    [void]$sb.AppendLine('</worksheet>')
    $sheetXml = $sb.ToString()

    $workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="$(Escape-Xml $SheetName)" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"@

    $relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@

    $workbookRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"@

    $contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"@

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
            Add-ZipEntry $zip "xl/worksheets/sheet1.xml" $sheetXml
        }
        finally { $zip.Dispose() }
    }
    finally { $fs.Dispose() }
}

function Resolve-NinjaBaseUrl([string]$InstanceValue) {
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

# ----------- KB helpers (overwrite: delete old then upload) -----------

function Invoke-KbRequest {
    param(
        [Parameter(Mandatory=$true)][string]$BaseUrl,
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$Method,
        [Parameter(Mandatory=$true)][string]$PathAndQuery
    )
    $headers = @{ Authorization = "Bearer $AccessToken"; Accept = "application/json" }
    $uri = ($BaseUrl.TrimEnd('/') + $PathAndQuery)
    return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers
}

function Get-OrgKbArticles {
    param(
        [Parameter(Mandatory=$true)][string]$BaseUrl,
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][int]$OrganizationId
    )

    $paths = @(
        "/v2/knowledgebase/articles?organizationId=$OrganizationId",
        "/api/v2/knowledgebase/articles?organizationId=$OrganizationId"
    )

    $lastErr = $null
    foreach ($p in $paths) {
        try {
            Write-Log ("KB list try: {0}" -f $p) "DEBUG"
            $resp = Invoke-KbRequest -BaseUrl $BaseUrl -AccessToken $AccessToken -Method "GET" -PathAndQuery $p
            return @(Normalize-ToArray $resp)
        } catch {
            $lastErr = $_.Exception.Message
        }
    }

    Write-Log ("KB list failed on all endpoints: {0}" -f $lastErr) "WARN"
    return @()
}

function Get-ArticleIdAny {
    param([object]$a)
    foreach ($k in @('id','articleId','knowledgeBaseArticleId','kbArticleId')) {
        $v = Get-Prop $a $k
        if ($v) { return [string]$v }
    }
    return $null
}

function Get-ArticleNameAny {
    param([object]$a)
    foreach ($k in @('name','title','documentName','fileName')) {
        $v = Get-Prop $a $k
        if ($v) { return [string]$v }
    }
    return $null
}

function Remove-OrgKbArticleById {
    param(
        [Parameter(Mandatory=$true)][string]$BaseUrl,
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$ArticleId
    )

    $paths = @(
        "/v2/knowledgebase/articles/$ArticleId",
        "/api/v2/knowledgebase/articles/$ArticleId"
    )

    $lastErr = $null
    foreach ($p in $paths) {
        try {
            Write-Log ("KB delete try: {0}" -f $p) "DEBUG"
            $null = Invoke-KbRequest -BaseUrl $BaseUrl -AccessToken $AccessToken -Method "DELETE" -PathAndQuery $p
            return $true
        } catch {
            $lastErr = $_.Exception.Message
        }
    }

    Write-Log ("KB delete failed for articleId={0}: {1}" -f $ArticleId, $lastErr) "WARN"
    return $false
}

function Cleanup-OrgKbDuplicatesByTitlePrefix {
    param(
        [Parameter(Mandatory=$true)][string]$BaseUrl,
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][int]$OrganizationId,
        [Parameter(Mandatory=$true)][string]$TitlePrefix
    )

    $all = @(Get-OrgKbArticles -BaseUrl $BaseUrl -AccessToken $AccessToken -OrganizationId $OrganizationId)
    if (-not $all -or $all.Count -eq 0) {
        Write-Log "KB list empty (nothing to cleanup)." "INFO"
        return 0
    }

    # Match the exact title OR prefixed duplicates like "(1)", "(2)", etc.
    $matches = New-Object System.Collections.Generic.List[object]
    foreach ($a in $all) {
        $nm = Get-ArticleNameAny $a
        if (-not $nm) { continue }

        if ($nm -eq $TitlePrefix -or $nm -like ("$TitlePrefix*")) {
            # Accept also "Title (1)", "Title (2)", etc.
            $matches.Add($a) | Out-Null
        }
    }

    if ($matches.Count -eq 0) {
        Write-Log ("No existing KB items found matching '{0}'" -f $TitlePrefix) "INFO"
        return 0
    }

    Write-Log ("Found {0} existing KB item(s) matching '{1}' -> deleting to overwrite." -f $matches.Count, $TitlePrefix) "INFO"

    $deleted = 0
    foreach ($m in $matches) {
        $id = Get-ArticleIdAny $m
        $nm = Get-ArticleNameAny $m
        if (-not $id) {
            Write-Log ("Skip delete (no id) for '{0}'" -f $nm) "WARN"
            continue
        }
        if (Remove-OrgKbArticleById -BaseUrl $BaseUrl -AccessToken $AccessToken -ArticleId $id) {
            Write-Log ("Deleted KB item: {0} (id={1})" -f $nm, $id) "INFO"
            $deleted++
        } else {
            Write-Log ("Unable to delete KB item: {0} (id={1})" -f $nm, $id) "WARN"
        }
    }

    return $deleted
}

function Upload-FileToOrgKbRoot {
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

    # ORG KB ROOT: pass ONLY organizationId (no folderId, no folderPath)
    $multipart.Add((New-Object System.Net.Http.StringContent($OrganizationId.ToString())), "organizationId")

    $stream = [System.IO.File]::OpenRead($FilePath)
    try {
        $fileContent = New-Object System.Net.Http.StreamContent($stream)
        $fileContent.Headers.ContentType =
            [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/octet-stream")

        # IMPORTANT: field name is "files" (plural)
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

# -------------------- Module load + Connect --------------------
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

try {
    $range = Get-PreviousMonthRange
    $afterEpoch  = [DateTimeOffset]::new($range.Start.ToUniversalTime()).ToUnixTimeSeconds()
    $beforeEpoch = [DateTimeOffset]::new($range.End.ToUniversalTime()).ToUnixTimeSeconds()

    Write-Log ("Range (prev month): {0} -> {1}" -f $range.Start, $range.End) "INFO"
    Write-Log ("Customer/OrganizationId: {0}" -f $Customer) "INFO"

    # Org name
    $orgName = ("OrgId {0}" -f $Customer)
    try {
        $orgs = Invoke-NinjaOneRequest -Method GET -Path "organizations"
        $o = (Normalize-ToArray $orgs) | Where-Object { $_ -and (Get-Prop $_ 'id') -and [int]$_.id -eq $Customer } | Select-Object -First 1
        if ($o -and (Get-Prop $o 'name')) { $orgName = $o.name }
    } catch {}
    Write-Log ("Organization: {0}" -f $orgName) "INFO"

    # Devices
    Write-Log "Fetching organization devices..." "INFO"
    $orgDevices = Normalize-ToArray (Invoke-NinjaOneRequest -Method GET -Path ("organization/{0}/devices" -f $Customer))

    $deviceList = New-Object System.Collections.Generic.List[object]
    foreach ($d in $orgDevices) {
        $idVal = Get-Prop $d 'id'
        if (-not $idVal) { continue }
        $id = [int]$idVal

        $nm = $null
        foreach ($k in @('systemName','hostname','displayName','name')) {
            $tmp = Get-Prop $d $k
            if ($tmp) { $nm = [string]$tmp; break }
        }
        if (-not $nm) { $nm = "DeviceId $id" }
        $deviceList.Add([pscustomobject]@{ id=$id; name=$nm }) | Out-Null
    }
    Write-Log ("Org devices loaded: {0}" -f $deviceList.Count) "INFO"

    # Collect activities (single page per device)
    $all = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $devicesPossiblyTruncated = 0

    $i = 0
    foreach ($dev in $deviceList) {
        $i++
        if (($i % $LogEveryNDevices) -eq 1) {
            Write-Log ("Fetching device activities: device {0}/{1} (deviceId={2})" -f $i, $deviceList.Count, $dev.id) "INFO"
        }

        try {
            $actsRaw = Get-DeviceActivitiesAddedSinglePage -DeviceId $dev.id
            $acts    = @(Normalize-ToArray $actsRaw)

            if (@($acts).Count -eq $PageSize) { $devicesPossiblyTruncated++ }

            foreach ($a in $acts) {
                $ep = Get-EventEpoch -activity $a
                if (-not $ep) { continue }
                if (-not (In-Range -epoch $ep -AfterEpoch $afterEpoch -BeforeEpoch $beforeEpoch)) { continue }

                $aid = Get-Prop $a 'id'
                if ($aid) {
                    if ($seen.Add([string]$aid)) { $all.Add($a) | Out-Null }
                } else {
                    $all.Add($a) | Out-Null
                }
            }
        }
        catch {
            if ($ContinueOnDeviceError) {
                Write-Log ("WARN deviceId={0} ({1}) skipped: {2}" -f $dev.id, $dev.name, $_.Exception.Message) "WARN"
                continue
            }
            throw
        }
    }

    Write-Log ("Total activities collected (deduped, in-range): {0}" -f $all.Count) "INFO"

    # Build rows
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($a in $all) {
        $did = [int](Get-Prop $a 'deviceId')

        $devObj = ($deviceList | Where-Object { $_.id -eq $did } | Select-Object -First 1)
        $devName = if ($devObj) { $devObj.name } else { "DeviceId $did" }

        $params  = Get-ParamsSafe $a
        $msgText = [string](Get-Prop $a 'message')

        $parsed = $null
        if (-not $params) { $parsed = Parse-FromMessage -Message $msgText }

        $diskName      = $null
        $diskModel     = $null
        $diskInterface = $null
        $mediaType     = $null

        if ($params) {
            $diskName      = [string](Get-Prop $params 'disk_name')
            $diskModel     = [string](Get-Prop $params 'disk_model')
            $diskInterface = [string](Get-Prop $params 'disk_interface')
            $mediaType     = [string](Get-Prop $params 'disk_media_type')
        } elseif ($parsed) {
            $diskName      = [string]$parsed.disk_name
            $diskModel     = [string]$parsed.disk_model
            $diskInterface = [string]$parsed.disk_interface
            $mediaType     = [string]$parsed.disk_media_type
        }

        $eventEpoch = Get-EventEpoch -activity $a
        $eventTime  = if ($eventEpoch) { Convert-EpochToLocalString $eventEpoch } else { "" }

        if (-not $diskName)      { $diskName      = "-" }
        if (-not $diskModel) {
            $m = $msgText
            if ($m -and $m.Length -gt $MessageMaxLen) { $m = $m.Substring(0,$MessageMaxLen) + "..." }
            $diskModel = if ($m) { $m } else { "-" }
        }
        if (-not $diskInterface) { $diskInterface = "-" }
        if (-not $mediaType)     { $mediaType     = "-" }

        $rows.Add([pscustomobject]@{
            EventTime     = $eventTime
            Device        = $devName
            DiskName      = $diskName
            DiskModel     = $diskModel
            DiskInterface = $diskInterface
            MediaType     = $mediaType
        }) | Out-Null
    }

    $sorted  = @($rows | Sort-Object -Property EventTime -Descending)
    $publish = $sorted
    if (@($sorted).Count -gt $MaxRowsToPublish) { $publish = @($sorted | Select-Object -First $MaxRowsToPublish) }

    Write-Log ("Rows total={0}, exporting={1}" -f @($sorted).Count, @($publish).Count) "INFO"

    # Create XLSX
    $safeOrg = ($orgName -replace '[\\/:*?"<>|]+', ' ').Trim()
    $titlePrefix = ("{0} {1} - {2}" -f $DocBaseTitle, $range.Label, $safeOrg)
    $fileName = ("{0}.xlsx" -f $titlePrefix)
    $xlsxPath = Join-Path $outDir $fileName

    New-MinimalXlsxFromObjects -Rows $publish -Path $xlsxPath -SheetName "RemovableMedia"
    Write-Log ("XLSX created: {0}" -f $xlsxPath) "INFO"

    # Upload to ORG KB ROOT (OVERWRITE = delete old first)
    $baseUrl = Resolve-NinjaBaseUrl $NinjaOneInstance
    $token   = Get-OAuthToken -BaseUrl $baseUrl -ClientId $NinjaOneClientId -ClientSecret $NinjaOneClientSecret

    # Cleanup duplicates (same title / title (1) etc.)
    $deletedCount = Cleanup-OrgKbDuplicatesByTitlePrefix -BaseUrl $baseUrl -AccessToken $token -OrganizationId $Customer -TitlePrefix $titlePrefix
    Write-Log ("Cleanup completed. Deleted: {0}" -f $deletedCount) "INFO"

    Write-Log ("Uploading XLSX to ORG KB ROOT (orgId={0})" -f $Customer) "INFO"
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

} # end process