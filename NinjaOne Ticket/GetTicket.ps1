[CmdletBinding()]
param(
    [Parameter()]
    [string]$CsvFileName = "NinjaOne_Tickets.csv",

    [Parameter()]
    [string]$OutDir = "C:\NinjaOneTicket",

    # FORZATO: Tutti i Ticket = 2
    [Parameter()]
    [int]$BoardId = 2
)

begin {
    if ($env:csvFileName -and $env:csvFileName -notlike "null") { $CsvFileName = [string]$env:csvFileName }
    if ($env:outDir -and $env:outDir -notlike "null") { $OutDir = [string]$env:outDir }
    if ($env:boardId -and $env:boardId -notlike "null") { $BoardId = [int]$env:boardId }
}

process {

# ==================== CONFIG ====================
$OrgNameRegexInclude    = '\(\d+\)\s*$'
$OrgNameAlwaysInclude   = @("Ninja Laboratory")
$OrgNameAlwaysExclude   = @("VECTOR S.P.A. (36638)")

$PageSize               = 500
$LogEveryNPages         = 1

$OAuthScopes            = "monitoring management control"
$OAuthRedirectUri       = "http://localhost"

# OPEN/Aperto filter
$OpenStatusIds          = @(2000)
$OpenStatusNames        = @("Aperto","Open")
# ===============================================

New-Item -Path $OutDir -ItemType Directory -Force | Out-Null
$CsvPath   = Join-Path $OutDir $CsvFileName
$StatePath = Join-Path $OutDir "NinjaOne_Tickets.state.json"

# -------------------- PS 7+ bootstrap (SAFE) --------------------
if ($PSVersionTable.PSVersion.Major -lt 7) {
    try {
        $pwsh = $null
        if (Test-Path "$env:SystemDrive\Program Files\PowerShell\7\pwsh.exe") {
            $pwsh = "$env:SystemDrive\Program Files\PowerShell\7\pwsh.exe"
        } elseif (Get-Command pwsh -ErrorAction SilentlyContinue) {
            $pwsh = "pwsh"
        }

        if (-not $pwsh) {
            Write-Output "PowerShell 7 not found. Install PowerShell 7 and retry."
            exit 1
        }

        & $pwsh -File $PSCommandPath @PSBoundParameters
        exit $LASTEXITCODE
    } catch {
        Write-Output "Failed to relaunch with PowerShell 7: $($_.Exception.Message)"
        exit 1
    }
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

    foreach ($k in @('items','data','results','organizations','tickets','boards','value')) {
        if ($obj.PSObject -and $obj.PSObject.Properties.Match($k).Count -gt 0 -and $obj.$k) {
            return (Normalize-ToArray $obj.$k)
        }
    }
    return @($obj)
}

function Get-EnvVar([string]$name) {
    try { return [Environment]::GetEnvironmentVariable($name) } catch { return $null }
}
function Get-EnvOrValue([string]$envName, $fallback) {
    $v = Get-EnvVar $envName
    if ($v -and $v -notlike "null") { return $v }
    return $fallback
}
function Normalize-Instance([string]$inst) {
    if (-not $inst) { return $null }
    $inst = $inst.Trim()
    if ($inst -match '^https?://') { return $inst.TrimEnd('/') }
    return ("https://{0}" -f $inst.TrimEnd('/'))
}

function Normalize-OrgName([string]$name) {
    if (-not $name) { return "" }
    $n = $name.Trim()
    # comprime spazi multipli
    $n = ($n -replace '\s{2,}', ' ')
    return $n
}

# ---- createTime normalizer: restituisce millisecondi (Int64) ----
function Get-EpochMs {
    param([object]$value)

    if ($null -eq $value) { return 0L }

    if ($value -is [int] -or $value -is [long]) {
        $n = [int64]$value
        if ($n -gt 1000000000000L) { return $n }              # ms
        if ($n -gt 1000000000L) { return ($n * 1000L) }       # sec -> ms
        return $n
    }

    $s = [string]$value
    $s = $s.Trim()
    if (-not $s) { return 0L }

    $s2 = $s.Replace(',', '.')
    $d2 = 0.0
    if ([double]::TryParse($s2, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d2)) {
        if ($d2 -gt 1000000000000.0) { return [int64][Math]::Floor($d2) }
        if ($d2 -gt 1000000000.0) { return [int64][Math]::Floor($d2 * 1000.0) }
        return [int64][Math]::Floor($d2)
    }

    return 0L
}

function EpochMs-ToUtcIso([int64]$ms) {
    if ($ms -le 0) { return "" }
    return ([DateTimeOffset]::FromUnixTimeMilliseconds($ms).UtcDateTime).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
}

function Get-TicketOrgName {
    param($t)

    foreach ($k in @("organization","org","client","customer")) {
        $v = Get-Prop $t $k
        if ($null -eq $v) { continue }

        if ($v -is [string]) {
            if ($v.Trim()) { return (Normalize-OrgName $v) }
        } else {
            $n = Get-Prop $v "name"
            if ($n) { return (Normalize-OrgName ([string]$n)) }
        }
    }
    return ""
}

function Get-TicketOrgId {
    param($t)

    foreach ($k in @("organizationId","orgId","orgID","clientId","clientID","customerId")) {
        $v = Get-Prop $t $k
        if ($null -ne $v -and [string]$v -ne "") { return [int]$v }
    }

    $o = Get-Prop $t "organization"
    if ($o -and -not ($o -is [string])) {
        $id = Get-Prop $o "id"
        if ($id) { return [int]$id }
    }

    return $null
}

function Org-IsAllowed {
    param([string]$orgName)

    if (-not $orgName) { return $false }
    if ($OrgNameAlwaysExclude -contains $orgName) { return $false }
    if ($OrgNameAlwaysInclude -contains $orgName) { return $true }
    return ($orgName -match $OrgNameRegexInclude)
}

function Ticket-IsOpen {
    param($t)

    $st = Get-Prop $t "status"
    if ($st) {
        $sid = Get-Prop $st "statusId"
        $dn  = Get-Prop $st "displayName"
        if ($sid -and ($OpenStatusIds -contains [int]$sid)) { return $true }
        if ($dn -and ($OpenStatusNames -contains ([string]$dn))) { return $true }
    }

    $st3 = Get-Prop $t "status"
    if ($st3 -is [string] -and ($OpenStatusNames -contains ([string]$st3))) { return $true }

    $st2 = Get-Prop $t "state"
    if ($st2 -and ($OpenStatusNames -contains ([string]$st2))) { return $true }

    return $false
}

function Get-NinjaAccessToken {
    param(
        [Parameter(Mandatory)][string]$InstanceBase,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter(Mandatory)][string]$Scope,
        [Parameter(Mandatory)][string]$RedirectUri
    )

    $tokenPaths = @("/oauth/token", "/ws/oauth/token")
    $lastErr = $null

    foreach ($tp in $tokenPaths) {
        try {
            $uri = $InstanceBase.TrimEnd('/') + $tp
            $body = @{
                grant_type    = "client_credentials"
                client_id     = $ClientId
                client_secret = $ClientSecret
                scope         = $Scope
                redirect_uri  = $RedirectUri
            }

            $resp = Invoke-RestMethod -Method POST -Uri $uri -Body $body -ContentType "application/x-www-form-urlencoded" -Headers @{ "Accept"="application/json" }
            if ($resp -and $resp.access_token) { return [string]$resp.access_token }
        } catch {
            $lastErr = $_.Exception.Message
        }
    }

    throw "Unable to fetch access token. Last error: $lastErr"
}

function Invoke-NinjaApi {
    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string[]]$ApiBases,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$AccessToken,
        [object]$Body = $null
    )

    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Accept"        = "application/json"
    }

    $lastErr = $null
    foreach ($base in $ApiBases) {
        try {
            $uri = $base.TrimEnd('/') + $Path
            if ($null -ne $Body) {
                $headers["Content-Type"] = "application/json"
                $json = ($Body | ConvertTo-Json -Depth 30)
                return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -Body $json
            } else {
                return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers
            }
        } catch {
            $lastErr = $_.Exception.Message
        }
    }

    throw "All API base variants failed for $Method $Path. Last error: $lastErr"
}

# -------------------- Credenziali --------------------
$instanceProp = $null
$clientIdProp = $null
$secretProp   = $null
try {
    if (Get-Command -Name "Ninja-Property-Get" -ErrorAction SilentlyContinue) {
        $instanceProp = Ninja-Property-Get ninjaoneInstance
        $clientIdProp = Ninja-Property-Get ninjaoneClientId
        $secretProp   = Ninja-Property-Get ninjaoneClientSecret
    }
} catch {}

$instanceRaw = Get-EnvOrValue "ninjaoneInstance" $instanceProp
$clientIdRaw = Get-EnvOrValue "ninjaoneClientId" $clientIdProp
$secretRaw   = Get-EnvOrValue "ninjaoneClientSecret" $secretProp

$NinjaOneInstance     = Normalize-Instance ([string]$instanceRaw)
$NinjaOneClientId     = [string]$clientIdRaw
$NinjaOneClientSecret = [string]$secretRaw

if (-not $NinjaOneInstance -or -not $NinjaOneClientId -or -not $NinjaOneClientSecret) {
    Write-Log "Missing required API credentials (ninjaoneInstance / ninjaoneClientId / ninjaoneClientSecret)." "ERROR"
    exit 1
}

$apiBases = @(
    ($NinjaOneInstance.TrimEnd('/') + "/api/v2"),
    ($NinjaOneInstance.TrimEnd('/') + "/ws/api/v2")
) | Select-Object -Unique

Write-Log ("Fetching token (instance={0})..." -f $NinjaOneInstance) "INFO"
$AccessToken = Get-NinjaAccessToken -InstanceBase $NinjaOneInstance -ClientId $NinjaOneClientId -ClientSecret $NinjaOneClientSecret -Scope $OAuthScopes -RedirectUri $OAuthRedirectUri
Write-Log "Fetched token." "INFO"

# -------------------- Organizations lookup (Name -> Id) --------------------
Write-Log "Fetching organizations for lookup..." "INFO"
$orgsRaw = Invoke-NinjaApi -Method "GET" -ApiBases $apiBases -Path "/organizations" -AccessToken $AccessToken
$orgs = @(Normalize-ToArray $orgsRaw)

$OrgNameToId = @{}
foreach ($o in $orgs) {
    $id = Get-Prop $o "id"
    $name = Get-Prop $o "name"
    if ($id -eq $null -or -not $name) { continue }
    $nn = Normalize-OrgName ([string]$name)
    if (-not $OrgNameToId.ContainsKey($nn)) {
        $OrgNameToId[$nn] = [int]$id
    }
}
Write-Log ("Organizations loaded for lookup: {0}" -f $OrgNameToId.Count) "INFO"

# -------------------- Load State (ms) --------------------
$state = @{
    lastCreateTimeMs = 0L
    idsAtLastTime    = @()
}
if (Test-Path $StatePath) {
    try {
        $loaded = Get-Content $StatePath -Raw | ConvertFrom-Json
        if ($loaded.lastCreateTimeMs -ne $null) { $state.lastCreateTimeMs = [int64]$loaded.lastCreateTimeMs }
        if ($loaded.idsAtLastTime) { $state.idsAtLastTime = @($loaded.idsAtLastTime) }
        Write-Log ("State loaded: lastCreateTimeMs={0}, idsAtLastTime={1}" -f $state.lastCreateTimeMs, $state.idsAtLastTime.Count) "INFO"
    } catch {
        Write-Log ("State file unreadable, starting from 0. ({0})" -f $StatePath) "WARN"
    }
} else {
    Write-Log "No state file found, starting from 0." "INFO"
}

Write-Log ("Board forced: id={0}" -f $BoardId) "INFO"
Write-Log ("Incremental filter: createTimeMs > {0}" -f $state.lastCreateTimeMs) "INFO"

# -------------------- Fetch tickets (board run) --------------------
$candidates = New-Object System.Collections.Generic.List[object]
$cursor = 0L
$page = 0

while ($true) {
    $page++
    if (($page % $LogEveryNPages) -eq 0) {
        Write-Log ("Fetching tickets page {0} (cursor={1})..." -f $page, $cursor) "INFO"
    }

    $body = @{
        sortBy = @(@{ field = "createTime"; direction = "DESC" })
        pageSize = $PageSize
        lastCursorId = $cursor
    }

    $resp = $null
    try {
        $resp = Invoke-NinjaApi -Method "POST" -ApiBases $apiBases -Path ("/ticketing/trigger/board/{0}/run" -f $BoardId) -AccessToken $AccessToken -Body $body
    } catch {
        $resp = Invoke-NinjaApi -Method "POST" -ApiBases $apiBases -Path ("/ticketing/trigger/boards/{0}/run" -f $BoardId) -AccessToken $AccessToken -Body $body
    }

    $data = @()
    if (Has-Prop $resp "data") { $data = @(Normalize-ToArray $resp.data) }
    else { $data = @(Normalize-ToArray $resp) }

    if (-not $data -or $data.Count -eq 0) { break }

    $oldestMs = Get-EpochMs (Get-Prop $data[-1] "createTime")
    $mayStopAfterThisPage = ($oldestMs -lt [int64]$state.lastCreateTimeMs)

    foreach ($t in $data) {
        $ticketId = [int](Get-Prop $t "id")
        $ctMs     = Get-EpochMs (Get-Prop $t "createTime")

        if ($ctMs -lt [int64]$state.lastCreateTimeMs) { continue }
        if ($ctMs -eq [int64]$state.lastCreateTimeMs -and ($state.idsAtLastTime -contains $ticketId)) { continue }

        if (-not (Ticket-IsOpen $t)) { continue }

        $orgName = Get-TicketOrgName $t
        if (-not (Org-IsAllowed $orgName)) { continue }

        [void]$candidates.Add($t)
    }

    $nextCursor = $null
    if (Has-Prop $resp "metadata" -and (Has-Prop $resp.metadata "lastCursorId")) {
        $nextCursor = [int64]$resp.metadata.lastCursorId
    }
    if (-not $nextCursor -or $nextCursor -eq $cursor) { break }
    $cursor = $nextCursor

    if ($mayStopAfterThisPage) { break }
}

if ($candidates.Count -eq 0) {
    Write-Log "No new OPEN tickets to export (after filters)." "INFO"
    exit 0
}

Write-Log ("Tickets to export: {0}" -f $candidates.Count) "INFO"

# -------------------- Build rows & export (orgId lookup by name) --------------------
$rows = foreach ($t in $candidates) {
    $ticketId = [int](Get-Prop $t "id")
    $ctMs     = Get-EpochMs (Get-Prop $t "createTime")
    $utMs     = Get-EpochMs (Get-Prop $t "lastUpdated")

    $orgName  = Get-TicketOrgName $t
    $orgId    = Get-TicketOrgId $t

    # Se manca orgId nel ticket -> lookup dal nome org
    if ($null -eq $orgId -and $orgName) {
        $nn = Normalize-OrgName $orgName
        if ($OrgNameToId.ContainsKey($nn)) {
            $orgId = [int]$OrgNameToId[$nn]
        }
    }

    $subject = $null
    foreach ($k in @("subject","title","summary")) {
        $v = Get-Prop $t $k
        if ($v) { $subject = [string]$v; break }
    }

    $statusName = ""
    $statusId   = ""
    $st = Get-Prop $t "status"
    if ($st) {
        $statusId = [string](Get-Prop $st "statusId")
        $statusName = [string](Get-Prop $st "displayName")
    } elseif ((Get-Prop $t "status") -is [string]) {
        $statusName = [string](Get-Prop $t "status")
    }

    [PSCustomObject]@{
        ExportedAtUtc      = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        TicketId           = $ticketId
        OrganizationId     = $orgId
        OrganizationName   = $orgName
        StatusId           = $statusId
        Status             = $statusName
        Subject            = $subject
        CreateTimeMs       = $ctMs
        CreateTimeUtc      = (EpochMs-ToUtcIso $ctMs)
        LastUpdatedMs      = $utMs
        LastUpdatedUtc     = (EpochMs-ToUtcIso $utMs)
        RawJson            = ($t | ConvertTo-Json -Depth 30 -Compress)
    }
}

$rowsSorted = $rows | Sort-Object -Property CreateTimeMs, TicketId
if (Test-Path $CsvPath) { $rowsSorted | Export-Csv -Path $CsvPath -Append -NoTypeInformation -Encoding UTF8 }
else { $rowsSorted | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 }

# -------------------- Save state (ms) --------------------
$maxCtMs = ($rowsSorted | Measure-Object -Property CreateTimeMs -Maximum).Maximum
$idsAtMax = @(
    $rowsSorted |
    Where-Object { [int64]$_.CreateTimeMs -eq [int64]$maxCtMs } |
    ForEach-Object { [int]$_.TicketId }
)

$newState = @{
    lastCreateTimeMs = [int64]$maxCtMs
    idsAtLastTime    = $idsAtMax
}
($newState | ConvertTo-Json -Depth 5) | Set-Content -Path $StatePath -Encoding UTF8

Write-Log ("Export complete: {0} tickets appended to {1}" -f $rowsSorted.Count, $CsvPath) "INFO"
Write-Log ("State updated: lastCreateTimeMs={0} (ids at max={1})" -f $newState.lastCreateTimeMs, $newState.idsAtLastTime.Count) "INFO"

} # end process