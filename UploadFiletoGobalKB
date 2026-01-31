<#
.SYNOPSIS
  Upload a file to the NinjaOne *Global* Knowledge Base into a specific folder (folderId).

.DESCRIPTION
  Authenticates via OAuth2 Client Credentials and uploads a file to the NinjaOne Knowledge Base
  using multipart/form-data.

  Compatible with Windows PowerShell 5.1 (no Invoke-RestMethod -Form), so it uses .NET HttpClient
  and MultipartFormDataContent.

.PREREQUISITES
  - Windows PowerShell 5.1
  - A NinjaOne API application (Client ID / Client Secret) with permission to manage Knowledge Base
  - Set environment variables:
      NINJA_CLIENT_ID
      NINJA_CLIENT_SECRET

  Example (current session only):
      $env:NINJA_CLIENT_ID     = "xxxx"
      $env:NINJA_CLIENT_SECRET = "yyyy"

.USAGE
  1) Save as: Upload-NinjaOneGlobalKB.ps1
  2) Set env vars (see above)
  3) Run:

      .\Upload-NinjaOneGlobalKB.ps1 -FolderId 24 -FilePath "C:\Temp\report.xlsx"

  Optional: specify a different BaseUrl if your tenant uses a custom domain (e.g. *.rmmservice.eu):
      .\Upload-NinjaOneGlobalKB.ps1 -BaseUrl "https://yourtenant.rmmservice.eu" -FolderId 24 -FilePath "C:\Temp\report.xlsx"

.PARAMETER BaseUrl
  NinjaOne regional base URL.
  - Europe: https://eu.ninjarmm.com
  - North America: https://app.ninjarmm.com
  - Oceania: https://oc.ninjarmm.com

  Note: some environments use a custom tenant domain (e.g. https://company.rmmservice.eu).
  If your OAuth token request fails, try passing your tenant domain via -BaseUrl.

.PARAMETER FolderId
  Target Knowledge Base folder ID (the numeric id you see in the UI URL:
  .../#/systemDashboard/knowledgeBase/<id>).

.PARAMETER FilePath
  Full path to the file to upload.

.NOTES
  - This script targets the *Global* Knowledge Base (no organizationId is sent).
  - The upload endpoint expects the file field name to be "files" (plural).
  - Some tenants expose endpoints under /v2, others under /api/v2; the script tries both.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)]
  [string]$BaseUrl = "https://eu.ninjarmm.com",

  [Parameter(Mandatory = $true)]
  [int]$FolderId,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$FilePath
)

$ErrorActionPreference = "Stop"

# -------------------------
# Validate inputs
# -------------------------
if (-not (Test-Path -LiteralPath $FilePath)) {
  throw "File not found: $FilePath"
}

$ClientId     = $env:NINJA_CLIENT_ID
$ClientSecret = $env:NINJA_CLIENT_SECRET

if ([string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($ClientSecret)) {
  throw "Missing credentials. Please set NINJA_CLIENT_ID and NINJA_CLIENT_SECRET environment variables."
}

# Normalize BaseUrl (remove trailing slash if present)
$BaseUrl = $BaseUrl.TrimEnd('/')

# -------------------------
# 1) OAuth token (Client Credentials)
# -------------------------
try {
  $tokenResp = Invoke-RestMethod -Method Post `
    -Uri "$BaseUrl/ws/oauth/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{
      grant_type    = "client_credentials"
      client_id     = $ClientId
      client_secret = $ClientSecret
      scope         = "monitoring management"
    }
}
catch {
  throw "Failed to obtain OAuth token from '$BaseUrl'. Error: $($_.Exception.Message)"
}

$accessToken = $tokenResp.access_token
if ([string]::IsNullOrWhiteSpace($accessToken)) {
  throw "OAuth token is empty. Response: $($tokenResp | ConvertTo-Json -Depth 10)"
}

# -------------------------
# 2) Multipart upload (HttpClient)
# -------------------------
Add-Type -AssemblyName System.Net.Http

$http = New-Object System.Net.Http.HttpClient
$http.DefaultRequestHeaders.Authorization =
  New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $accessToken)
$http.DefaultRequestHeaders.Accept.Add(
  (New-Object System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"))
)

$multipart = New-Object System.Net.Http.MultipartFormDataContent

# Global KB: do NOT include organizationId
$multipart.Add((New-Object System.Net.Http.StringContent($FolderId.ToString())), "folderId")

$stream = [System.IO.File]::OpenRead($FilePath)
try {
  $fileContent = New-Object System.Net.Http.StreamContent($stream)
  $fileContent.Headers.ContentType =
    [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/octet-stream")

  # IMPORTANT: field name must be "files" (plural)
  $multipart.Add($fileContent, "files", [System.IO.Path]::GetFileName($FilePath))

  # Try /v2 first, then /api/v2 as fallback
  $uris = @(
    "$BaseUrl/v2/knowledgebase/articles/upload",
    "$BaseUrl/api/v2/knowledgebase/articles/upload"
  )

  $resp = $null
  $body = $null
  $lastError = $null

  foreach ($u in $uris) {
    try {
      $resp = $http.PostAsync($u, $multipart).Result
      $body = $resp.Content.ReadAsStringAsync().Result

      if ($resp.IsSuccessStatusCode) {
        $lastError = $null
        break
      } else {
        $lastError = "Attempt: $u`nHTTP $([int]$resp.StatusCode) $($resp.ReasonPhrase)`n$body"
      }
    }
    catch {
      $lastError = "Attempt: $u`nException: $($_.Exception.Message)"
    }
  }

  if (-not $resp -or -not $resp.IsSuccessStatusCode) {
    throw "Upload failed.`n$lastError"
  }

  # Output: try to parse JSON, otherwise return raw text
  try { $body | ConvertFrom-Json } catch { $body }
}
finally {
  if ($stream) { $stream.Dispose() }
  if ($multipart) { $multipart.Dispose() }
  if ($http) { $http.Dispose() }
}
