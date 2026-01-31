# Upload-NinjaOneGlobalKB.ps1

Upload a file to the **NinjaOne Global Knowledge Base** into a specific folder (`folderId`) using **OAuth2 Client Credentials** and **multipart/form-data**.

This script is compatible with **Windows PowerShell 5.1** (where `Invoke-RestMethod -Form` is not available), so it uses **.NET `HttpClient`** and `MultipartFormDataContent`.

---

## Features

- OAuth2 **Client Credentials** authentication
- Uploads to the **Global** Knowledge Base (**no `organizationId`** is sent)
- Uses multipart upload with the required file field name: **`files`** (plural)
- Automatically tries both possible API paths:
  - `/v2/knowledgebase/articles/upload`
  - `/api/v2/knowledgebase/articles/upload` (fallback)

---

## Requirements

- **Windows PowerShell 5.1**
- A **NinjaOne API application** (Client ID / Client Secret) with permission to manage the Knowledge Base
- Environment variables configured:
  - `NINJA_CLIENT_ID`
  - `NINJA_CLIENT_SECRET`

---

## Setup

1. Save the script as:

   `Upload-NinjaOneGlobalKB.ps1`

2. Set the environment variables (example for current session only):

```powershell
$env:NINJA_CLIENT_ID     = "xxxx"
$env:NINJA_CLIENT_SECRET = "yyyy"