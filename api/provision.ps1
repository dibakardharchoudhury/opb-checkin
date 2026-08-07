# Bare-minimum Azure provisioning for the OPB check-in backend.
# All resources are created in tenant ad340c84-1886-4202-a483-2da2cb9168eb
# (the previous NorkappTrip tenant/subscription is being decommissioned, so we do
# NOT reuse it). Footprint: 1 resource group + 1 Free (F1) Linux plan + 1 web app.
# No Key Vault, no managed identity, no premium connectors.
#
# The workbook itself lives on a PERSONAL OneDrive and is reached via a delegated
# Graph refresh token (see authorize.js) — NOT via any Azure identity here.
#
# Usage:
#   az login --tenant ad340c84-1886-4202-a483-2da2cb9168eb
#   ./provision.ps1                     # uses the tenant-ad340c84 defaults below
# Then set the three secrets once the app registration + authorize.js are done:
#   az webapp config appsettings set -n $App -g $Rg --settings `
#       OPB_CLIENT_ID=... OPB_CLIENT_SECRET=... OPB_REFRESH_TOKEN=...
# (GRAPH_WORKBOOK="path:/Oslo Durgotsav 2026_Test.xlsx" and TABLE_NAME are preset below.)

param(
  # Subscription in tenant ad340c84 (default). Other option: ME-MngEnvMCAP218279-didharch-1.
  [string]$Subscription = "ME-MngEnvMCAP218279-didharch-2",
  [string]$PagesOrigin  = "https://dibakardharchoudhury.github.io",
  [string]$Tenant       = "ad340c84-1886-4202-a483-2da2cb9168eb",
  [string]$Rg           = "rg-opb-checkin",
  [string]$Plan         = "opb-checkin-plan",
  [string]$Loc          = "swedencentral",
  [string]$App          = "opb-checkin-api"
)
$ErrorActionPreference = "Continue"

"=== account ==="
az account set --subscription $Subscription | Out-Null
az account show --query "{tenant:tenantId, sub:name}" -o json

"=== resource group ==="
az group create -n $Rg -l $Loc 2>&1 | Out-Null; "rg exit=$LASTEXITCODE"

"=== Free (F1) Linux plan (create if missing) ==="
$planId = az appservice plan show -n $Plan -g $Rg --query id -o tsv 2>$null
if (-not $planId) {
  az appservice plan create -n $Plan -g $Rg -l $Loc --is-linux --sku F1 2>&1 | Out-Null
  "plan created exit=$LASTEXITCODE"
} else { "plan reused: $Plan" }

"=== web app (Node 20) ==="
az webapp create -n $App -g $Rg -p $Plan --runtime "NODE:20-lts" 2>&1 | Out-Null
"webapp exit=$LASTEXITCODE"

"=== base app settings (secrets set separately, once) ==="
# A random session-signing secret for the login JWTs (generated once per provision).
$sessionSecret = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 48 | ForEach-Object { [char]$_ })
az webapp config appsettings set -n $App -g $Rg --settings `
  ALLOWED_ORIGINS="$PagesOrigin" `
  GRAPH_WORKBOOK="path:/Oslo Durgotsav 2026_Test.xlsx" `
  TABLE_NAME="Table1" `
  TZ_NAME="Europe/Oslo" `
  SESSION_SECRET="$sessionSecret" `
  ADMIN_EMAILS="" `
  USER_EMAILS="" `
  GOOGLE_CLIENT_ID="" `
  MS_CLIENT_ID="" `
  SCM_DO_BUILD_DURING_DEPLOYMENT="true" `
  WEBSITE_NODE_DEFAULT_VERSION="~20" 2>&1 | Out-Null
"settings exit=$LASTEXITCODE"

"=== startup command ==="
az webapp config set -n $App -g $Rg --startup-file "node server.js" 2>&1 | Out-Null
"startup exit=$LASTEXITCODE"

"=== result ==="
az webapp show -n $App -g $Rg --query "{host:defaultHostName, state:state}" -o json
"`nNext: set OPB_CLIENT_ID / OPB_CLIENT_SECRET / OPB_REFRESH_TOKEN / GRAPH_WORKBOOK, then deploy with push.ps1."
