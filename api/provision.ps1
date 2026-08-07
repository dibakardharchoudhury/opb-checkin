# Bare-minimum Azure provisioning for the OPB check-in backend.
# Hosting tenant: ad340c84-1886-4202-a483-2da2cb9168eb (Azure resources only).
# The workbook itself lives on a PERSONAL OneDrive and is reached via a delegated
# Graph refresh token (see authorize.js) — NOT via any Azure identity here.
#
# Resources created: 1 resource group + 1 Free (F1) Linux App Service plan + 1 web app.
# That's it. No Key Vault, no managed identity, no premium connectors.
#
# Usage:
#   az login --tenant ad340c84-1886-4202-a483-2da2cb9168eb
#   ./provision.ps1 -Subscription "<your sub name or id>" -PagesOrigin "https://<user>.github.io"
# Then set the three secrets once the app registration + authorize.js are done:
#   az webapp config appsettings set -n $App -g $Rg --settings `
#       OPB_CLIENT_ID=... OPB_CLIENT_SECRET=... OPB_REFRESH_TOKEN=...
# (GRAPH_WORKBOOK="path:/Oslo Durgotsav 2026_Test.xlsx" and TABLE_NAME are preset below.)

param(
  [Parameter(Mandatory = $true)][string]$Subscription,
  [Parameter(Mandatory = $true)][string]$PagesOrigin,
  [string]$Tenant = "ad340c84-1886-4202-a483-2da2cb9168eb",
  [string]$Rg     = "rg-opb-checkin",
  [string]$Loc    = "westeurope",
  [string]$Plan   = "opb-checkin-plan",
  [string]$App    = "opb-checkin-api"
)
$ErrorActionPreference = "Continue"

"=== account ==="
az account set --subscription $Subscription | Out-Null
az account show --query "{tenant:tenantId, sub:name}" -o json

"=== resource group ==="
az group create -n $Rg -l $Loc 2>&1 | Out-Null; "rg exit=$LASTEXITCODE"

"=== Free (F1) Linux plan ==="
az appservice plan create -n $Plan -g $Rg -l $Loc --is-linux --sku F1 2>&1 | Out-Null
"plan exit=$LASTEXITCODE"

"=== web app (Node 20) ==="
az webapp create -n $App -g $Rg -p $Plan --runtime "NODE:20-lts" 2>&1 | Out-Null
"webapp exit=$LASTEXITCODE"

"=== base app settings (secrets set separately, once) ==="
az webapp config appsettings set -n $App -g $Rg --settings `
  ALLOWED_ORIGINS="$PagesOrigin" `
  GRAPH_WORKBOOK="path:/Oslo Durgotsav 2026_Test.xlsx" `
  TABLE_NAME="Table1" `
  TZ_NAME="Europe/Oslo" `
  SCM_DO_BUILD_DURING_DEPLOYMENT="true" `
  WEBSITE_NODE_DEFAULT_VERSION="~20" 2>&1 | Out-Null
"settings exit=$LASTEXITCODE"

"=== startup command ==="
az webapp config set -n $App -g $Rg --startup-file "node server.js" 2>&1 | Out-Null
"startup exit=$LASTEXITCODE"

"=== result ==="
az webapp show -n $App -g $Rg --query "{host:defaultHostName, state:state}" -o json
"`nNext: set OPB_CLIENT_ID / OPB_CLIENT_SECRET / OPB_REFRESH_TOKEN / GRAPH_WORKBOOK, then deploy with push.ps1."
