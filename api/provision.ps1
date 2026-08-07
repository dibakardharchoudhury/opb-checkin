# Bare-minimum Azure provisioning for the OPB check-in backend — REUSES the existing
# NorkappTrip infrastructure (no new plan, no new resource group, no extra cost).
#
# Hosting tenant: ad340c84-1886-4202-a483-2da2cb9168eb (the subscription below lives here).
# We add ONE web app onto the existing B1 Linux plan "nordkapp-ai-plan" in "rg-agentmcp".
# A B1 App Service plan hosts multiple apps for free, so this adds no monthly cost.
#
# The workbook itself lives on a PERSONAL OneDrive and is reached via a delegated
# Graph refresh token (see authorize.js) — NOT via any Azure identity here. So unlike
# NorkappTrip there is no managed identity / role assignment to make.
#
# Usage:
#   az login --tenant ad340c84-1886-4202-a483-2da2cb9168eb
#   ./provision.ps1                     # uses the reused-infra defaults below
# Then set the three secrets once the app registration + authorize.js are done:
#   az webapp config appsettings set -n $App -g $Rg --settings `
#       OPB_CLIENT_ID=... OPB_CLIENT_SECRET=... OPB_REFRESH_TOKEN=...
# (GRAPH_WORKBOOK="path:/Oslo Durgotsav 2026_Test.xlsx" and TABLE_NAME are preset below.)

param(
  [string]$Subscription = "ME-MngEnvMCAP677316-didharch-1",  # same sub as NorkappTrip
  [string]$PagesOrigin  = "https://dibakardharchoudhury.github.io",
  [string]$Tenant       = "ad340c84-1886-4202-a483-2da2cb9168eb",
  [string]$Rg           = "rg-agentmcp",       # reuse NorkappTrip resource group
  [string]$Plan         = "nordkapp-ai-plan",  # reuse NorkappTrip B1 Linux plan
  [string]$Loc          = "swedencentral",     # only used if the plan must be created
  [string]$App          = "opb-checkin-api"
)
$ErrorActionPreference = "Continue"

"=== account ==="
az account set --subscription $Subscription | Out-Null
az account show --query "{tenant:tenantId, sub:name}" -o json

"=== ensure resource group (reused) ==="
az group create -n $Rg -l $Loc 2>&1 | Out-Null; "rg exit=$LASTEXITCODE"

"=== ensure plan (reuse nordkapp-ai-plan; create B1 only if missing) ==="
$planId = az appservice plan show -n $Plan -g $Rg --query id -o tsv 2>$null
if (-not $planId) {
  az appservice plan create -n $Plan -g $Rg -l $Loc --is-linux --sku B1 2>&1 | Out-Null
  "plan created exit=$LASTEXITCODE"
} else { "plan reused: $Plan" }

"=== web app (Node 20 on the shared plan) ==="
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
