# Zip-deploy the backend to Azure App Service (tenant ad340c84).
#   az login --tenant ad340c84-1886-4202-a483-2da2cb9168eb
#   ./push.ps1
param(
  [string]$Subscription = "ME-MngEnvMCAP218279-didharch-2",   # tenant ad340c84
  [string]$Rg  = "rg-opb-checkin",
  [string]$App = "opb-checkin-api"
)
$ErrorActionPreference = "Stop"
az account set --subscription $Subscription | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Could not select Azure subscription $Subscription." }
$subscriptionId = az account show --query id -o tsv
if ($LASTEXITCODE -ne 0 -or -not $subscriptionId) { throw "Could not resolve Azure subscription." }

if (Test-Path app.zip) { Remove-Item app.zip -Force }
Compress-Archive -Path server.js, graph.js, rules.js, auth.js, userstore.js, configstore.js, package.json, package-lock.json -DestinationPath app.zip -Force

"=== deploy (SCM publishing is restored to disabled afterward) ==="
$scmPolicyId = "/subscriptions/$subscriptionId/resourceGroups/$Rg/providers/Microsoft.Web/sites/$App/basicPublishingCredentialsPolicies/scm"
try {
  az resource update --ids $scmPolicyId --api-version 2022-03-01 --set properties.allow=true --output none
  if ($LASTEXITCODE -ne 0) { throw "Could not enable SCM publishing for deployment." }
  az webapp deploy --name $App --resource-group $Rg --src-path app.zip --type zip --clean true --restart true
  if ($LASTEXITCODE -ne 0) { throw "Azure ZIP deployment failed." }
} finally {
  az resource update --ids $scmPolicyId --api-version 2022-03-01 --set properties.allow=false --output none
  if ($LASTEXITCODE -ne 0) { throw "CRITICAL: could not disable SCM publishing after deployment." }
}
"Health: https://$App.azurewebsites.net/health"
