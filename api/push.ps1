# Zip-deploy the backend to Azure App Service (tenant ad340c84).
#   az login --tenant ad340c84-1886-4202-a483-2da2cb9168eb
#   ./push.ps1
param(
  [string]$Subscription = "ME-MngEnvMCAP218279-didharch-2",   # tenant ad340c84
  [string]$Rg  = "rg-opb-checkin",
  [string]$App = "opb-checkin-api"
)
$ErrorActionPreference = "Continue"
az account set --subscription $Subscription | Out-Null

if (Test-Path app.zip) { Remove-Item app.zip -Force }
Compress-Archive -Path server.js, graph.js, rules.js, auth.js, userstore.js, configstore.js, package.json, package-lock.json -DestinationPath app.zip -Force
"zip exit=$LASTEXITCODE"

"=== deploy ==="
az webapp deploy --name $App --resource-group $Rg --src-path app.zip --type zip 2>&1 | Select-Object -Last 6
"deploy exit=$LASTEXITCODE"
"Health: https://$App.azurewebsites.net/health"
