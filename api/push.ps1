# Zip-deploy the backend to Azure App Service.
#   az login --tenant ad340c84-1886-4202-a483-2da2cb9168eb
#   ./push.ps1 -Subscription "<your sub>"
param(
  [Parameter(Mandatory = $true)][string]$Subscription,
  [string]$Rg  = "rg-opb-checkin",
  [string]$App = "opb-checkin-api"
)
$ErrorActionPreference = "Stop"
az account set --subscription $Subscription | Out-Null

$zip = Join-Path $env:TEMP "opb-checkin-api.zip"
Remove-Item $zip -ErrorAction SilentlyContinue
$files = "server.js", "graph.js", "rules.js", "package.json", "package-lock.json"
Compress-Archive -Path $files -DestinationPath $zip -Force

"Deploying $zip -> $App ..."
az webapp deploy -n $App -g $Rg --src-path $zip --type zip
"Health: https://$App.azurewebsites.net/health"
