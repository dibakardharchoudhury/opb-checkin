targetScope = 'resourceGroup'

@description('Name of the existing App Service API.')
param appServiceName string = 'opb-checkin-api'

@description('Name of the existing App Service plan.')
param appServicePlanName string = 'opb-checkin-plan'

@description('Name of the Azure Front Door profile.')
param frontDoorProfileName string = 'afd-opb-checkin'

@description('Globally unique Azure Front Door endpoint name.')
param frontDoorEndpointName string = 'opb-checkin-${uniqueString(subscription().id, resourceGroup().id)}'

@description('Tags applied to the new Azure Front Door resources.')
param tags object = {
  application: 'opb-checkin'
  environment: 'production'
  managedBy: 'bicep'
}

resource appService 'Microsoft.Web/sites@2024-11-01' existing = {
  name: appServiceName
}

resource appServicePlan 'Microsoft.Web/serverfarms@2024-11-01' existing = {
  name: appServicePlanName
}

module frontDoor './modules/front-door.bicep' = {
  name: 'deploy-front-door'
  params: {
    endpointName: frontDoorEndpointName
    originHostName: appService.properties.defaultHostName
    profileName: frontDoorProfileName
    tags: tags
  }
}

output appServiceId string = appService.id
output appServicePlanId string = appServicePlan.id
output frontDoorEndpointHostName string = frontDoor.outputs.endpointHostName
output frontDoorEndpointUrl string = 'https://${frontDoor.outputs.endpointHostName}'
output frontDoorProfileId string = frontDoor.outputs.profileId