using './main.bicep'

param appServiceName = 'opb-checkin-api'
param appServicePlanName = 'opb-checkin-plan'
param frontDoorProfileName = 'afd-opb-checkin'
param tags = {
  application: 'opb-checkin'
  environment: 'production'
  managedBy: 'bicep'
}