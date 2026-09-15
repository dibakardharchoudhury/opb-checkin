@description('Name of the Azure Front Door profile.')
param profileName string

@description('Globally unique name of the Azure Front Door endpoint.')
param endpointName string

@description('DNS hostname of the existing App Service origin.')
param originHostName string

@description('Tags applied to Azure Front Door resources.')
param tags object

var originGroupName = 'opb-checkin-origin-group'
var originName = 'opb-checkin-api-origin'
var routeName = 'api-route'

resource profile 'Microsoft.Cdn/profiles@2025-06-01' = {
  name: profileName
  location: 'global'
  sku: {
    name: 'Standard_AzureFrontDoor'
  }
  properties: {
    originResponseTimeoutSeconds: 60
  }
  tags: tags
}

resource originGroup 'Microsoft.Cdn/profiles/originGroups@2025-06-01' = {
  parent: profile
  name: originGroupName
  properties: {
    healthProbeSettings: {
      probeIntervalInSeconds: 120
      probePath: '/health'
      probeProtocol: 'Https'
      probeRequestType: 'GET'
    }
    loadBalancingSettings: {
      additionalLatencyInMilliseconds: 50
      sampleSize: 4
      successfulSamplesRequired: 3
    }
    sessionAffinityState: 'Disabled'
  }
}

resource origin 'Microsoft.Cdn/profiles/originGroups/origins@2025-06-01' = {
  parent: originGroup
  name: originName
  properties: {
    enabledState: 'Enabled'
    enforceCertificateNameCheck: true
    hostName: originHostName
    httpPort: 80
    httpsPort: 443
    originHostHeader: originHostName
    priority: 1
    weight: 1000
  }
}

resource endpoint 'Microsoft.Cdn/profiles/afdEndpoints@2025-06-01' = {
  parent: profile
  name: endpointName
  location: 'global'
  properties: {
    enabledState: 'Enabled'
  }
  tags: tags
}

resource route 'Microsoft.Cdn/profiles/afdEndpoints/routes@2025-06-01' = {
  parent: endpoint
  name: routeName
  properties: {
    enabledState: 'Enabled'
    forwardingProtocol: 'HttpsOnly'
    httpsRedirect: 'Enabled'
    linkToDefaultDomain: 'Enabled'
    originGroup: {
      id: originGroup.id
    }
    patternsToMatch: [
      '/*'
    ]
    supportedProtocols: [
      'Https'
    ]
  }
  dependsOn: [
    origin
  ]
}

output endpointHostName string = endpoint.properties.hostName
output profileId string = profile.id