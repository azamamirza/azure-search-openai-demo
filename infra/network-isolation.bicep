metadata description = 'Sets up private networking for all resources, using VNet, private endpoints, and DNS zones.'

param vnetResourceGroupName string

@description('The name of the VNet to create')
param vnetName string

@description('The location to create the VNet and private endpoints')
param location string = resourceGroup().location

@description('The tags to apply to all resources')
param tags object = {}

@description('The name of an existing App Service Plan to connect to the VNet')
param appServicePlanName string

param backendSubnetName string
param appSubnetName string

param usePrivateEndpoint bool = false

@allowed(['appservice', 'containerapps'])
param deploymentTarget string

resource appServicePlan 'Microsoft.Web/serverfarms@2022-03-01' existing = if (deploymentTarget == 'appservice') {
  name: appServicePlanName
}

resource vnet 'Microsoft.Network/virtualNetworks@2024-05-01' existing = if (usePrivateEndpoint) {
  scope: resourceGroup(vnetResourceGroupName)
  name: vnetName
}
resource backendSubnet 'Microsoft.Network/virtualNetworks/subnets@2024-05-01' existing = if (usePrivateEndpoint) {
  parent: vnet
  name: backendSubnetName
}
resource appSubnet 'Microsoft.Network/virtualNetworks/subnets@2024-05-01' existing = if (usePrivateEndpoint) {
  parent: vnet 
  name: appSubnetName
}

output appSubnetId string = usePrivateEndpoint ? appSubnet.id : ''
output backendSubnetId string = usePrivateEndpoint ? backendSubnet.id : ''
output vnetName string = usePrivateEndpoint ? vnet.name : ''
