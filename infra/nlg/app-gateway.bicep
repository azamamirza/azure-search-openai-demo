param name string
param location string = resourceGroup().location
param vnetName string
param ipStart string
param ipCidr string
var subnetTier = 'AGW'
param hubGatewayIP string
param zones array = []


module subnet 'app-gateway-subnet.bicep' = {
  name: 'AppGwSubnet'
  params: {
    hubGatewayIP: hubGatewayIP
    ipCidr: ipCidr
    ipStart: ipStart
    subnetTier: subnetTier
    vnetName: vnetName
  }
}

resource publicIP 'Microsoft.Network/publicIPAddresses@2024-05-01' = {
  name: '${toUpper(name)}PIP'
  location: location
  sku: {
    name: 'Standard'
  }
  zones: zones
  properties: {
    publicIPAddressVersion: 'IPv4'
    publicIPAllocationMethod: 'Static'
  }
}

module appgw 'br/public:avm/res/network/application-gateway:0.5.1' = {
  name: 'AppGw'
  params: {
    name: name 
    sku: 'Standard_v2'
    lock: { name: 'None' }
    zones: zones
  }
}
