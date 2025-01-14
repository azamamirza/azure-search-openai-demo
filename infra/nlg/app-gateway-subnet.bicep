param vnetName string
param ipStart string
param ipCidr string
param subnetTier string = 'AGW'
param hubGatewayIP string

resource vnet 'Microsoft.Network/virtualNetworks@2024-05-01' existing = {
  name: vnetName
}

module rt_nsg 'app-gateway-rt-nsg.bicep' = {
  name: 'RoutesAndNSGRules'
  params: {
    HubGatewayIP: hubGatewayIP
    NSGName: '${vnet.name}${subnetTier}NSG'
    RouteTableName: '${vnet.name}${subnetTier}RT'
  }
}

resource subnet 'Microsoft.Network/virtualNetworks/subnets@2024-05-01' = {
  parent: vnet
  name: 'SN_${ipStart}_${subnetTier}'
  properties: {
    addressPrefix: '${ipStart}/${ipCidr}'
    networkSecurityGroup: {
      id: rt_nsg.outputs.nsg_id
    }
    routeTable: {
      id: rt_nsg.outputs.rt_id
    }
  }
}

output subnet_id string = subnet.id
