param NSGName string
param RouteTableName string
param HubGatewayIP string
param location string = resourceGroup().location



resource rt 'Microsoft.Network/routeTables@2024-05-01' = {
  name: RouteTableName
  location: location
  properties: {
    disableBgpRoutePropagation: true
  }

  resource monitor_to_internet 'routes' = {
    name: 'Monitor-To-Internet'
    properties: {
      addressPrefix: 'AzureMonitor'
      nextHopType: 'Internet'
      //hasBgpOverride: true
    }
  }
  
  resource other_to_internet 'routes' = {
    name: 'Other-To-Internet'
    properties: {
      addressPrefix: '0.0.0.0/0'
      nextHopType: 'Internet'
    }
  }
  
  resource internal_to_hub 'routes' = {
    name: 'Internal-To-Hub'
    properties: {
      addressPrefix: '10.0.0.0/8'
      nextHopType: 'VirtualAppliance'
      nextHopIpAddress: HubGatewayIP
    }
  }
}

resource nsg 'Microsoft.Network/networkSecurityGroups@2024-05-01' = {
  name: NSGName
  location: location
  properties: {}

  resource allow_controlplane 'securityRules' = {
    name: 'Allow-AppGW-Inbound'
    properties: {
      protocol: 'Tcp'
      access: 'Allow'
      direction: 'Inbound'
      sourceAddressPrefix: 'Internet'
      sourcePortRange: '*'
      destinationAddressPrefix: 'VirtualNetwork'
      destinationPortRange: '443'
      priority: 101
    }
  }
  
  resource allow_akamai 'securityRules' = {
    name: 'Allow-AzureLoadBalancer-Inbound'
    properties: {
      protocol: 'Tcp'
      access: 'Allow'
      direction: 'Inbound'
      sourceAddressPrefix: 'AzureLoadBalancer'
      sourcePortRange: '*'
      destinationAddressPrefix: 'VirtualNetwork'
      destinationPortRange: '6390'
      priority: 102
    }
  }
  
  resource allow_SQL 'securityRules' = {
    name: 'Allow-SQL-Outbound'
    properties: {
      protocol: 'Tcp'
      direction: 'Outbound'
      access: 'Allow'
      sourceAddressPrefix: 'VirtualNetwork'
      sourcePortRange: '*'
      destinationAddressPrefix: 'SQL'
      destinationPortRange: '1433'
      priority: 104
    }
  }
  
  resource allow_key_vault 'securityRules' = {
    name: 'Allow-Key-Vault-Outbound'
    properties: {
      protocol: 'Tcp'
      direction: 'Outbound'
      access: 'Allow'
      sourceAddressPrefix: 'VirtualNetwork'
      sourcePortRange: '*'
      destinationAddressPrefix: 'AzureKeyVault'
      destinationPortRange: '443'
      priority: 105
    }
  }
  
  resource allow_monitor_1886 'securityRules' = {
    name: 'Allow-Monitor-1886-Outbound'
    properties: {
      protocol: 'Tcp'
      access: 'Allow'
      direction: 'Outbound'
      sourceAddressPrefix: 'VirtualNetwork'
      sourcePortRange: '*'
      destinationAddressPrefix: 'AzureMonitor'
      destinationPortRange: '1886'
      priority: 106
    }
  }
  resource allow_gwman_65200_65535 'securityRules' = {
    name: 'Allow-GatewayManager-65200-65535-Inbound'
    properties: {
      protocol: 'Tcp'
      access: 'Allow'
      direction: 'Inbound'
      sourceAddressPrefix: 'GatewayManager'
      sourcePortRange: '*'
      destinationAddressPrefix: '*'
      destinationPortRange: '65200-65535'
      priority: 107
    }
  }
}

output rt_id string = rt.id
output nsg_id string = nsg.id
