targetScope = 'subscription'

@minLength(1)
@maxLength(64)
@description('Nom de l\'environnement azd (préfixe des ressources)')
param environmentName string

@minLength(1)
@description('Région Azure principale (ex. westeurope, francecentral)')
param location string

@description('Object ID de l\'utilisateur qui déploie (fourni automatiquement par azd)')
param principalId string = ''

@description('Tenant Microsoft 365 (GUID ou domaine) pour l\'auth Graph')
param outlookTenantId string

@description('Application (client) ID de l\'app Azure AD pour l\'auth Graph')
param outlookClientId string

@secure()
@description('Clé d\'API protégeant l\'endpoint MCP public')
param mcpApiKey string

@description('Image conteneur mutualisée (ghcr.io…). Vide = build local via ACR (mode perso).')
param containerImage string = ''

var tags = { 'azd-env-name': environmentName }
var resourceToken = toLower(uniqueString(subscription().id, environmentName, location))

resource rg 'Microsoft.Resources/resourceGroups@2024-03-01' = {
  name: 'rg-${environmentName}'
  location: location
  tags: tags
}

module resources 'resources.bicep' = {
  name: 'resources'
  scope: rg
  params: {
    location: location
    tags: tags
    resourceToken: resourceToken
    principalId: principalId
    outlookTenantId: outlookTenantId
    outlookClientId: outlookClientId
    mcpApiKey: mcpApiKey
    containerImage: containerImage
  }
}

output AZURE_CONTAINER_REGISTRY_ENDPOINT string = resources.outputs.AZURE_CONTAINER_REGISTRY_ENDPOINT
output AZURE_RESOURCE_GROUP string = rg.name
output AZURE_KEY_VAULT_NAME string = resources.outputs.keyVaultName
output OUTLOOK_MCP_KEYVAULT_URL string = resources.outputs.keyVaultUri
output SERVICE_API_NAME string = resources.outputs.containerAppName
output SERVICE_API_URI string = resources.outputs.containerAppUri
output MCP_ENDPOINT string = '${resources.outputs.containerAppUri}/mcp'
