@description('Région des ressources')
param location string

@description('Étiquettes communes')
param tags object

@description('Jeton d\'unicité des noms')
param resourceToken string

@description('Object ID de l\'utilisateur déployeur (accès secret pour amorcer le jeton)')
param principalId string = ''

param outlookTenantId string
param outlookClientId string

@secure()
param mcpApiKey string

@description('Image conteneur mutualisée (ex. ghcr.io/org/img:tag). Vide = build local via ACR (mode perso).')
param containerImage string = ''

var abbrs = {
  containerApp: 'ca-api-${resourceToken}'
  env: 'cae-${resourceToken}'
  registry: 'acr${resourceToken}'
  identity: 'id-${resourceToken}'
  keyvault: 'kv-${resourceToken}'
  logs: 'log-${resourceToken}'
}

// Rôle intégré Azure (AcrPull pour tirer l'image depuis l'ACR, mode perso)
var roleAcrPull = subscriptionResourceId('Microsoft.Authorization/roleDefinitions', '7f951dda-4ed3-4680-a7ca-43fe172d538d')
var kvSecretPerms = { secrets: [ 'get', 'list', 'set' ] }

var cacheSecretName = 'outlook-token-cache'
var placeholderImage = 'mcr.microsoft.com/azuredocs/containerapps-helloworld:latest'

// Mode « image mutualisée » : si containerImage est fourni (ex. ghcr.io public),
// on tire cette image et on ne crée PAS d'ACR → ~0 € côté client.
// Si vide : build local poussé dans un ACR dédié (mode perso, via `azd up`).
var useAcr = empty(containerImage)
var appImage = useAcr ? placeholderImage : containerImage

// ----------------------------------------------------------------------------
// Identité managée portée par l'application
// ----------------------------------------------------------------------------
resource uami 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: abbrs.identity
  location: location
  tags: tags
}

// ----------------------------------------------------------------------------
// Observabilité + environnement Container Apps
// ----------------------------------------------------------------------------
resource logs 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: abbrs.logs
  location: location
  tags: tags
  properties: {
    sku: { name: 'PerGB2018' }
    retentionInDays: 30
  }
}

resource env 'Microsoft.App/managedEnvironments@2024-03-01' = {
  name: abbrs.env
  location: location
  tags: tags
  properties: {
    appLogsConfiguration: {
      destination: 'log-analytics'
      logAnalyticsConfiguration: {
        customerId: logs.properties.customerId
        sharedKey: logs.listKeys().primarySharedKey
      }
    }
  }
}

// ----------------------------------------------------------------------------
// Registre de conteneurs — UNIQUEMENT en mode build local (perso)
// ----------------------------------------------------------------------------
resource registry 'Microsoft.ContainerRegistry/registries@2023-07-01' = if (useAcr) {
  name: abbrs.registry
  location: location
  tags: tags
  sku: { name: 'Basic' }
  properties: {
    adminUserEnabled: false
  }
}

resource acrPull 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (useAcr) {
  name: guid(resourceToken, 'acrpull')
  scope: registry
  properties: {
    roleDefinitionId: roleAcrPull
    principalId: uami.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

// ----------------------------------------------------------------------------
// Key Vault : stockage du jeton délégué (cache MSAL)
// ----------------------------------------------------------------------------
resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: abbrs.keyvault
  location: location
  tags: tags
  properties: {
    sku: { family: 'A', name: 'standard' }
    tenantId: subscription().tenantId
    // Stratégies d'accès (au lieu du RBAC) : déterministe, sans dépendre de la
    // propagation des définitions de rôles sur un abonnement neuf.
    enableRbacAuthorization: false
    accessPolicies: concat(
      [
        {
          // L'app (identité managée) lit ET écrit le secret (refresh du jeton).
          tenantId: subscription().tenantId
          objectId: uami.properties.principalId
          permissions: kvSecretPerms
        }
      ],
      empty(principalId) ? [] : [
        {
          // Le déployeur humain : pour amorcer le jeton en local (commande `login`).
          tenantId: subscription().tenantId
          objectId: principalId
          permissions: kvSecretPerms
        }
      ]
    )
    enableSoftDelete: true
    softDeleteRetentionInDays: 7
    publicNetworkAccess: 'Enabled'
  }
}

// ----------------------------------------------------------------------------
// Container App
// ----------------------------------------------------------------------------
resource containerApp 'Microsoft.App/containerApps@2024-03-01' = {
  name: abbrs.containerApp
  location: location
  tags: union(tags, { 'azd-service-name': 'api' })
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${uami.id}': {}
    }
  }
  properties: {
    managedEnvironmentId: env.id
    configuration: {
      activeRevisionsMode: 'Single'
      ingress: {
        external: true
        targetPort: 8000
        transport: 'auto'
        allowInsecure: false
      }
      // En mode image mutualisée (publique), aucun registre privé n'est requis.
      registries: useAcr ? [
        {
          server: registry.?properties.loginServer ?? ''
          identity: uami.id
        }
      ] : []
      secrets: [
        {
          name: 'mcp-api-key'
          value: mcpApiKey
        }
      ]
    }
    template: {
      containers: [
        {
          name: 'api'
          image: appImage
          resources: {
            cpu: json('0.5')
            memory: '1Gi'
          }
          env: [
            { name: 'OUTLOOK_MCP_CLIENT_ID', value: outlookClientId }
            { name: 'OUTLOOK_MCP_TENANT_ID', value: outlookTenantId }
            { name: 'OUTLOOK_MCP_CACHE_BACKEND', value: 'keyvault' }
            { name: 'OUTLOOK_MCP_KEYVAULT_URL', value: keyVault.properties.vaultUri }
            { name: 'OUTLOOK_MCP_CACHE_SECRET', value: cacheSecretName }
            { name: 'OUTLOOK_MCP_TRANSPORT', value: 'http' }
            { name: 'OUTLOOK_MCP_IMMUTABLE_IDS', value: 'true' }
            { name: 'AZURE_CLIENT_ID', value: uami.properties.clientId }
            { name: 'PORT', value: '8000' }
            { name: 'OUTLOOK_MCP_API_KEY', secretRef: 'mcp-api-key' }
          ]
        }
      ]
      scale: {
        minReplicas: 0
        maxReplicas: 1
      }
    }
  }
}

output AZURE_CONTAINER_REGISTRY_ENDPOINT string = registry.?properties.loginServer ?? ''
output keyVaultName string = keyVault.name
output keyVaultUri string = keyVault.properties.vaultUri
output containerAppName string = containerApp.name
output containerAppUri string = 'https://${containerApp.properties.configuration.ingress.fqdn}'
