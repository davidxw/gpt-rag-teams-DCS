// Cross-RG role assignment helper. Scoped at the storage account so the
// bot's UMI can read blobs via DefaultAzureCredential.
@description('Existing storage account name in the current resource group scope.')
param storageAccountName string

@description('Object (principal) id of the identity that needs Storage Blob Data Reader.')
param principalId string

// Storage Blob Data Reader — built-in role.
// https://learn.microsoft.com/azure/role-based-access-control/built-in-roles/storage#storage-blob-data-reader
var storageBlobDataReaderRoleId = '2a2b9908-6ea1-4ae2-8e65-a410df84e7d1'

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' existing = {
  name: storageAccountName
}

resource roleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(storageAccount.id, principalId, storageBlobDataReaderRoleId)
  scope: storageAccount
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      storageBlobDataReaderRoleId
    )
    principalId: principalId
    principalType: 'ServicePrincipal'
  }
}
