/**
 * Configuration for Microsoft Graph MCP server
 */

export const config = {
  azure: {
    clientId: process.env.AZURE_CLIENT_ID || '',
    tenantId: process.env.AZURE_TENANT_ID || 'common',
    redirectUri: process.env.AZURE_REDIRECT_URI || 'alfred://auth/microsoft/callback',
    scopes: (process.env.SCOPES || 'Files.ReadWrite.All offline_access').split(' '),
  },
  
  onedrive: {
    uploadRoot: process.env.ONEDRIVE_UPLOAD_ROOT || 'Apps/Alfred/Uploads',
  },
  
  addin: {
    id: process.env.ADDIN_ID || '',
    sourceUrl: process.env.ADDIN_SOURCE_URL || 'https://localhost:3900/office-addin/',
    allowedOrigins: (process.env.ADDIN_ALLOWED_ORIGINS || '').split(',').filter(Boolean),
  },
  
  http: {
    port: parseInt(process.env.HTTP_PORT || '3100', 10),
    corsOrigins: (process.env.CORS_ORIGINS || 'https://localhost:3900,alfred://*').split(','),
  },
};

