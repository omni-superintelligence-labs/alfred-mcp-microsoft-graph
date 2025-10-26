# Alfred Microsoft Graph MCP - Setup Guide

This package provides both an MCP server (stdio) and an HTTP API server for Microsoft Graph integration with Excel.

## Prerequisites

1. **Azure App Registration**
   - Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
   - Create a new registration
   - Set redirect URI: `alfred://auth/microsoft/callback`
   - Add API permissions: `Files.ReadWrite.All`, `offline_access`
   - Copy Client ID and Tenant ID

2. **Office Add-in Certificate** (for HTTPS dev server)
   ```bash
   npx office-addin-dev-certs install
   ```

## Configuration

1. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

2. Edit `.env` and fill in your values:
   ```env
   AZURE_CLIENT_ID=your-client-id-here
   AZURE_TENANT_ID=common
   # ... other values
   ```

3. Generate a GUID for your add-in:
   ```bash
   uuidgen  # or use an online generator
   ```
   
   Update `ADDIN_ID` in `.env` and replace `YOUR-ADDIN-GUID-HERE` in `config/manifest.dev.xml`

## Running

### HTTP API Server (for Office Add-in)

```bash
npm run dev:http
```

This starts the HTTP server on port 3100 (configurable via `HTTP_PORT` env var).

Endpoints:
- `POST /microsoft-graph/workbook/apply` - Apply workbook changes
- `GET /microsoft-graph/workbook/metadata?itemId=...` - Get workbook metadata

### MCP Server (stdio)

```bash
npm run dev
```

This starts the MCP server for integration with Alfred's MCP manager.

## Sideloading the Office Add-in

### Development (Excel for Web)

1. Start the frontend dev server with HTTPS (must run on `https://localhost:3900`):
   ```bash
   cd ../alfred-frontend
   npm run dev:addin
   ```

2. Start the HTTP API server:
   ```bash
   npm run dev:http
   ```

3. Open Excel Online in your browser
4. Go to Insert → Add-ins → Upload My Add-in
5. Upload `config/manifest.dev.xml`

### Development (Excel Desktop - macOS/Windows)

Use the [Office Add-ins CLI](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins):

```bash
npx office-addin-debugging start config/manifest.dev.xml desktop
```

## Architecture

```
┌─────────────────┐
│ Excel for Web   │
│  (Office.js)    │
└────────┬────────┘
         │
         │ Task Pane loads from
         │ https://localhost:3900/office-addin/
         │
         v
┌─────────────────┐
│ Office Add-in   │
│  (React + TS)   │
└────────┬────────┘
         │
         │ HTTP API calls
         │ (with MS Graph token)
         │
         v
┌─────────────────┐
│  HTTP Server    │
│  (port 3100)    │
└────────┬────────┘
         │
         │ Microsoft Graph
         │ Workbook APIs
         │
         v
┌─────────────────┐
│ Microsoft Graph │
│   (cloud)       │
└─────────────────┘
```

## Next Steps

- Implement actual Graph Workbook API calls in `src/http-server.ts`
- Add Graph client library (`@microsoft/microsoft-graph-client`)
- Implement token validation and refresh
- Add error handling and retry logic (429 throttling)
- Production manifest and deployment config

