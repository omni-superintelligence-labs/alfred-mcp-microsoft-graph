# Alfred Microsoft Graph MCP Server

MCP server providing Microsoft Graph API integration for Alfred chat assistant.

## Tools

- `excel_read_range` - Read cells from Excel workbook
- `excel_write_range` - Write data to Excel workbook
- `outlook_list_messages` - List recent emails
- `outlook_send_message` - Send email
- `teams_send_chat_message` - Send Teams chat message

## Usage

Launched by Alfred Electron app via stdio transport.

**Environment variables**:
- `BACKEND_BASE_URL` - Alfred backend URL (default: http://localhost:3000)
- `SERVICE_JWT` - JWT token from Alfred identity service

## Development

```bash
npm install
npm run dev
```

## Production

```bash
npm run build
npm start
```

