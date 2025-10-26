/**
 * HTTP API server for Office Add-in communication
 * Provides REST endpoints for workbook operations via Graph API
 */

import http from 'http';
import { config } from './config.js';

interface WorkbookOperation {
  type: 'insert' | 'update' | 'delete' | 'format' | 'chart' | 'table';
  target: string;
  data?: any;
  options?: Record<string, any>;
}

interface ApplyChangesRequest {
  workbookId: string;
  driveId?: string;
  operations: WorkbookOperation[];
  clientContext?: Record<string, any>;
}

/**
 * Parse JSON body from request
 */
function parseBody(req: http.IncomingMessage): Promise<any> {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        resolve(JSON.parse(body));
      } catch (err) {
        reject(new Error('Invalid JSON'));
      }
    });
    req.on('error', reject);
  });
}

/**
 * Send JSON response
 */
function sendJson(res: http.ServerResponse, status: number, data: any) {
  res.writeHead(status, { 
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*', // In production, use config.http.corsOrigins
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  });
  res.end(JSON.stringify(data));
}

/**
 * Apply workbook changes via Graph Workbook API
 */
async function applyWorkbookChanges(
  token: string,
  request: ApplyChangesRequest
): Promise<any> {
  // TODO: Implement actual Graph API calls
  // For now, return mock success
  console.log('[HTTP] Apply workbook changes:', {
    workbookId: request.workbookId,
    operationCount: request.operations.length,
  });

  return {
    success: true,
    appliedOperations: request.operations,
    message: 'Operations queued (mock implementation)',
  };
}

/**
 * Get workbook metadata
 */
async function getWorkbookMetadata(token: string, itemId: string): Promise<any> {
  // TODO: Implement actual Graph API call
  console.log('[HTTP] Get workbook metadata:', itemId);

  return {
    id: itemId,
    name: 'Sample Workbook',
    worksheets: ['Sheet1'],
  };
}

/**
 * Create HTTP server
 */
const server = http.createServer(async (req, res) => {
  // Handle CORS preflight
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    });
    res.end();
    return;
  }

  // Health check
  if (req.url === '/health') {
    sendJson(res, 200, { status: 'ok' });
    return;
  }

  // Apply workbook changes
  if (req.url === '/microsoft-graph/workbook/apply' && req.method === 'POST') {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader) {
        sendJson(res, 401, { error: 'Missing authorization header' });
        return;
      }

      const token = authHeader.replace('Bearer ', '');
      const body: ApplyChangesRequest = await parseBody(req);

      const result = await applyWorkbookChanges(token, body);
      sendJson(res, 200, result);
    } catch (err) {
      console.error('[HTTP] Error applying changes:', err);
      sendJson(res, 500, { error: err instanceof Error ? err.message : 'Internal error' });
    }
    return;
  }

  // Get workbook metadata
  if (req.url?.startsWith('/microsoft-graph/workbook/metadata') && req.method === 'GET') {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader) {
        sendJson(res, 401, { error: 'Missing authorization header' });
        return;
      }

      const token = authHeader.replace('Bearer ', '');
      const url = new URL(req.url, `http://${req.headers.host}`);
      const itemId = url.searchParams.get('itemId');

      if (!itemId) {
        sendJson(res, 400, { error: 'Missing itemId parameter' });
        return;
      }

      const result = await getWorkbookMetadata(token, itemId);
      sendJson(res, 200, result);
    } catch (err) {
      console.error('[HTTP] Error getting metadata:', err);
      sendJson(res, 500, { error: err instanceof Error ? err.message : 'Internal error' });
    }
    return;
  }

  // 404
  sendJson(res, 404, { error: 'Not found' });
});

// Start server
const port = config.http.port;
server.listen(port, () => {
  console.log(`[HTTP] Microsoft Graph API server listening on http://localhost:${port}`);
  console.log(`[HTTP] Endpoints:`);
  console.log(`  POST /microsoft-graph/workbook/apply`);
  console.log(`  GET  /microsoft-graph/workbook/metadata`);
});

// Handle shutdown
process.on('SIGINT', () => {
  console.log('[HTTP] Shutting down...');
  server.close();
  process.exit(0);
});

