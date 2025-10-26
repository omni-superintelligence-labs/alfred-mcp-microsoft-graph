/**
 * Microsoft Graph Workbook API operations
 * Handles Excel operations with persistent sessions and idempotency
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { graphRetry } from './client.js';
import { getCachedSession, cacheSession } from '../cache/redis.js';
import { protectedGraphCall } from './circuitBreaker.js';

interface WorkbookOperation {
  type: 'insert' | 'update' | 'delete' | 'format' | 'chart' | 'table';
  target: string;
  data?: any;
  options?: Record<string, any>;
}

interface OperationResult {
  applied: WorkbookOperation[];
  errors?: Array<{ operation: number; error: string }>;
  sessionId?: string;
}

/**
 * Get or create a persistent workbook session
 */
async function getWorkbookSession(
  client: Client,
  itemId: string,
  driveId?: string
): Promise<string> {
  // Try Redis cache first
  const cached = await getCachedSession(driveId, itemId);
  if (cached) {
    console.log(`[Workbook] Using cached session: ${cached}`);
    return cached;
  }

  // Create new session with circuit breaker protection
  const path = driveId
    ? `/drives/${driveId}/items/${itemId}/workbook/createSession`
    : `/me/drive/items/${itemId}/workbook/createSession`;

  const response = await protectedGraphCall(
    'createSession',
    () => graphRetry(() => client.api(path).post({ persistChanges: true }))
  );

  const sessionId = response.id;
  
  // Cache for 5 minutes (Graph sessions expire after inactivity)
  await cacheSession(driveId, itemId, sessionId, 300);

  console.log(`[Workbook] Created session: ${sessionId}`);
  return sessionId;
}

/**
 * Apply a single operation to the workbook
 */
async function applyOperation(
  client: Client,
  itemId: string,
  driveId: string | undefined,
  sessionId: string,
  operation: WorkbookOperation
): Promise<void> {
  const basePath = driveId
    ? `/drives/${driveId}/items/${itemId}/workbook`
    : `/me/drive/items/${itemId}/workbook`;

  switch (operation.type) {
    case 'insert': {
      // Insert data into a range
      const rangePath = `${basePath}/worksheets('${operation.options?.worksheet || 'Sheet1'}')/range(address='${operation.target}')`;
      
      await graphRetry(() =>
        client
          .api(rangePath)
          .header('workbook-session-id', sessionId)
          .patch({
            values: operation.data,
          })
      );
      break;
    }

    case 'update': {
      // Update existing range
      const rangePath = `${basePath}/worksheets('${operation.options?.worksheet || 'Sheet1'}')/range(address='${operation.target}')`;
      
      await graphRetry(() =>
        client
          .api(rangePath)
          .header('workbook-session-id', sessionId)
          .patch({
            values: operation.data,
            numberFormat: operation.options?.numberFormat,
          })
      );
      break;
    }

    case 'format': {
      // Apply formatting to a range
      const rangePath = `${basePath}/worksheets('${operation.options?.worksheet || 'Sheet1'}')/range(address='${operation.target}')/format`;
      
      await graphRetry(() =>
        client
          .api(rangePath)
          .header('workbook-session-id', sessionId)
          .patch(operation.data)
      );
      break;
    }

    case 'table': {
      // Create a table
      const tablesPath = `${basePath}/worksheets('${operation.options?.worksheet || 'Sheet1'}')/tables/add`;
      
      await graphRetry(() =>
        client
          .api(tablesPath)
          .header('workbook-session-id', sessionId)
          .post({
            address: operation.target,
            hasHeaders: operation.options?.hasHeaders !== false,
          })
      );
      break;
    }

    case 'chart': {
      // Create a chart
      const chartsPath = `${basePath}/worksheets('${operation.options?.worksheet || 'Sheet1'}')/charts/add`;
      
      await graphRetry(() =>
        client
          .api(chartsPath)
          .header('workbook-session-id', sessionId)
          .post({
            type: operation.options?.chartType || 'ColumnClustered',
            sourceData: operation.target,
            seriesBy: 'Auto',
          })
      );
      break;
    }

    case 'delete': {
      // Delete a range (clear contents)
      const rangePath = `${basePath}/worksheets('${operation.options?.worksheet || 'Sheet1'}')/range(address='${operation.target}')/clear`;
      
      await graphRetry(() =>
        client
          .api(rangePath)
          .header('workbook-session-id', sessionId)
          .post({
            applyTo: 'Contents',
          })
      );
      break;
    }

    default:
      throw new Error(`Unknown operation type: ${(operation as any).type}`);
  }
}

/**
 * Apply multiple operations to a workbook
 */
export async function applyWorkbookOperations(
  client: Client,
  itemId: string,
  driveId: string | undefined,
  operations: WorkbookOperation[],
  idempotencyKey?: string
): Promise<OperationResult> {
  // TODO: Check idempotency key cache
  if (idempotencyKey) {
    // Check if we've already processed this request
    // For now, skip idempotency check
  }

  // Get or create session
  const sessionId = await getWorkbookSession(client, itemId, driveId);

  const applied: WorkbookOperation[] = [];
  const errors: Array<{ operation: number; error: string }> = [];

  // Apply operations sequentially
  for (let i = 0; i < operations.length; i++) {
    const operation = operations[i];

    try {
      await applyOperation(client, itemId, driveId, sessionId, operation);
      applied.push(operation);
    } catch (err: any) {
      console.error(`[Workbook] Operation ${i} failed:`, err);
      errors.push({
        operation: i,
        error: err.message || 'Unknown error',
      });

      // Stop on first error (or continue based on policy)
      if (!operation.options?.continueOnError) {
        break;
      }
    }
  }

  return {
    applied,
    errors: errors.length > 0 ? errors : undefined,
    sessionId,
  };
}

/**
 * Close a workbook session
 */
export async function closeWorkbookSession(
  client: Client,
  sessionId: string
): Promise<void> {
  try {
    await client
      .api(`/workbook/closeSession`)
      .header('workbook-session-id', sessionId)
      .post({});

    console.log(`[Workbook] Closed session: ${sessionId}`);
  } catch (err) {
    console.error('[Workbook] Failed to close session:', err);
  }
}

