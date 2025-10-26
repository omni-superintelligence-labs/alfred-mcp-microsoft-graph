/**
 * Production Fastify HTTP API for Microsoft Graph workbook operations
 * Implements OBO auth, Graph client, rate limiting, and observability
 */

// Initialize OpenTelemetry before any other imports
import { initializeTracing } from './observability/tracing.js';
initializeTracing();

import Fastify from 'fastify';
import cors from '@fastify/cors';
import helmet from '@fastify/helmet';
import rateLimit from '@fastify/rate-limit';
import { config } from './config.js';
import { createGraphClient, getOBOToken } from './graph/client.js';
import { applyWorkbookOperations } from './graph/workbook.js';
import { validateJWT, issueAddinToken } from './auth/jwt.js';
import { checkRateLimit, checkIdempotencyKey, getIdempotencyResult, storeIdempotencyKey, getRedisClient } from './cache/redis.js';
import { workbookOperationsCounter, operationDurationHistogram } from './observability/tracing.js';
import type { FastifyRequest, FastifyReply } from 'fastify';
import { z } from 'zod';

// Validation schemas
const ApplyRequestSchema = z.object({
  itemId: z.string(),
  driveId: z.string().optional(),
  operations: z.array(z.object({
    type: z.enum(['insert', 'update', 'delete', 'format', 'chart', 'table']),
    target: z.string(),
    data: z.any().optional(),
    options: z.record(z.any()).optional(),
  })),
  clientContext: z.record(z.any()).optional(),
  idempotencyKey: z.string().optional(),
});

// Create Fastify instance
const fastify = Fastify({
  logger: {
    transport: {
      target: 'pino-pretty',
      options: {
        translateTime: 'HH:MM:ss Z',
        ignore: 'pid,hostname',
      },
    },
  },
  requestIdHeader: 'x-request-id',
  requestIdLogLabel: 'reqId',
});

// Register plugins
await fastify.register(helmet, {
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'"],
    },
  },
});

await fastify.register(cors, {
  origin: config.http.corsOrigins,
  credentials: true,
});

await fastify.register(rateLimit, {
  global: true,
  max: 100, // requests per timeWindow
  timeWindow: '1 minute',
  cache: 10000,
  allowList: ['127.0.0.1'],
  redis: undefined, // TODO: Add Redis for distributed rate limiting
  keyGenerator: (req) => {
    // Rate limit per user (from JWT sub claim)
    const auth = req.headers.authorization;
    if (!auth) return req.ip;
    
    try {
      const token = auth.replace('Bearer ', '');
      const decoded = validateJWT(token);
      return `user:${decoded.sub}`;
    } catch {
      return req.ip;
    }
  },
});

// Auth middleware
fastify.addHook('preHandler', async (request, reply) => {
  // Skip auth for health check and token issuance
  if (request.url === '/health' || request.url === '/auth/issue-addin-token') return;

  const auth = request.headers.authorization;
  if (!auth) {
    reply.code(401).send({ error: 'Missing authorization header' });
    return;
  }

  try {
    const token = auth.replace('Bearer ', '');
    const decoded = validateJWT(token);
    
    // Attach user info to request
    (request as any).user = decoded;
  } catch (err) {
    request.log.error({ err }, 'JWT validation failed');
    reply.code(401).send({ error: 'Invalid token' });
  }
});

// Health check
fastify.get('/health', async (request, reply) => {
  return { status: 'ok', timestamp: new Date().toISOString() };
});

// Metrics endpoint (TODO: Prometheus format)
fastify.get('/metrics', async (request, reply) => {
  return {
    uptime: process.uptime(),
    memory: process.memoryUsage(),
    cpu: process.cpuUsage(),
  };
});

// Issue short-lived token for add-in
fastify.post('/auth/issue-addin-token', async (request: FastifyRequest, reply: FastifyReply) => {
  try {
    const { userId, tabId, nonce } = request.body as { userId: string; tabId: string; nonce: string };
    
    if (!userId || !tabId || !nonce) {
      reply.code(400).send({ error: 'Missing required fields: userId, tabId, nonce' });
      return;
    }

    const token = await issueAddinToken(userId, tabId, nonce);
    
    return { token, expiresIn: 600 }; // 10 minutes
  } catch (err) {
    request.log.error({ err }, 'Failed to issue add-in token');
    reply.code(500).send({ error: 'Internal server error' });
  }
});

// Apply workbook operations
fastify.post('/microsoft-graph/workbook/apply', async (request: FastifyRequest, reply: FastifyReply) => {
  const startTime = Date.now();
  
  try {
    // Validate request
    const body = ApplyRequestSchema.parse(request.body);
    const user = (request as any).user;

    // Check Redis-based rate limit
    const rateLimit = await checkRateLimit(user.sub, 100, 60);
    if (!rateLimit.allowed) {
      reply.header('X-RateLimit-Limit', '100');
      reply.header('X-RateLimit-Remaining', '0');
      reply.header('X-RateLimit-Reset', rateLimit.resetAt.toString());
      reply.code(429).send({
        error: 'Too many requests',
        retryAfter: Math.ceil((rateLimit.resetAt - Date.now()) / 1000),
      });
      return;
    }

    reply.header('X-RateLimit-Limit', '100');
    reply.header('X-RateLimit-Remaining', rateLimit.remaining.toString());
    reply.header('X-RateLimit-Reset', rateLimit.resetAt.toString());

    // Check idempotency key
    if (body.idempotencyKey) {
      const isDuplicate = await checkIdempotencyKey(body.idempotencyKey);
      
      if (isDuplicate) {
        const cachedResult = await getIdempotencyResult(body.idempotencyKey);
        if (cachedResult) {
          request.log.info({ idempotencyKey: body.idempotencyKey }, 'Returning cached result');
          return cachedResult;
        }
      }
    }

    request.log.info({
      itemId: body.itemId,
      operationCount: body.operations.length,
      userId: user.sub,
    }, 'Apply workbook operations');

    // Get OBO token for Graph
    const userToken = request.headers.authorization!.replace('Bearer ', '');
    const graphToken = await getOBOToken(userToken);

    // Create Graph client
    const graphClient = createGraphClient(graphToken);

    // Apply operations
    const result = await applyWorkbookOperations(
      graphClient,
      body.itemId,
      body.driveId,
      body.operations,
      body.idempotencyKey
    );

    const duration = Date.now() - startTime;

    request.log.info({
      itemId: body.itemId,
      appliedCount: result.applied.length,
      errorCount: result.errors?.length || 0,
      durationMs: duration,
    }, 'Operations applied');

    const response = {
      success: true,
      applied: result.applied,
      errors: result.errors,
      durationMs: duration,
      sessionId: result.sessionId,
    };

    // Record metrics
    result.applied.forEach(op => {
      workbookOperationsCounter.add(1, { operation_type: op.type });
    });
    
    operationDurationHistogram.record(duration / 1000, {
      itemId: body.itemId,
      operationCount: body.operations.length,
    });

    // Store idempotency result
    if (body.idempotencyKey) {
      await storeIdempotencyKey(body.idempotencyKey, response);
    }

    return response;
  } catch (err: any) {
    const duration = Date.now() - startTime;
    
    request.log.error({
      err,
      durationMs: duration,
    }, 'Failed to apply operations');

    // Handle specific Graph errors
    if (err.statusCode === 429) {
      const retryAfter = err.headers?.['retry-after'] || 60;
      reply.header('Retry-After', retryAfter);
      reply.code(429).send({
        error: 'Rate limited by Microsoft Graph',
        retryAfter: parseInt(retryAfter, 10),
      });
      return;
    }

    if (err.statusCode === 423) {
      reply.code(423).send({
        error: 'Workbook is locked',
        message: 'The workbook is currently locked by another user or process',
      });
      return;
    }

    if (err.statusCode === 409) {
      reply.code(409).send({
        error: 'Conflict',
        message: 'The workbook has been modified. Please refresh and try again',
      });
      return;
    }

    // Generic error
    reply.code(500).send({
      error: 'Internal server error',
      message: err.message,
    });
  }
});

// Get workbook metadata
fastify.get('/microsoft-graph/workbook/metadata', async (request: FastifyRequest, reply: FastifyReply) => {
  try {
    const query = request.query as { itemId?: string };
    
    if (!query.itemId) {
      reply.code(400).send({ error: 'Missing itemId parameter' });
      return;
    }

    const user = (request as any).user;
    request.log.info({
      itemId: query.itemId,
      userId: user.sub,
    }, 'Get workbook metadata');

    // Get OBO token
    const userToken = request.headers.authorization!.replace('Bearer ', '');
    const graphToken = await getOBOToken(userToken);

    // Create Graph client
    const graphClient = createGraphClient(graphToken);

    // Get workbook info
    const workbook = await graphClient
      .api(`/me/drive/items/${query.itemId}/workbook`)
      .get();

    // Get worksheets
    const worksheets = await graphClient
      .api(`/me/drive/items/${query.itemId}/workbook/worksheets`)
      .get();

    return {
      id: query.itemId,
      worksheets: worksheets.value.map((ws: any) => ({
        id: ws.id,
        name: ws.name,
        position: ws.position,
        visibility: ws.visibility,
      })),
    };
  } catch (err: any) {
    request.log.error({ err }, 'Failed to get metadata');
    reply.code(500).send({
      error: 'Failed to get metadata',
      message: err.message,
    });
  }
});

// Initialize Redis
try {
  const redis = getRedisClient();
  await redis.ping();
  console.log('[Fastify] Redis connection verified');
} catch (err) {
  console.warn('[Fastify] Redis connection failed, will retry on demand:', err);
}

// Start server
const port = config.http.port;
const host = '0.0.0.0';

try {
  await fastify.listen({ port, host });
  console.log(`[Fastify] Microsoft Graph API server listening on http://${host}:${port}`);
  console.log(`[Fastify] Endpoints:`);
  console.log(`  POST /microsoft-graph/workbook/apply`);
  console.log(`  GET  /microsoft-graph/workbook/metadata`);
  console.log(`  GET  /health`);
  console.log(`  GET  /metrics`);
} catch (err) {
  fastify.log.error(err);
  process.exit(1);
}

// Graceful shutdown
const signals = ['SIGINT', 'SIGTERM'];
signals.forEach((signal) => {
  process.on(signal, async () => {
    fastify.log.info(`Received ${signal}, closing server`);
    await fastify.close();
    process.exit(0);
  });
});

