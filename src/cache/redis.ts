/**
 * Redis client for caching sessions, idempotency keys, and rate limiting
 */

import Redis from 'ioredis';

const REDIS_HOST = process.env.REDIS_HOST || 'localhost';
const REDIS_PORT = parseInt(process.env.REDIS_PORT || '6379', 10);
const REDIS_PASSWORD = process.env.REDIS_PASSWORD;
const REDIS_DB = parseInt(process.env.REDIS_DB || '0', 10);

let redisClient: Redis | null = null;

/**
 * Get or create Redis client
 */
export function getRedisClient(): Redis {
  if (redisClient) return redisClient;

  redisClient = new Redis({
    host: REDIS_HOST,
    port: REDIS_PORT,
    password: REDIS_PASSWORD,
    db: REDIS_DB,
    retryStrategy: (times) => {
      const delay = Math.min(times * 50, 2000);
      return delay;
    },
    maxRetriesPerRequest: 3,
    enableReadyCheck: true,
    lazyConnect: false,
  });

  redisClient.on('error', (err) => {
    console.error('[Redis] Connection error:', err);
  });

  redisClient.on('connect', () => {
    console.log('[Redis] Connected to Redis');
  });

  redisClient.on('ready', () => {
    console.log('[Redis] Redis client ready');
  });

  return redisClient;
}

/**
 * Cache workbook session ID
 */
export async function cacheSession(
  driveId: string | undefined,
  itemId: string,
  sessionId: string,
  ttlSeconds: number = 300
): Promise<void> {
  const client = getRedisClient();
  const key = `workbookSession:${driveId || 'default'}:${itemId}`;
  
  await client.setex(key, ttlSeconds, sessionId);
  console.log(`[Redis] Cached session: ${key}`);
}

/**
 * Get cached workbook session ID
 */
export async function getCachedSession(
  driveId: string | undefined,
  itemId: string
): Promise<string | null> {
  const client = getRedisClient();
  const key = `workbookSession:${driveId || 'default'}:${itemId}`;
  
  const sessionId = await client.get(key);
  
  if (sessionId) {
    // Renew TTL on access (keep-alive pattern)
    await client.expire(key, 300);
  }
  
  return sessionId;
}

/**
 * Check idempotency key
 */
export async function checkIdempotencyKey(key: string): Promise<boolean> {
  const client = getRedisClient();
  const redisKey = `idem:${key}`;
  
  const exists = await client.exists(redisKey);
  return exists === 1;
}

/**
 * Store idempotency key (24 hour TTL)
 */
export async function storeIdempotencyKey(
  key: string,
  result: any
): Promise<void> {
  const client = getRedisClient();
  const redisKey = `idem:${key}`;
  
  await client.setex(redisKey, 86400, JSON.stringify(result));
  console.log(`[Redis] Stored idempotency key: ${key}`);
}

/**
 * Get cached idempotency result
 */
export async function getIdempotencyResult(key: string): Promise<any | null> {
  const client = getRedisClient();
  const redisKey = `idem:${key}`;
  
  const result = await client.get(redisKey);
  return result ? JSON.parse(result) : null;
}

/**
 * Token bucket rate limiter
 */
export async function checkRateLimit(
  userId: string,
  maxRequests: number = 100,
  windowSeconds: number = 60
): Promise<{ allowed: boolean; remaining: number; resetAt: number }> {
  const client = getRedisClient();
  const key = `rate:${userId}`;
  const now = Date.now();
  
  // Use Redis sorted set for sliding window
  const windowStart = now - (windowSeconds * 1000);
  
  // Remove old entries
  await client.zremrangebyscore(key, 0, windowStart);
  
  // Count current requests
  const count = await client.zcard(key);
  
  if (count >= maxRequests) {
    // Get oldest entry to determine reset time
    const oldest = await client.zrange(key, 0, 0, 'WITHSCORES');
    const resetAt = oldest.length > 1 
      ? parseInt(oldest[1], 10) + (windowSeconds * 1000)
      : now + (windowSeconds * 1000);
    
    return {
      allowed: false,
      remaining: 0,
      resetAt,
    };
  }
  
  // Add current request
  await client.zadd(key, now, `${now}-${Math.random()}`);
  await client.expire(key, windowSeconds);
  
  return {
    allowed: true,
    remaining: maxRequests - count - 1,
    resetAt: now + (windowSeconds * 1000),
  };
}

/**
 * Close Redis connection
 */
export async function closeRedis(): Promise<void> {
  if (redisClient) {
    await redisClient.quit();
    redisClient = null;
    console.log('[Redis] Connection closed');
  }
}

