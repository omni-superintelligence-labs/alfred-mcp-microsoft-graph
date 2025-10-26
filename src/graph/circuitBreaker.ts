/**
 * Circuit breaker for Microsoft Graph API calls
 * Prevents cascading failures during Graph API outages
 */

import CircuitBreaker from 'opossum';

const CIRCUIT_BREAKER_OPTIONS = {
  timeout: 30000, // 30s timeout
  errorThresholdPercentage: 50, // Open circuit if 50% error rate
  resetTimeout: 30000, // Try to close after 30s
  rollingCountTimeout: 10000, // 10s rolling window
  rollingCountBuckets: 10,
  volumeThreshold: 10, // Need at least 10 requests before opening
};

// Cache circuit breakers per endpoint
const breakers = new Map<string, CircuitBreaker>();

/**
 * Get or create circuit breaker for an endpoint
 */
export function getCircuitBreaker(name: string): CircuitBreaker {
  if (breakers.has(name)) {
    return breakers.get(name)!;
  }

  const breaker = new CircuitBreaker(
    async (fn: () => Promise<any>) => fn(),
    {
      ...CIRCUIT_BREAKER_OPTIONS,
      name,
    }
  );

  // Event listeners
  breaker.on('open', () => {
    console.error(`[CircuitBreaker] ${name} opened (too many failures)`);
  });

  breaker.on('halfOpen', () => {
    console.warn(`[CircuitBreaker] ${name} half-open (testing if recovered)`);
  });

  breaker.on('close', () => {
    console.log(`[CircuitBreaker] ${name} closed (service recovered)`);
  });

  breaker.on('fallback', (result) => {
    console.warn(`[CircuitBreaker] ${name} fallback triggered`);
  });

  breakers.set(name, breaker);
  return breaker;
}

/**
 * Wrap Graph API call with circuit breaker
 */
export async function protectedGraphCall<T>(
  name: string,
  fn: () => Promise<T>,
  fallback?: () => Promise<T>
): Promise<T> {
  const breaker = getCircuitBreaker(name);

  if (fallback) {
    breaker.fallback(fallback);
  }

  return breaker.fire(fn);
}

/**
 * Get circuit breaker stats
 */
export function getCircuitBreakerStats(): Record<string, any> {
  const stats: Record<string, any> = {};

  for (const [name, breaker] of breakers) {
    stats[name] = {
      state: breaker.opened ? 'open' : breaker.halfOpen ? 'half-open' : 'closed',
      stats: breaker.stats,
    };
  }

  return stats;
}

