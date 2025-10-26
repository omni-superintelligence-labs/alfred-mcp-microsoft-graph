/**
 * k6 load test for /workbook/apply endpoint
 * Run with: k6 run tests/load/apply-operations.js
 */

import http from 'k6/http';
import { check, sleep } from 'k6';
import { Rate, Trend } from 'k6/metrics';

// Custom metrics
const errorRate = new Rate('errors');
const applyDuration = new Trend('apply_duration');
const throttleRate = new Rate('throttled_requests');

// Test configuration
export const options = {
  stages: [
    { duration: '30s', target: 10 },  // Ramp up to 10 users
    { duration: '1m', target: 50 },   // Ramp up to 50 users
    { duration: '2m', target: 50 },   // Stay at 50 users
    { duration: '30s', target: 100 }, // Spike to 100 users
    { duration: '1m', target: 100 },  // Stay at 100 users
    { duration: '30s', target: 0 },   // Ramp down
  ],
  thresholds: {
    'http_req_duration': ['p(95)<4000'], // 95% of requests under 4s
    'errors': ['rate<0.01'],              // Error rate under 1%
    'throttled_requests': ['rate<0.1'],   // Throttle rate under 10%
  },
};

const BASE_URL = __ENV.API_URL || 'http://localhost:3100';
const AUTH_TOKEN = __ENV.AUTH_TOKEN || 'mock-token-for-testing';

export default function () {
  const payload = JSON.stringify({
    itemId: `test-workbook-${__VU}-${__ITER}`,
    operations: [
      {
        type: 'insert',
        target: 'A1:B2',
        data: [
          ['Name', 'Value'],
          ['Test', 123],
        ],
        options: { worksheet: 'Sheet1' },
      },
      {
        type: 'format',
        target: 'A1:B1',
        data: {
          font: { bold: true },
        },
        options: { worksheet: 'Sheet1' },
      },
    ],
    idempotencyKey: `idem-${__VU}-${__ITER}`,
  });

  const params = {
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${AUTH_TOKEN}`,
    },
  };

  const startTime = new Date().getTime();
  const response = http.post(`${BASE_URL}/microsoft-graph/workbook/apply`, payload, params);
  const duration = new Date().getTime() - startTime;

  // Record metrics
  applyDuration.add(duration);

  // Check response
  const success = check(response, {
    'status is 200': (r) => r.status === 200,
    'has success field': (r) => JSON.parse(r.body).success !== undefined,
    'duration under 4s': () => duration < 4000,
  });

  if (!success) {
    errorRate.add(1);
  } else {
    errorRate.add(0);
  }

  // Track throttling
  if (response.status === 429) {
    throttleRate.add(1);
    
    // Respect Retry-After header
    const retryAfter = parseInt(response.headers['Retry-After'] || '1', 10);
    sleep(retryAfter);
  } else {
    throttleRate.add(0);
    
    // Normal think time between requests
    sleep(1);
  }
}

export function handleSummary(data) {
  return {
    'summary.json': JSON.stringify(data, null, 2),
    stdout: textSummary(data, { indent: ' ', enableColors: true }),
  };
}

function textSummary(data, options) {
  const { indent, enableColors } = options;
  let summary = '\n';
  
  summary += `${indent}✓ checks........................: ${data.metrics.checks.values.rate * 100}% ✓ ${data.metrics.checks.values.passes} ✗ ${data.metrics.checks.values.fails}\n`;
  summary += `${indent}✓ http_req_duration.............: avg=${data.metrics.http_req_duration.values.avg.toFixed(2)}ms p(95)=${data.metrics.http_req_duration.values['p(95)'].toFixed(2)}ms\n`;
  summary += `${indent}✓ apply_duration................: avg=${data.metrics.apply_duration.values.avg.toFixed(2)}ms p(95)=${data.metrics.apply_duration.values['p(95)'].toFixed(2)}ms\n`;
  summary += `${indent}✓ errors........................: ${(data.metrics.errors.values.rate * 100).toFixed(2)}%\n`;
  summary += `${indent}✓ throttled_requests............: ${(data.metrics.throttled_requests.values.rate * 100).toFixed(2)}%\n`;
  
  return summary;
}

