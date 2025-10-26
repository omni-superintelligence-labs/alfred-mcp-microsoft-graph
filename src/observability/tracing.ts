/**
 * OpenTelemetry instrumentation setup
 */

import { NodeSDK } from '@opentelemetry/sdk-node';
import { getNodeAutoInstrumentations } from '@opentelemetry/auto-instrumentations-node';
import { PrometheusExporter } from '@opentelemetry/exporter-prometheus';
import { Resource } from '@opentelemetry/resources';
import { SemanticResourceAttributes } from '@opentelemetry/semantic-conventions';

const prometheusExporter = new PrometheusExporter({
  port: 9464, // Prometheus metrics port
  endpoint: '/metrics',
});

export const sdk = new NodeSDK({
  resource: new Resource({
    [SemanticResourceAttributes.SERVICE_NAME]: 'alfred-graph-api',
    [SemanticResourceAttributes.SERVICE_VERSION]: process.env.npm_package_version || '1.0.0',
    [SemanticResourceAttributes.DEPLOYMENT_ENVIRONMENT]: process.env.NODE_ENV || 'development',
  }),
  metricReader: prometheusExporter,
  instrumentations: [
    getNodeAutoInstrumentations({
      '@opentelemetry/instrumentation-fs': { enabled: false }, // Too verbose
      '@opentelemetry/instrumentation-http': { enabled: true },
      '@opentelemetry/instrumentation-fastify': { enabled: true },
      '@opentelemetry/instrumentation-ioredis': { enabled: true },
    }),
  ],
});

/**
 * Initialize OpenTelemetry
 */
export function initializeTracing(): void {
  sdk.start();
  console.log('[OpenTelemetry] Tracing initialized');
  console.log('[OpenTelemetry] Prometheus metrics available on :9464/metrics');

  // Graceful shutdown
  process.on('SIGTERM', () => {
    sdk
      .shutdown()
      .then(() => console.log('[OpenTelemetry] Shut down successfully'))
      .catch((error) => console.error('[OpenTelemetry] Error shutting down', error))
      .finally(() => process.exit(0));
  });
}

/**
 * Custom metrics for business logic
 */
import { metrics } from '@opentelemetry/api';

const meter = metrics.getMeter('alfred-graph-api');

// Counters
export const workbookOperationsCounter = meter.createCounter('workbook_operations_total', {
  description: 'Total workbook operations by type',
});

export const graphApiCallsCounter = meter.createCounter('graph_api_calls_total', {
  description: 'Total Graph API calls',
});

export const throttledRequestsCounter = meter.createCounter('graph_api_throttled_total', {
  description: 'Total requests throttled by Graph (429)',
});

export const retriesCounter = meter.createCounter('graph_api_retry_total', {
  description: 'Total retry attempts',
});

// Histograms
export const operationDurationHistogram = meter.createHistogram('workbook_operation_duration_seconds', {
  description: 'Duration of workbook operations',
});

export const graphApiDurationHistogram = meter.createHistogram('graph_api_call_duration_seconds', {
  description: 'Duration of Graph API calls',
});

