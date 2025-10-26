# Microsoft Graph API - Production Deployment Guide

## Prerequisites

1. **Azure Resources**
   - App Registration with OBO configured
   - Client secret stored in Azure Key Vault
   - API permissions granted: Files.ReadWrite.All

2. **Infrastructure**
   - Kubernetes cluster (1.24+)
   - Redis cluster for caching
   - External Secrets Operator
   - Prometheus + Grafana for monitoring

3. **CI/CD**
   - Container registry (ECR, ACR, or GCR)
   - ArgoCD for GitOps deployments

## Azure App Registration Setup

### 1. Create API App Registration

```bash
# Create the API app
az ad app create \
  --display-name "Alfred Graph API - Production" \
  --identifier-uris "api://alfred-graph-api-prod" \
  --sign-in-audience AzureADMyOrg

# Note the Application (client) ID
export API_CLIENT_ID="<client-id>"

# Create a client secret
az ad app credential reset \
  --id $API_CLIENT_ID \
  --append \
  --display-name "Prod Secret"

# Note the client secret (only shown once!)
export API_CLIENT_SECRET="<secret>"
```

### 2. Expose API Scope

In Azure Portal:
1. Go to App Registration → Expose an API
2. Add scope: `access_as_user`
   - Admin consent display name: "Access Graph API as user"
   - Admin consent description: "Allows the app to access Microsoft Graph on behalf of the signed-in user"
   - State: Enabled

### 3. Configure API Permissions

```bash
# Add Microsoft Graph permissions
az ad app permission add \
  --id $API_CLIENT_ID \
  --api 00000003-0000-0000-c000-000000000000 \
  --api-permissions e1fe6dd8-ba31-4d61-89e7-88639da4683d=Scope

# Grant admin consent
az ad app permission admin-consent --id $API_CLIENT_ID
```

## Secrets Management

### 1. Store Secrets in Azure Key Vault

```bash
az keyvault secret set \
  --vault-name alfred-prod-kv \
  --name azure-client-id \
  --value "$API_CLIENT_ID"

az keyvault secret set \
  --vault-name alfred-prod-kv \
  --name azure-client-secret \
  --value "$API_CLIENT_SECRET"
```

### 2. Configure External Secrets

```yaml
# config/external-secrets/graph-api-secrets.yaml
apiVersion: external-secrets.io/v1beta1
kind: ExternalSecret
metadata:
  name: azure-credentials
  namespace: alfred
spec:
  refreshInterval: 1h
  secretStoreRef:
    name: azure-keyvault
    kind: SecretStore
  target:
    name: azure-credentials
    creationPolicy: Owner
  data:
    - secretKey: AZURE_CLIENT_ID
      remoteRef:
        key: azure-client-id
    - secretKey: AZURE_CLIENT_SECRET
      remoteRef:
        key: azure-client-secret
    - secretKey: AZURE_API_CLIENT_ID
      remoteRef:
        key: azure-api-client-id
```

## Build and Push Container

```bash
# Build
docker build -t alfred-mcp-microsoft-graph:latest .

# Tag for registry
docker tag alfred-mcp-microsoft-graph:latest \
  <registry>/alfred-mcp-microsoft-graph:v1.0.0

# Push
docker push <registry>/alfred-mcp-microsoft-graph:v1.0.0
```

## Deploy to Kubernetes

### 1. Create Namespace

```bash
kubectl create namespace alfred
```

### 2. Deploy External Secrets

```bash
kubectl apply -f config/external-secrets/graph-api-secrets.yaml
```

### 3. Deploy via Helm

```bash
helm upgrade --install alfred-graph-api ./helm \
  --namespace alfred \
  --set image.repository=<registry>/alfred-mcp-microsoft-graph \
  --set image.tag=v1.0.0 \
  --set redis.host=redis-master.alfred.svc.cluster.local \
  --values helm/values-prod.yaml
```

### 4. Verify Deployment

```bash
# Check pods
kubectl get pods -n alfred -l app.kubernetes.io/name=alfred-mcp-microsoft-graph

# Check logs
kubectl logs -n alfred -l app.kubernetes.io/name=alfred-mcp-microsoft-graph -f

# Check health
kubectl port-forward -n alfred svc/alfred-mcp-microsoft-graph 3100:3100
curl http://localhost:3100/health
```

## Monitoring Setup

### 1. ServiceMonitor for Prometheus

```yaml
# Automatically created by Helm if serviceMonitor.enabled=true
# Scrapes /metrics endpoint every 30s
```

### 2. Grafana Dashboard

Import dashboard from `monitoring/grafana-dashboard.json`:
- Request rate, latency, error rate
- Graph API call distribution
- 429 throttle rate
- Memory and CPU usage

### 3. Alerts

```yaml
# monitoring/alerts.yaml
groups:
  - name: graph-api
    rules:
      - alert: HighErrorRate
        expr: rate(http_requests_total{status=~"5.."}[5m]) > 0.01
        for: 5m
        labels:
          severity: critical
        annotations:
          summary: "High error rate on Graph API"
      
      - alert: HighLatency
        expr: histogram_quantile(0.95, rate(http_request_duration_seconds_bucket[5m])) > 4
        for: 10m
        labels:
          severity: warning
        annotations:
          summary: "p95 latency above 4s"
      
      - alert: HighThrottleRate
        expr: rate(graph_api_throttled_total[5m]) > 0.1
        for: 5m
        labels:
          severity: warning
        annotations:
          summary: "Microsoft Graph throttling > 10%"
```

## Rollout Strategy

### 1. Canary Deployment (via Argo Rollouts)

```yaml
# helm/templates/rollout.yaml
apiVersion: argoproj.io/v1alpha1
kind: Rollout
metadata:
  name: {{ include "alfred-mcp-microsoft-graph.fullname" . }}
spec:
  replicas: {{ .Values.replicaCount }}
  strategy:
    canary:
      steps:
        - setWeight: 10
        - pause: { duration: 5m }
        - setWeight: 25
        - pause: { duration: 5m }
        - setWeight: 50
        - pause: { duration: 10m }
        - setWeight: 75
        - pause: { duration: 5m }
      analysis:
        templates:
          - templateName: success-rate
        startingStep: 1
        args:
          - name: service-name
            value: alfred-mcp-microsoft-graph
```

### 2. Feature Flags

Environment variables:
```yaml
- name: FEATURE_EXCEL_TABS_ENABLED
  value: "true"
- name: FEATURE_SITES_SELECTED_ENABLED
  value: "false"
```

### 3. Kill Switch

```bash
# Disable Excel tabs immediately
kubectl set env deployment/alfred-mcp-microsoft-graph \
  -n alfred \
  FEATURE_EXCEL_TABS_ENABLED=false
```

## Performance Tuning

### 1. HPA Custom Metrics

```yaml
# Install metrics server and custom metrics adapter
# Configure HPA to scale on request rate
metrics:
  - type: Pods
    pods:
      metric:
        name: graph_api_requests_per_second
      target:
        type: AverageValue
        averageValue: "50"
```

### 2. Connection Pooling

```typescript
// In src/graph/client.ts
const httpAgent = new http.Agent({
  keepAlive: true,
  maxSockets: 50,
  maxFreeSockets: 10,
  timeout: 60000,
});
```

### 3. Redis Configuration

```yaml
# values-prod.yaml
redis:
  enabled: true
  host: redis-master
  port: 6379
  maxConnections: 100
  auth:
    enabled: true
    existingSecret: redis-auth
```

## Troubleshooting

### High 429 Rate

1. Check Graph API quotas in Azure Portal
2. Verify retry logic is working
3. Check logs for retry patterns
4. Consider request batching

### Memory Issues

1. Check for memory leaks: `kubectl top pods -n alfred`
2. Review session cache size
3. Adjust memory limits in values.yaml
4. Check for Redis connection pooling

### Auth Failures

1. Verify secrets are loaded: `kubectl get secret azure-credentials -n alfred -o yaml`
2. Check JWKS endpoint is accessible
3. Verify tenant ID configuration
4. Check token expiration times

## Production Checklist

- [ ] Azure app registration configured
- [ ] Client secret rotated and stored in Key Vault
- [ ] External Secrets operator installed
- [ ] Redis cluster deployed
- [ ] Secrets synced to Kubernetes
- [ ] Container image built and pushed
- [ ] Helm chart deployed
- [ ] Health checks passing
- [ ] Metrics being collected
- [ ] Dashboards configured
- [ ] Alerts configured
- [ ] Runbooks documented
- [ ] Load tested (100 RPS)
- [ ] Security scan passed
- [ ] SBOM generated

## Scaling Guidance

### Expected Load
- 10k DAU
- ~50k Excel operations/day
- Peak: 100 RPS
- Average: 10-20 RPS

### Recommended Configuration
- Min replicas: 2
- Max replicas: 10
- CPU: 500m-2000m
- Memory: 512Mi-2Gi
- Redis: 2GB memory, persistence enabled

## Support Contacts

- **Ops Issues:** ops@chatwithalfred.com
- **Security Issues:** security@chatwithalfred.com
- **On-call Escalation:** Use PagerDuty integration

