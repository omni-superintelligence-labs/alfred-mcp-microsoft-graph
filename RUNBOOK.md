# Microsoft Graph API - Operations Runbook

## Common Issues and Solutions

### 1. High 429 Throttle Rate

**Symptoms:**
- Dashboard shows throttle rate > 10%
- User reports "Rate limited" errors
- Logs show repeated 429 responses

**Diagnosis:**
```bash
# Check throttle rate
kubectl logs -n alfred -l app=alfred-mcp-microsoft-graph | grep "429"

# Check retry patterns
kubectl logs -n alfred -l app=alfred-mcp-microsoft-graph | grep "GraphRetry"
```

**Solution:**
1. Verify retry logic is working (exponential backoff)
2. Check if batch operations can be consolidated
3. Review Graph API quotas in Azure Portal
4. Consider implementing request queuing
5. Temporary: Scale up pods to distribute load

**Prevention:**
- Implement pre-emptive rate limiting (90% of quota)
- Add request batching for similar operations
- Cache workbook metadata aggressively

---

### 2. OBO Token Exchange Failures

**Symptoms:**
- 401 errors from Graph API
- "Failed to acquire OBO token" in logs
- Users see "Authentication failed" in add-in

**Diagnosis:**
```bash
# Check JWT validation errors
kubectl logs -n alfred -l app=alfred-mcp-microsoft-graph | grep "JWT"

# Verify secrets are loaded
kubectl get secret azure-credentials -n alfred -o jsonpath='{.data.AZURE_CLIENT_SECRET}' | base64 -d
```

**Solution:**
1. Verify client secret is correct and not expired
2. Check JWKS endpoint is accessible
3. Verify API scope is configured correctly
4. Check tenant ID matches token issuer
5. Rotate secret if compromised

**Prevention:**
- Implement secret rotation alerts (90 days)
- Monitor OBO success rate
- Log detailed OBO errors (sanitized)

---

### 3. High Latency (p95 > 4s)

**Symptoms:**
- Dashboard shows p95 latency above threshold
- Users report slow Excel operations
- Timeout errors in logs

**Diagnosis:**
```bash
# Check Graph API latency
kubectl logs -n alfred -l app=alfred-mcp-microsoft-graph | grep "durationMs"

# Check for slow queries
kubectl logs -n alfred -l app=alfred-mcp-microsoft-graph | grep "durationMs" | awk '{print $NF}' | sort -n | tail -20
```

**Solution:**
1. Check if Graph API is slow (external issue)
2. Verify persistent sessions are working
3. Review operation batching
4. Check Redis latency
5. Scale up pods if CPU-bound

**Prevention:**
- Implement operation batching
- Cache workbook metadata
- Use persistent Excel sessions
- Monitor external dependencies

---

### 4. Memory Leaks

**Symptoms:**
- Pods restarting due to OOM
- Memory usage trending upward
- Slow degradation over time

**Diagnosis:**
```bash
# Check memory usage
kubectl top pods -n alfred -l app=alfred-mcp-microsoft-graph

# Get heap snapshots (if enabled)
kubectl exec -n alfred alfred-mcp-microsoft-graph-xxx -- node --expose-gc -e "global.gc(); console.log(process.memoryUsage())"
```

**Solution:**
1. Restart affected pods
2. Review session cache size limits
3. Check for uncleared interval timers
4. Verify Graph client connection pooling
5. Take heap snapshot and analyze

**Prevention:**
- Implement session cache eviction (LRU)
- Set max cache size limits
- Monitor memory trends
- Regular heap profiling in staging

---

### 5. Redis Connection Failures

**Symptoms:**
- "Redis connection failed" in logs
- Fallback to in-memory cache
- Session cache misses

**Diagnosis:**
```bash
# Check Redis status
kubectl get pods -n alfred -l app=redis

# Test connection
kubectl exec -n alfred alfred-mcp-microsoft-graph-xxx -- wget -O- http://redis-master:6379
```

**Solution:**
1. Verify Redis pods are healthy
2. Check network policies allow connection
3. Verify Redis auth secret is correct
4. Restart Redis if needed (with caution)

**Prevention:**
- Monitor Redis health
- Implement connection retry logic
- Alert on cache miss rate > 20%

---

### 6. Workbook Session Errors

**Symptoms:**
- "Session not found" errors
- Frequent session recreation
- Increased latency

**Diagnosis:**
```bash
# Check session cache hits
kubectl logs -n alfred -l app=alfred-mcp-microsoft-graph | grep "session"
```

**Solution:**
1. Verify session caching is working
2. Check session expiration times (default 5min)
3. Ensure session IDs are preserved across operations
4. Review concurrent operation handling

**Prevention:**
- Extend session timeout if allowed
- Implement session keep-alive
- Monitor session recreation rate

---

## Emergency Procedures

### Kill Switch (Disable Feature)

```bash
# Disable Excel tabs immediately
kubectl set env deployment/alfred-mcp-microsoft-graph -n alfred FEATURE_EXCEL_TABS_ENABLED=false

# Verify
kubectl rollout status deployment/alfred-mcp-microsoft-graph -n alfred
```

### Force Rollback

```bash
# Rollback to previous version
kubectl rollout undo deployment/alfred-mcp-microsoft-graph -n alfred

# Check status
kubectl rollout status deployment/alfred-mcp-microsoft-graph -n alfred
```

### Scale Down (Traffic Surge)

```bash
# Temporarily reduce max replicas to prevent resource exhaustion
kubectl patch hpa alfred-mcp-microsoft-graph -n alfred -p '{"spec":{"maxReplicas":5}}'
```

### Drain and Restart

```bash
# Restart all pods (rolling)
kubectl rollout restart deployment/alfred-mcp-microsoft-graph -n alfred

# Force delete pod (if stuck)
kubectl delete pod alfred-mcp-microsoft-graph-xxx -n alfred --grace-period=0 --force
```

---

## Metrics Reference

### RED Metrics
- **Rate:** `rate(http_requests_total[5m])`
- **Errors:** `rate(http_requests_total{status=~"5.."}[5m])`
- **Duration:** `histogram_quantile(0.95, rate(http_request_duration_seconds_bucket[5m]))`

### Business Metrics
- **Excel Operations:** `rate(workbook_operations_total[5m])`
- **Graph API Calls:** `rate(graph_api_calls_total[5m])`
- **Cache Hit Rate:** `redis_hits / (redis_hits + redis_misses)`

### SLIs
- **Availability:** `(1 - errors/requests) * 100`
- **Latency:** `p95(http_request_duration)`
- **Success Rate:** `(requests - errors) / requests * 100`

---

## On-Call Procedures

### 1. Alert: High Error Rate

1. Check dashboard for spike timing
2. Review logs for error patterns
3. Check external dependencies (Graph API status)
4. If isolated issue: investigate and fix
5. If widespread: activate kill switch

### 2. Alert: High Latency

1. Check Graph API status page
2. Review concurrent request count
3. Check database/Redis latency
4. Consider scaling up temporarily
5. Review recent deployments

### 3. Alert: Pod Crash Loop

1. Get pod logs: `kubectl logs pod-name -n alfred --previous`
2. Check for OOM killer
3. Review recent config changes
4. Rollback if needed
5. Scale up resources if legitimate load

---

## Maintenance Windows

### Weekly Tasks
- Review error logs for patterns
- Check secret expiration dates
- Verify backup status (Redis)
- Review performance trends

### Monthly Tasks
- Rotate client secrets (if approaching 90 days)
- Review and update HPA settings
- Audit access logs
- Update dependencies

### Quarterly Tasks
- Load test with expected growth
- Penetration testing
- Disaster recovery drill
- Review and update runbooks

---

## Contacts

- **Primary On-Call:** Check PagerDuty schedule
- **Escalation:** ops-lead@chatwithalfred.com
- **Microsoft Support:** Azure Premier Support (if applicable)
- **Security Incidents:** security@chatwithalfred.com

---

## Related Documentation

- [DEPLOYMENT.md](./DEPLOYMENT.md) - Deployment procedures
- [SETUP.md](./SETUP.md) - Development setup
- [Architecture Diagrams](./docs/architecture.md)
- [API Documentation](./docs/api.md)

