# Multi-stage build for alfred-mcp-microsoft-graph HTTP API

# Stage 1: Dependencies
FROM node:20-alpine AS deps
WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm ci --only=production

# Stage 2: Builder
FROM node:20-alpine AS builder
WORKDIR /app

COPY package.json package-lock.json* tsconfig.json ./
COPY src ./src

RUN npm ci
RUN npm run build

# Stage 3: Runner
FROM node:20-alpine AS runner
WORKDIR /app

# Security: Run as non-root user
RUN addgroup --system --gid 1001 alfred && \
    adduser --system --uid 1001 alfred

# Copy dependencies and built app
COPY --from=deps --chown=alfred:alfred /app/node_modules ./node_modules
COPY --from=builder --chown=alfred:alfred /app/dist ./dist
COPY --from=builder --chown=alfred:alfred /app/package.json ./

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD node -e "require('http').get('http://localhost:3100/health', (r) => { process.exit(r.statusCode === 200 ? 0 : 1); })"

# Switch to non-root user
USER alfred

# Expose port
EXPOSE 3100

# Set production environment
ENV NODE_ENV=production

# Start the HTTP API server
CMD ["node", "dist/fastify-server.js"]

