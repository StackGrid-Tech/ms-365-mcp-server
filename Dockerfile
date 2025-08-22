# ---------- Builder stage ----------
FROM node:22-slim AS builder

# Install build dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    make \
    g++ \
    libsecret-1-dev \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy package files first (for better caching)
COPY package*.json ./

# Install all dependencies (including dev)
RUN npm ci

# Copy rest of the source
COPY . .

# Build the project
RUN npm run generate && npm run build

# Remove devDependencies for smaller image
RUN npm prune --production && npm cache clean --force

# ---------- Production stage ----------
FROM node:22-slim AS production

# Create non-root user
RUN apt-get update && apt-get install -y --no-install-recommends libsecret-1-dev && rm -rf /var/lib/apt/lists/* && \
    addgroup --system nodejs && \
    adduser --system --uid 1001 --ingroup nodejs mcp

WORKDIR /app

COPY --from=builder --chown=mcp:nodejs /app/dist ./dist
COPY --from=builder --chown=mcp:nodejs /app/node_modules ./node_modules
COPY --from=builder --chown=mcp:nodejs /app/package.json ./package.json
COPY --from=builder --chown=mcp:nodejs /app/.env ./.env

USER mcp

EXPOSE 3000
CMD ["sh", "-c", "node dist/index.js --org-mode --http ${PORT:-3000}"]
