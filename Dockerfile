# ═══════════════════════════════════════════════════════════════════════════════
# Excel AI Copilot — Multi-stage Dockerfile
#
# Stage 1 (frontend-build): Node.js — compiles the React/Webpack add-in.
#   FRONTEND_URL is baked into manifest.xml at build time.
#   Pass it as a build arg:
#     docker build --build-arg FRONTEND_URL=https://your-app.example.com .
#
# Stage 2 (final): Python — runs the FastAPI backend.
#   The built frontend (dist/) is copied to ./static and served as static files.
#   All secrets/keys are injected at runtime via environment variables.
#
# Usage:
#   Build:  docker build --build-arg FRONTEND_URL=https://your-app.example.com -t excel-ai-copilot .
#   Run:    docker run -p 8080:8080 -e LLM_API_KEY=sk-... excel-ai-copilot
# ═══════════════════════════════════════════════════════════════════════════════

# ── Stage 1: Build frontend ───────────────────────────────────────────────────
FROM node:20-alpine AS frontend-build

WORKDIR /build

# Install npm dependencies first (separate layer for caching)
COPY frontend/package*.json ./
RUN npm ci --prefer-offline

# Copy source
COPY frontend/ ./

# FRONTEND_URL is baked into manifest.xml by webpack's CopyWebpackPlugin transform.
# Default keeps localhost:3000 so a plain `docker build .` still works for testing.
ARG FRONTEND_URL=https://localhost:3000
ENV FRONTEND_URL=$FRONTEND_URL

RUN npm run build

# ── Stage 2: Python backend ───────────────────────────────────────────────────
FROM python:3.11-slim AS final

WORKDIR /app

# Install Python dependencies
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy backend source
COPY backend/ .

# Copy built frontend into ./static (FastAPI serves this at "/" in OpenShift mode)
COPY --from=frontend-build /build/dist ./static

# OpenShift runs containers as an arbitrary non-root UID.
# Making /app group-writable ensures the app can write logs/temp files.
RUN chmod -R g+rwX /app

# ── Default runtime environment ───────────────────────────────────────────────
# These are production defaults — override at runtime via env vars or OpenShift
# Secrets/ConfigMaps. LLM_API_KEY must be set at runtime (never hardcoded).
ENV OPENSHIFT=true \
    SERVE_STATIC=true \
    STATIC_DIR=./static \
    PORT=8080 \
    DEBUG=false \
    LLM_MODEL=claude-sonnet-4-20250514 \
    LLM_API_KEY="" \
    LLM_BASE_URL="" \
    LLM_MAX_TOKENS=4096 \
    LLM_TEMPERATURE=0.1

EXPOSE 8080

# Use sh -c so $PORT expansion works
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080}"]
