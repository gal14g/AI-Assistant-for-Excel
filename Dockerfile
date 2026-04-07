# ═══════════════════════════════════════════════════════════════════════════════
# AI Assistant For Excel — Multi-stage Dockerfile
#
# Stage 1 (frontend-build): Node.js — compiles the React/Webpack add-in.
#   FRONTEND_URL is baked into manifest.xml at build time.
#   Pass it as a build arg:
#     docker build --build-arg FRONTEND_URL=https://your-app.example.com .
#
# Stage 2 (deps): Python — installs dependencies into a virtual env.
#
# Stage 3 (final): Python slim — runs the FastAPI backend + serves frontend.
#   All secrets/keys are injected at runtime via environment variables.
#
# Usage:
#   Build:  docker build --build-arg FRONTEND_URL=https://your-app.example.com -t ai-assistant-for-excel .
#   Run:    docker run -p 8080:8080 -e LLM_API_KEY=sk-... ai-assistant-for-excel
# ═══════════════════════════════════════════════════════════════════════════════

# ── Stage 1: Build frontend ───────────────────────────────────────────────────
FROM node:20-alpine AS frontend-build

WORKDIR /build

# Install npm dependencies first (separate layer for caching)
COPY frontend/package*.json ./
RUN npm ci --no-audit --no-fund

# Copy source and build
COPY frontend/ ./

ARG FRONTEND_URL=https://localhost:3000
ENV FRONTEND_URL=$FRONTEND_URL

# For enclosed networks: pass --build-arg OFFICE_JS_SRC=/assets/office.js
# AND drop a downloaded office.js into frontend/public/assets/ beforehand.
ARG OFFICE_JS_SRC=
ENV OFFICE_JS_SRC=$OFFICE_JS_SRC

RUN npm run build

# ── Stage 2: Install Python dependencies ─────────────────────────────────────
FROM python:3.11-slim AS deps

WORKDIR /deps

COPY backend/requirements.txt .
RUN python -m venv /deps/venv && \
    /deps/venv/bin/pip install --no-cache-dir --upgrade pip && \
    /deps/venv/bin/pip install --no-cache-dir -r requirements.txt

# ── Stage 3: Final production image ──────────────────────────────────────────
FROM python:3.11-slim AS final

LABEL org.opencontainers.image.source="https://github.com/your-org/ai-assistant-for-excel"
LABEL org.opencontainers.image.description="AI Assistant For Excel - Natural language spreadsheet assistant"
LABEL org.opencontainers.image.version="1.1.0"

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PATH="/app/venv/bin:$PATH"

WORKDIR /app

# Copy pre-built virtual environment from deps stage
COPY --from=deps /deps/venv ./venv

# Copy backend source
COPY backend/ .

# Copy built frontend into ./static (FastAPI serves this at "/" in OpenShift mode)
COPY --from=frontend-build /build/dist ./static

# Create data directory and set permissions for OpenShift (arbitrary non-root UID)
RUN mkdir -p /app/data && \
    chmod -R g+rwX /app

# ── Default runtime environment ───────────────────────────────────────────────
# These are production defaults — override at runtime via env vars or OpenShift
# Secrets/ConfigMaps. LLM_API_KEY must be set at runtime (never hardcoded).
ENV OPENSHIFT=true \
    SERVE_STATIC=true \
    STATIC_DIR=./static \
    PORT=8080 \
    DEBUG=false \
    LLM_MODEL=gpt-4o \
    LLM_API_KEY="" \
    LLM_BASE_URL="" \
    LLM_MAX_TOKENS=4096 \
    LLM_TEMPERATURE=0.1 \
    ANONYMIZED_TELEMETRY=False \
    HF_HUB_OFFLINE=1 \
    TRANSFORMERS_OFFLINE=1

VOLUME ["/app/data"]

EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8080/health')" || exit 1

# Use sh -c so $PORT expansion works
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080} --workers 1"]
