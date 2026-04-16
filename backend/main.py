"""
AI Assistant For Excel – FastAPI Backend

Entry point for the backend server.

Local dev:   uvicorn main:app --reload --port 8000
Docker/OCP:  uvicorn main:app --host 0.0.0.0 --port 8080
             (or just run the container — CMD handles this)

Deployment modes (controlled via environment variables):
  OPENSHIFT=false  (default) — local dev, no static file serving
  OPENSHIFT=true             — production, serves built frontend from ./static
  SERVE_STATIC=true          — serve static files regardless of OPENSHIFT flag
"""

import logging
import os

from fastapi import FastAPI, Request, Response
from fastapi.middleware.cors import CORSMiddleware
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.errors import RateLimitExceeded
from slowapi.util import get_remote_address
from starlette.middleware.base import BaseHTTPMiddleware

from app.config import settings
from app.routers import chat, feedback, conversations
from app.routers import analyze

logger = logging.getLogger(__name__)

# ── Rate limiter ─────────────────────────────────────────────────────────────
limiter = Limiter(key_func=get_remote_address)

app = FastAPI(
    title="AI Assistant For Excel API",
    description="Backend for the AI Assistant For Excel add-in. Provides LLM-powered plan generation and validation.",
    version="1.1.0",
    docs_url="/docs" if settings.debug else None,
    redoc_url="/redoc" if settings.debug else None,
)

app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)


# ── Security headers middleware ──────────────────────────────────────────────
class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response: Response = await call_next(request)
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
        response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
        if settings.openshift:
            response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
            response.headers["Content-Security-Policy"] = (
                "default-src 'self'; "
                "script-src 'self' https://appsforoffice.microsoft.com; "
                "style-src 'self' 'unsafe-inline'; "
                "connect-src 'self'"
            )
        return response


app.add_middleware(SecurityHeadersMiddleware)

# ── CORS ──────────────────────────────────────────────────────────────────────
# Always use explicit origins — never wildcard in production.
# In OpenShift the frontend is served from the same origin, but Office.js
# iframes may require CORS, so use the configured origins list.
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PATCH", "DELETE"],
    allow_headers=["Content-Type", "X-API-Key", "Authorization"],
)

# ── API routers ───────────────────────────────────────────────────────────────
app.include_router(chat.router)
app.include_router(analyze.router)
app.include_router(feedback.router)
app.include_router(conversations.router)


@app.on_event("startup")
async def startup_event():
    """Initialise vector stores, example store, and feedback database."""
    from app.persistence.factory import get_repositories, get_vector_store
    from app.services.capability_store import init_store
    from app.services.example_store import init_example_store

    # Relational: SQLite (default) or Postgres, per settings.database_url.
    await get_repositories().initialize()
    # Vector: ChromaDB (default) or pgvector, per settings.vector_store_url.
    get_vector_store().initialize()

    # Seed the capability + example collections (idempotent).
    init_store()
    await init_example_store()


@app.on_event("shutdown")
async def shutdown_event():
    """Release persistence resources on shutdown.

    Both the relational repositories and the vector store own long-lived
    resources:
      - `SqliteRepositories` — an aiosqlite connection
      - `PostgresRepositories` — an asyncpg pool
      - `PgVectorStore`       — an asyncpg pool + a background event-loop thread

    Skipping the vector-store close leaks the background thread + pool, which
    Kubernetes then has to SIGKILL on pod termination (truncating in-flight
    writes). Each close is guarded so one failure can't mask the others.
    """
    from app.persistence.factory import get_repositories, get_vector_store

    try:
        await get_repositories().close()
    except Exception:  # pragma: no cover — shutdown best-effort
        logger.exception("Error closing repositories during shutdown")

    try:
        close = getattr(get_vector_store(), "close", None)
        if callable(close):
            close()
    except Exception:  # pragma: no cover — shutdown best-effort
        logger.exception("Error closing vector store during shutdown")


@app.get("/health")
async def health():
    return {
        "status": "ok",
        "version": "1.1.0",
        "mode": "openshift" if settings.openshift else "local",
    }


@app.get("/ready")
async def readiness():
    """Readiness probe — checks that the LLM model config is valid."""
    errors = []
    if not settings.llm_model:
        errors.append("LLM_MODEL is not set")
    # LLM_API_KEY is required for every provider except Ollama (local).
    is_ollama = settings.llm_model.lower().startswith("ollama/") or (
        settings.llm_base_url and "11434" in settings.llm_base_url
    )
    if not is_ollama and not settings.llm_api_key:
        errors.append("LLM_API_KEY is not set")
    from app.services.capability_store import is_ready as store_ready
    if not store_ready():
        errors.append("Capability store not initialized")
    if errors:
        from fastapi.responses import JSONResponse
        return JSONResponse(status_code=503, content={"ready": False, "errors": errors})
    return {"ready": True, "model": settings.llm_model}


# ── Static file serving (production / OpenShift) ──────────────────────────────
# Mount AFTER the API routes so /api/* is always handled by FastAPI first.
# The static mount is a catch-all: only fires when no API route matched.
if settings.openshift or settings.serve_static:
    from fastapi.staticfiles import StaticFiles

    static_dir = os.path.abspath(settings.static_dir)
    if os.path.isdir(static_dir):
        app.mount("/", StaticFiles(directory=static_dir, html=True), name="static")
        print(f"[startup] Serving frontend static files from: {static_dir}")
    else:
        print(
            f"[startup] WARNING: static_dir '{static_dir}' not found — "
            "frontend will not be served. Run `npm run build` and copy dist/ to that path."
        )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "main:app",
        host=settings.host,
        port=settings.port,
        reload=settings.debug and not settings.openshift,
    )
