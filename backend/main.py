"""
Excel AI Copilot – FastAPI Backend

Entry point for the backend server.

Local dev:   uvicorn main:app --reload --port 8000
Docker/OCP:  uvicorn main:app --host 0.0.0.0 --port 8080
             (or just run the container — CMD handles this)

Deployment modes (controlled via environment variables):
  OPENSHIFT=false  (default) — local dev, no static file serving
  OPENSHIFT=true             — production, serves built frontend from ./static
  SERVE_STATIC=true          — serve static files regardless of OPENSHIFT flag
"""

import os

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.config import settings
from app.routers import plan, stream, chat

app = FastAPI(
    title="Excel AI Copilot API",
    description="Backend for the Excel AI Copilot add-in. Provides LLM-powered plan generation and validation.",
    version="1.0.0",
)

# ── CORS ──────────────────────────────────────────────────────────────────────
# In OpenShift the frontend is served from the same origin as the API, so CORS
# is not strictly needed for same-origin requests. We still open it up for any
# external tooling (e.g. local dev hitting the deployed API, Office.js frame).
cors_origins = ["*"] if settings.openshift else settings.cors_origins

app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=not settings.openshift,  # credentials + wildcard origin is invalid
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── API routers ───────────────────────────────────────────────────────────────
app.include_router(chat.router)
app.include_router(plan.router)
app.include_router(stream.router)


@app.get("/health")
async def health():
    return {
        "status": "ok",
        "version": "1.0.0",
        "mode": "openshift" if settings.openshift else "local",
    }


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
