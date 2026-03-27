"""
Excel AI Copilot – FastAPI Backend

Entry point for the backend server.
Run with: uvicorn main:app --reload --port 8000
"""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.config import settings
from app.routers import plan, stream, chat

app = FastAPI(
    title="Excel AI Copilot API",
    description="Backend for the Excel AI Copilot add-in. Provides LLM-powered plan generation and validation.",
    version="1.0.0",
)

# CORS for the Office Add-in (served from localhost during dev)
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Register routers
app.include_router(chat.router)
app.include_router(plan.router)
app.include_router(stream.router)


@app.get("/health")
async def health():
    return {"status": "ok", "version": "1.0.0"}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "main:app",
        host=settings.host,
        port=settings.port,
        reload=settings.debug,
    )
