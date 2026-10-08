"""
Main entry point for the PO Translation System.

This module initializes the FastAPI application and includes all API routes.
"""

from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.api.upload import router as upload_router
from app.api.translate import router as translate_router
from app.api.download import router as download_router
from app.config import settings
from app.utils.logger import setup_logger

# Initialize logger
logger = setup_logger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Initialize application on startup."""
    logger.info("PO Translation System starting up")

    # Create necessary directories
    settings.upload_dir.mkdir(parents=True, exist_ok=True)
    settings.output_dir.mkdir(parents=True, exist_ok=True)
    yield
    logger.info("PO Translation System shutting down")


# Create FastAPI app
app = FastAPI(
    title=settings.app_name,
    description="A production-ready internal web application for translating Purchase Order Excel files",
    version="1.0.0",
    lifespan=lifespan
)

# Configure CORS for Streamlit frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include API routers
app.include_router(upload_router, prefix="/api", tags=["Upload"])
app.include_router(translate_router, prefix="/api", tags=["Translate"])
app.include_router(download_router, prefix="/api", tags=["Download"])


@app.get("/")
async def root():
    """Root endpoint - API information."""
    return {
        "application": settings.app_name,
        "version": "1.0.0",
        "description": "Translates Purchase Order Excel files to Tamil and Hindi",
        "endpoints": [
            "/api/upload - Upload Excel file",
            "/api/translate/{file_id} - Translate PO",
            "/api/download/{file_id} - Download translated file"
        ]
    }


@app.get("/health")
async def health():
    """Health check endpoint."""
    return {"status": "ok", "translation_provider": settings.translation_provider}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
