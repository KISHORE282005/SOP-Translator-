"""
Upload API endpoint for PO Translation System.

Handles file upload, validation, and initial processing.
"""

import uuid
from pathlib import Path
from typing import Any

from fastapi import APIRouter, File, UploadFile, HTTPException, status

from app.models.schemas import UploadResponse
from app.utils.validation import validate_excel_file
from app.utils.logger import setup_logger
from app.utils.filename import generate_unique_filename
from app.config import settings

# Initialize router and logger
router = APIRouter()
logger = setup_logger(__name__)

# Upload directory
UPLOAD_DIR = Path(settings.upload_dir)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


@router.post("/upload", response_model=UploadResponse)
async def upload_file(file: UploadFile = File(...)) -> Any:
    """
    Upload an Excel PO file for translation.

    Args:
        file: The uploaded Excel file

    Returns:
        UploadResponse: Contains file_id, filename, and status

    Raises:
        HTTPException: If file validation fails
    """
    try:
        logger.info(f"Received file upload: {file.filename}")

        # Validate file extension
        extension = Path(file.filename or "").suffix.lower()
        if extension not in settings.allowed_extensions:
            logger.warning(f"Invalid file extension for {file.filename}")
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Only {', '.join(sorted(settings.allowed_extensions))} files are supported."
            )

        # Validate file size (read one byte past the limit, never the whole oversized file)
        content = await file.read(settings.max_file_size + 1)
        if len(content) > settings.max_file_size:
            logger.warning(f"File too large: {file.filename}")
            raise HTTPException(
                status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                detail=f"File size exceeds maximum allowed size of {settings.max_file_size} bytes."
            )

        # Validate Excel file content
        if not validate_excel_file(content):
            logger.warning(f"Invalid Excel file content: {file.filename}")
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Unable to read the Excel file. Please upload a valid PO workbook."
            )

        # Generate unique file ID and save file
        file_id = str(uuid.uuid4())
        unique_filename = generate_unique_filename(file.filename, file_id)
        file_path = UPLOAD_DIR / unique_filename

        # Save the uploaded file
        with open(file_path, "wb") as f:
            f.write(content)

        logger.info(f"File saved successfully: {unique_filename} (ID: {file_id})")

        return UploadResponse(
            file_id=file_id,
            filename=file.filename,
            status="uploaded"
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Unexpected error during file upload: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An unexpected error occurred during file upload."
        )
