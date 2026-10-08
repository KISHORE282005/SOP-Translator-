"""
Download API endpoint for PO Translation System.

Handles file download after translation is complete.
"""

from pathlib import Path
from typing import Any

from fastapi import APIRouter, HTTPException, status
from fastapi.responses import FileResponse

from app.utils.filename import TRANSLATED_SUFFIX, strip_file_id
from app.utils.logger import setup_logger
from app.utils.validation import is_valid_file_id
from app.config import settings

# Initialize router and logger
router = APIRouter()
logger = setup_logger(__name__)


@router.get("/download/{file_id}")
async def download_file(file_id: str) -> Any:
    """
    Download the translated PO file.

    Args:
        file_id: Unique ID of the uploaded file (used to find the output file)

    Returns:
        FileResponse: The translated Excel file

    Raises:
        HTTPException: If file not found
    """
    try:
        logger.info(f"Download requested for file ID: {file_id}")

        # Find the output file
        output_dir = Path(settings.output_dir)
        matching_files = (
            list(output_dir.glob(f"{file_id}_*{TRANSLATED_SUFFIX}.xlsx"))
            if is_valid_file_id(file_id) else []
        )

        if not matching_files:
            logger.error(f"No translated file found for ID: {file_id}")
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Translated file not found. Please complete translation first."
            )

        output_file = matching_files[0]
        download_name = strip_file_id(output_file.name, file_id)
        logger.info(f"Serving file: {download_name}")

        return FileResponse(
            path=output_file,
            filename=download_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Download failed: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An error occurred while downloading the file."
        )
