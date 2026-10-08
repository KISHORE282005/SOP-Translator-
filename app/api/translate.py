"""
Translate API endpoint for PO Translation System.

Handles the translation process from English to Tamil and Hindi.
"""

import shutil
import tempfile
from pathlib import Path
from typing import Any

from fastapi import APIRouter, HTTPException, status

from app.services.excel_service import ExcelService
from app.services.translation_service import TranslationError, TranslationService
from app.models.schemas import TranslateResponse
from app.utils.filename import generate_output_filename, strip_file_id
from app.utils.logger import setup_logger
from app.utils.validation import is_valid_file_id
from app.config import settings

# Initialize router and logger
router = APIRouter()
logger = setup_logger(__name__)

# Initialize services
excel_service = ExcelService()
translation_service = TranslationService()


# Plain "def" on purpose: Excel processing and the translation calls are blocking,
# so FastAPI runs this endpoint in a worker thread instead of the event loop.
@router.post("/translate/{file_id}", response_model=TranslateResponse)
def translate_file(file_id: str) -> Any:
    """
    Translate the uploaded PO file to Tamil and Hindi.

    Args:
        file_id: Unique ID of the uploaded file

    Returns:
        TranslateResponse: Translation status and output filename

    Raises:
        HTTPException: If translation process fails
    """
    try:
        logger.info(f"Starting translation for file ID: {file_id}")

        # Find the uploaded file
        upload_dir = Path(settings.upload_dir)
        matching_files = (
            list(upload_dir.glob(f"{file_id}_*.xlsx")) if is_valid_file_id(file_id) else []
        )

        if not matching_files:
            logger.error(f"No file found for ID: {file_id}")
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found. Please upload the file again."
            )

        input_file = matching_files[0]
        logger.info(f"Processing file: {input_file.name}")

        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir) / input_file.name

            # Copy file to temporary location (original file protection)
            shutil.copy2(input_file, temp_path)
            logger.info(f"Created working copy: {temp_path}")

            # Read and process the Excel file
            logger.info("Reading Excel workbook")
            workbook = excel_service.read_workbook(temp_path)

            # Identify translatable content
            logger.info("Identifying translatable content")
            translatable_content = excel_service.identify_translatable_content(workbook)

            if not translatable_content['english_texts']:
                logger.warning("No translatable content found")
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="The uploaded workbook does not contain usable data for translation."
                )

            logger.info(f"Found {len(translatable_content['english_texts'])} English text segments to translate")

            # Generate translations
            logger.info("Generating Tamil translation")
            tamil_translations = translation_service.translate_to_tamil(translatable_content['english_texts'])

            logger.info("Generating Hindi translation")
            hindi_translations = translation_service.translate_to_hindi(translatable_content['english_texts'])

            # Create output workbook. It is stored with the file ID prefix so the
            # download endpoint can find it; the user sees the name without the ID.
            output_filename = generate_output_filename(strip_file_id(input_file.name, file_id))
            output_path = Path(settings.output_dir) / f"{file_id}_{output_filename}"
            temp_output_path = Path(temp_dir) / "output.xlsx"

            logger.info(f"Creating output workbook: {output_filename}")
            excel_service.generate_translated_workbook(
                workbook,
                temp_path,
                tamil_translations,
                hindi_translations,
                translatable_content['cell_mapping'],
                temp_output_path
            )
            workbook.close()

            # Store the output file for download (only complete files reach the output folder)
            shutil.move(str(temp_output_path), str(output_path))

            logger.info(f"Translation completed successfully. Output: {output_filename}")

            return TranslateResponse(
                status="completed",
                languages=["Tamil", "Hindi"],
                output_filename=output_filename,
                translated_cells=len(translatable_content['cell_mapping'])
            )

    except HTTPException:
        raise
    except TranslationError as e:
        logger.error(f"Translation provider failed: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Translation service is temporarily unavailable. Please try again."
        )
    except Exception as e:
        logger.error(f"Translation failed: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An error occurred while translating the file."
        )
