"""
Pydantic schemas for the PO Translation System API.

Defines request and response models for API validation.
"""

from pydantic import BaseModel, ConfigDict, Field
from typing import Any, List, Optional


class UploadResponse(BaseModel):
    """Response model for file upload endpoint."""

    model_config = ConfigDict(from_attributes=True)

    file_id: str = Field(..., description="Unique identifier for the uploaded file")
    filename: str = Field(..., description="Original filename of the uploaded file")
    status: str = Field(default="uploaded", description="Upload status")


class TranslateResponse(BaseModel):
    """Response model for translation endpoint."""

    model_config = ConfigDict(from_attributes=True)

    status: str = Field(..., description="Translation status")
    languages: List[str] = Field(..., description="List of languages translated to")
    output_filename: str = Field(..., description="Name of the generated output file")
    translated_cells: int = Field(default=0, description="Number of cells that were translated")


class TranslationContent(BaseModel):
    """Model for storing translation content."""

    model_config = ConfigDict(from_attributes=True)

    english_texts: List[str] = Field(default_factory=list, description="List of English text to translate")
    cell_mapping: List[dict] = Field(default_factory=list, description="Mapping of cells to translated content")


class TranslationResult(BaseModel):
    """Model for storing translation results."""

    model_config = ConfigDict(from_attributes=True)

    original_text: str = Field(..., description="Original English text")
    tamil_translation: str = Field(..., description="Tamil translation")
    hindi_translation: str = Field(..., description="Hindi translation")
    translated: bool = Field(default=True, description="Whether translation was successful")
    error_message: Optional[str] = Field(None, description="Error message if translation failed")


class CellInfo(BaseModel):
    """Model for storing cell information in Excel."""

    model_config = ConfigDict(from_attributes=True)

    row: int = Field(..., description="Row number (1-based)")
    column: int = Field(..., description="Column number (1-based)")
    value: Any = Field(..., description="Cell value")
    cell_type: str = Field(..., description="Type of cell (text, number, date, formula, etc.)")
    language_code: Optional[str] = Field(None, description="Language code if applicable")
    preserve_original: bool = Field(default=False, description="Whether to preserve original value")
