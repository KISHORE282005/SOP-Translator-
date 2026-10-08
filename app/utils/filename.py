"""
Filename generation utilities for the PO Translation System.

Handles safe filename generation and output file naming.
"""

from pathlib import Path

TRANSLATED_SUFFIX = "_Translated"


def generate_unique_filename(original_filename: str, file_id: str) -> str:
    """
    Generate a unique filename to avoid collisions.

    Args:
        original_filename: Original uploaded filename
        file_id: Unique ID for the file

    Returns:
        str: Unique filename
    """
    # Extract extension
    path = Path(get_safe_filename(original_filename))
    stem = path.stem
    suffix = path.suffix

    # Create unique name with file_id prefix
    unique_name = f"{file_id}_{stem}{suffix}"
    return unique_name


def strip_file_id(stored_filename: str, file_id: str) -> str:
    """
    Remove the file ID prefix added by generate_unique_filename.

    Args:
        stored_filename: Filename as stored on disk
        file_id: Unique ID for the file

    Returns:
        str: Filename without the file ID prefix
    """
    prefix = f"{file_id}_"
    if stored_filename.startswith(prefix):
        return stored_filename[len(prefix):]
    return stored_filename


def generate_output_filename(original_filename: str) -> str:
    """
    Generate output filename based on original filename.

    Args:
        original_filename: Original uploaded filename

    Returns:
        str: Output filename with _Translated suffix
    """
    path = Path(original_filename)
    stem = path.stem
    suffix = path.suffix

    # If the original filename already ends with _Translated, don't add again
    if stem.endswith(TRANSLATED_SUFFIX):
        return original_filename

    return f"{stem}{TRANSLATED_SUFFIX}{suffix}"


def get_safe_filename(filename: str) -> str:
    """
    Sanitize filename to prevent path traversal and other security issues.

    Args:
        filename: Filename to sanitize

    Returns:
        str: Safe filename
    """
    # Remove any path components (both separators, whatever the host OS is)
    safe_name = filename.replace("\\", "/").split("/")[-1]

    # Remove any potentially dangerous characters
    safe_name = safe_name.replace(" ", "_")
    safe_name = "".join(c for c in safe_name if c.isalnum() or c in "._-")

    # Ensure the filename is not empty
    if not Path(safe_name).stem.strip("."):
        safe_name = "unnamed_file.xlsx"

    return safe_name
