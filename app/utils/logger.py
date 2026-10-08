"""
Logger configuration for the PO Translation System.

Provides structured logging for all application events.
"""

import logging
import sys

import structlog
from structlog.stdlib import LoggerFactory

_configured = False


def _configure() -> None:
    """Configure logging and structlog once per process."""
    global _configured
    if _configured:
        return

    # Configure standard logging
    logging.basicConfig(
        format="%(message)s",
        stream=sys.stdout,
        level=logging.INFO,
    )

    # Configure structlog
    structlog.configure(
        processors=[
            structlog.stdlib.filter_by_level,
            structlog.stdlib.add_logger_name,
            structlog.stdlib.add_log_level,
            structlog.stdlib.PositionalArgumentsFormatter(),
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.processors.StackInfoRenderer(),
            structlog.processors.format_exc_info,
            structlog.processors.UnicodeDecoder(),
            # ensure_ascii=False keeps Tamil / Hindi text readable in the logs
            structlog.processors.JSONRenderer(ensure_ascii=False)
        ],
        context_class=dict,
        logger_factory=LoggerFactory(),
        wrapper_class=structlog.stdlib.BoundLogger,
        cache_logger_on_first_use=True,
    )
    _configured = True


def setup_logger(name: str) -> structlog.BoundLogger:
    """
    Set up structured logger for the given module.

    Args:
        name: Module name

    Returns:
        Configured structlog logger
    """
    _configure()
    logger = structlog.get_logger(name)
    return logger


def log_file_upload(filename: str, file_id: str, size: int) -> None:
    """Log file upload event."""
    logger = setup_logger("file_upload")
    logger.info("PO uploaded", filename=filename, file_id=file_id, size=size)


def log_workbook_opened(filename: str, sheet_count: int, sheet_names: list) -> None:
    """Log workbook opened event."""
    logger = setup_logger("workbook")
    logger.info("Workbook opened", filename=filename, sheet_count=sheet_count, sheet_names=sheet_names)


def log_translation_started(file_id: str, text_count: int) -> None:
    """Log translation start event."""
    logger = setup_logger("translation")
    logger.info("Translation started", file_id=file_id, text_count=text_count)


def log_tamil_translation_completed(file_id: str, count: int) -> None:
    """Log Tamil translation completion."""
    logger = setup_logger("translation")
    logger.info("Tamil translation completed", file_id=file_id, translated_count=count)


def log_hindi_translation_completed(file_id: str, count: int) -> None:
    """Log Hindi translation completion."""
    logger = setup_logger("translation")
    logger.info("Hindi translation completed", file_id=file_id, translated_count=count)


def log_translation_failed(file_id: str, error: str) -> None:
    """Log translation failure."""
    logger = setup_logger("translation")
    logger.error("Translation failed", file_id=file_id, error=error)


def log_output_generated(output_filename: str, size: int) -> None:
    """Log output file generation."""
    logger = setup_logger("output")
    logger.info("Output generated", filename=output_filename, size=size)


def log_download_requested(file_id: str) -> None:
    """Log download request."""
    logger = setup_logger("download")
    logger.info("Download requested", file_id=file_id)


def log_error(event: str, error: str, **kwargs) -> None:
    """Log generic error."""
    logger = setup_logger("error")
    logger.error(event, error=error, **kwargs)
