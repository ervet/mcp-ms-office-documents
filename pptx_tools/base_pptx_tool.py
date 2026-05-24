import logging
from typing import List, Dict, Any, Optional

from upload_tools import upload_file
from .slide_builder import PowerpointPresentation

logger = logging.getLogger(__name__)


def create_presentation(
    slides: List[Dict[str, Any]],
    format: str = "4:3",
    file_name: Optional[str] = None,
    author: Optional[str] = None,
    footer_text: Optional[str] = None,
    show_slide_numbers: bool = False,
) -> str:
    """Create a PowerPoint presentation from structured slides and upload it.

    :param slides: List of slide dicts with keys based on slide_type
    :param format: "4:3" or "16:9"
    :param file_name: Optional custom filename (without extension)
    :param author: Author name for document properties
    :param footer_text: Optional footer text displayed on all slides
    :param show_slide_numbers: Whether to show slide numbers
    :return: Upload status or URL text
    """
    if not slides:
        raise ValueError("No slides provided")

    logger.info(f"Starting create_presentation: slides={len(slides)}, format={format}")

    presentation = PowerpointPresentation(
        slides, format,
        author=author,
        footer_text=footer_text,
        show_slide_numbers=show_slide_numbers,
    )
    file_object = presentation.save()

    text = upload_file(file_object, "pptx", filename=file_name)
    file_object.close()

    logger.info("PowerPoint upload completed")
    return text

