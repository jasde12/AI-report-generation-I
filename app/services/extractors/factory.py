from __future__ import annotations

from pathlib import Path
from typing import Protocol

from app.models.schemas import ExtractedSource

from .docx_extractor import DocxExtractor
from .image_extractor import ImageExtractor
from .pdf_extractor import PdfExtractor
from .table_extractor import TableExtractor


class Extractor(Protocol):
    def extract(self, file_path: Path) -> ExtractedSource:
        ...


_EXTRACTOR_MAP: dict[str, Extractor] = {
    ".docx": DocxExtractor(),
    ".csv": TableExtractor(),
    ".xlsx": TableExtractor(),
    ".pdf": PdfExtractor(),
    ".jpg": ImageExtractor(),
    ".jpeg": ImageExtractor(),
    ".png": ImageExtractor(),
}


def get_extractor(file_path: Path) -> Extractor:
    extension = file_path.suffix.lower()
    extractor = _EXTRACTOR_MAP.get(extension)
    if extractor is None:
        supported = ", ".join(sorted(_EXTRACTOR_MAP))
        raise ValueError(f"不支援的副檔名: {extension or '(無)'}。支援格式: {supported}")
    return extractor


def extract_from_file(file_path: Path) -> ExtractedSource:
    extractor = get_extractor(file_path)
    return extractor.extract(file_path)


def supported_extensions() -> list[str]:
    return sorted(_EXTRACTOR_MAP)
