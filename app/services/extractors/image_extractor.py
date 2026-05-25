from __future__ import annotations

import base64
import mimetypes
from pathlib import Path

from app.models.schemas import ExtractedSource


class ImageExtractor:
    SUPPORTED_MIME_TYPES = {"image/jpeg", "image/png"}

    def extract(self, file_path: Path) -> ExtractedSource:
        mime_type, _ = mimetypes.guess_type(file_path.name)
        if mime_type not in self.SUPPORTED_MIME_TYPES:
            raise ValueError(f"不支援的圖片格式: {file_path.name}")

        try:
            encoded = base64.b64encode(file_path.read_bytes()).decode("utf-8")
        except Exception as exc:
            raise ValueError(f"無法讀取圖片檔案: {file_path.name}") from exc

        return ExtractedSource(
            type="image",
            source_name=file_path.name,
            mime_type=mime_type,
            data_url=f"data:{mime_type};base64,{encoded}",
            metadata={
                "size_bytes": file_path.stat().st_size,
                "path": str(file_path),
            },
        )
