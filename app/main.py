from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI
from fastapi import Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles

from app.config import get_settings
from app.routes.report import router as report_router


settings = get_settings()
BASE_DIR = Path(__file__).resolve().parent
INDEX_HTML_PATH = BASE_DIR / "templates" / "index.html"
STATIC_DIR = BASE_DIR / "static"

app = FastAPI(
    title=settings.app_name,
    version=settings.app_version,
    description="Upload fragmented files, normalize them, and generate a report docx via OpenAI.",
)

app.include_router(report_router)
app.mount("/outputs", StaticFiles(directory=str(settings.output_dir)), name="outputs")
app.mount("/assets", StaticFiles(directory=str(STATIC_DIR)), name="assets")


@app.get("/", response_model=None)
def read_root(request: Request, format: str | None = None):
    if _should_render_html(request=request, format=format):
        return HTMLResponse(INDEX_HTML_PATH.read_text(encoding="utf-8"))

    return _build_status_payload()


@app.get("/api/status")
def read_status() -> dict[str, str]:
    return _build_status_payload()


def _build_status_payload() -> dict[str, str]:
    return {
        "status": "ok",
        "service": settings.app_name,
        "version": settings.app_version,
    }


def _should_render_html(request: Request, format: str | None) -> bool:
    if format == "json":
        return False
    if format == "html":
        return True

    accept = request.headers.get("accept", "")
    return "text/html" in accept
