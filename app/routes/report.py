from __future__ import annotations

import shutil
from pathlib import Path
from typing import Annotated, Any
from uuid import uuid4

from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile, status
from fastapi.responses import FileResponse

from app.config import get_settings
from app.models.schemas import (
    ExtractedSource,
    GenerateReportResponse,
    NormalizedDocument,
    ValidateQuestionnaireResponse,
)
from app.parsers.questionnaire_docx_parser import parse_questionnaire_docx
from app.schemas.questionnaire_schema import QuestionnaireParseResponse
from app.services.docx_renderer import DocxRenderer
from app.services.extractors.factory import extract_from_file, supported_extensions
from app.services.image_analyzer import ImageAnalyzer
from app.services.normalizer import Normalizer
from app.services.report_planner import ReportPlanner


router = APIRouter(tags=["report"])


@router.post("/generate-report", response_model=GenerateReportResponse)
async def generate_report(
    request: Request,
    project_name: Annotated[str, Form(...)],
    template_id: Annotated[str, Form(...)],
    files: Annotated[list[UploadFile], File(...)],
    hospital_name: Annotated[str | None, Form()] = None,
) -> GenerateReportResponse:
    normalized_document = await _build_normalized_document(
        project_name=project_name,
        hospital_name=hospital_name,
        template_id=template_id,
        files=files,
    )

    report_planner = ReportPlanner()
    docx_renderer = DocxRenderer()

    try:
        report_json = report_planner.generate_report(normalized_document)
        output_path = docx_renderer.render(
            project_name=project_name,
            template_id=template_id,
            report=report_json,
            normalized_document=normalized_document,
        )
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        ) from exc
    except RuntimeError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(exc),
        ) from exc

    normalizer = Normalizer()
    normalized_preview = normalizer.build_preview(normalized_document)
    base_url = str(request.base_url).rstrip("/")
    download_url = f"{base_url}/download/{output_path.name}"

    return GenerateReportResponse(
        success=True,
        normalized_preview=normalized_preview,
        report_json=report_json,
        output_file=f"outputs/{output_path.name}",
        download_url=download_url,
    )


@router.post("/validate-questionnaire", response_model=ValidateQuestionnaireResponse)
async def validate_questionnaire(
    project_name: Annotated[str, Form(...)],
    template_id: Annotated[str, Form(...)],
    files: Annotated[list[UploadFile], File(...)],
    hospital_name: Annotated[str | None, Form()] = None,
) -> ValidateQuestionnaireResponse:
    _ = (project_name, template_id, hospital_name)

    try:
        parsed = await _parse_first_docx_upload(files)
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        ) from exc
    except RuntimeError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(exc),
        ) from exc

    return ValidateQuestionnaireResponse(
        success=True,
        normalized_preview=parsed["raw_json"],
        understanding_json=parsed["report_input_json"],
    )


@router.post("/convert-questionnaire", response_model=QuestionnaireParseResponse)
async def convert_questionnaire(
    project_name: Annotated[str, Form(...)],
    template_id: Annotated[str, Form(...)],
    files: Annotated[list[UploadFile], File(...)],
    hospital_name: Annotated[str | None, Form()] = None,
) -> QuestionnaireParseResponse:
    _ = (project_name, template_id, hospital_name)

    try:
        parsed = await _parse_first_docx_upload(files)
        return QuestionnaireParseResponse.model_validate(parsed)
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(exc),
        ) from exc
    except RuntimeError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(exc),
        ) from exc


@router.get("/download/{file_name}")
def download_report(file_name: str) -> FileResponse:
    safe_name = Path(file_name).name
    if safe_name != file_name:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="不合法的檔名。",
        )

    output_path = get_settings().output_dir / safe_name
    if not output_path.is_file():
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="找不到報告檔案。",
        )

    return FileResponse(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=safe_name,
    )


async def _build_normalized_document(
    project_name: str,
    hospital_name: str | None,
    template_id: str,
    files: list[UploadFile],
) -> NormalizedDocument:
    if not files:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="至少要上傳一個檔案。",
        )

    extracted_sources = await _extract_sources_from_uploads(files)
    if not extracted_sources:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="沒有可處理的有效檔案。",
        )

    normalizer = Normalizer()
    return normalizer.normalize(
        project_name=project_name,
        hospital_name=hospital_name,
        template_id=template_id,
        sources=extracted_sources,
    )


async def _extract_sources_from_uploads(files: list[UploadFile]) -> list[ExtractedSource]:
    settings = get_settings()
    batch_dir = settings.upload_dir / uuid4().hex
    batch_dir.mkdir(parents=True, exist_ok=True)

    extracted_sources: list[ExtractedSource] = []
    image_analyzer = ImageAnalyzer()

    for upload in files:
        original_name = Path(upload.filename or "").name
        if not original_name:
            continue

        try:
            saved_path = await _save_upload_file(upload=upload, destination_dir=batch_dir)
            source = extract_from_file(saved_path)
            source.source_name = original_name
        except HTTPException:
            raise
        except ValueError as exc:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=str(exc),
            ) from exc
        finally:
            await upload.close()

        if source.type == "image":
            try:
                source.image_analysis = image_analyzer.analyze(source)
                source.warnings.extend(source.image_analysis.warnings)
            except RuntimeError as exc:
                raise HTTPException(
                    status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                    detail=f"圖片分析失敗: {source.source_name} - {exc}",
                ) from exc

        extracted_sources.append(source)

    return extracted_sources


async def _save_first_docx_upload(files: list[UploadFile]) -> tuple[Path, str]:
    if not files:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="至少要上傳一個檔案。",
        )

    settings = get_settings()
    batch_dir = settings.upload_dir / uuid4().hex
    batch_dir.mkdir(parents=True, exist_ok=True)

    chosen_path: Path | None = None
    chosen_name: str | None = None

    for upload in files:
        original_name = Path(upload.filename or "").name
        try:
            if chosen_path is None and Path(original_name).suffix.lower() == ".docx":
                chosen_path = await _save_upload_file(upload=upload, destination_dir=batch_dir)
                chosen_name = original_name
        finally:
            await upload.close()

    if chosen_path is None or chosen_name is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="convert-questionnaire 需要至少一個 .docx 問卷檔案。",
        )

    return chosen_path, chosen_name


async def _parse_first_docx_upload(files: list[UploadFile]) -> dict[str, Any]:
    saved_path, original_name = await _save_first_docx_upload(files)
    parsed = parse_questionnaire_docx(str(saved_path))
    parsed["raw_json"]["source_file"] = original_name
    parsed["normalized_json"]["source_file"]["file_name"] = original_name
    parsed["report_input_json"]["source_file"]["file_name"] = original_name
    return parsed


async def _save_upload_file(upload: UploadFile, destination_dir: Path) -> Path:
    safe_name = Path(upload.filename or "upload.bin").name
    suffix = Path(safe_name).suffix.lower()
    if suffix not in supported_extensions():
        supported = ", ".join(supported_extensions())
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"不支援的副檔名: {suffix or '(無)'}。支援格式: {supported}",
        )

    destination = destination_dir / f"{uuid4().hex}{suffix}"
    with destination.open("wb") as buffer:
        shutil.copyfileobj(upload.file, buffer)
    return destination
