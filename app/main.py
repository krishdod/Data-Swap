from __future__ import annotations

import json
import os
import uuid
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .excel_utils import read_headers_xlsx, read_source_preview_xlsx, write_filled_template
from .mapping_suggest import suggest_sources_for_target


ROOT = Path(__file__).resolve().parent.parent
UPLOADS = ROOT / "uploads"
UPLOADS.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Sheet Updating")
templates = Jinja2Templates(directory=str(ROOT / "templates"))
app.mount("/static", StaticFiles(directory=str(ROOT / "static")), name="static")

MAX_SUGGESTION_COMPARISONS = 50_000
ASSET_VERSION = os.getenv("RENDER_GIT_COMMIT") or str(uuid.uuid4())


def _save_upload(f: UploadFile) -> Path:
    file_id = str(uuid.uuid4())
    safe_name = (f.filename or "upload.xlsx").replace("/", "_").replace("\\", "_")
    out = UPLOADS / f"{file_id}__{safe_name}"
    with out.open("wb") as w:
        while True:
            chunk = f.file.read(1024 * 1024)
            if not chunk:
                break
            w.write(chunk)
    return out


def _build_suggestions(source_headers: list[str], template_headers: list[str]) -> dict[str, list[dict[str, Any]]]:
    """
    Build top-3 suggestions unless header counts are too large.
    This keeps upload responsive for very wide files.
    """
    comparisons = len(source_headers) * len(template_headers)
    if comparisons > MAX_SUGGESTION_COMPARISONS:
        return {th: [] for th in template_headers}

    suggestions: dict[str, list[dict[str, Any]]] = {}
    for th in template_headers:
        suggestions[th] = [
            {"source": s.source_header, "score": round(s.score, 3)}
            for s in suggest_sources_for_target(th, source_headers, top_k=3)
        ]
    return suggestions


@app.get("/", response_class=HTMLResponse)
def index(request: Request) -> Any:
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/upload", response_class=HTMLResponse)
def upload(
    request: Request,
    source_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
) -> Any:
    source_path = _save_upload(source_file)
    template_path = _save_upload(template_file)

    source_preview = read_source_preview_xlsx(str(source_path))
    template_headers = read_headers_xlsx(str(template_path))

    suggestions = _build_suggestions(source_preview.headers, template_headers)

    payload = {
        "source": {
            "headers": source_preview.headers,
            "preview_rows": source_preview.rows,
            "row_count": source_preview.row_count,
        },
        "template": {"headers": template_headers},
        "suggestions": suggestions,
        "files": {"source_path": str(source_path), "template_path": str(template_path)},
    }

    return templates.TemplateResponse(
        "map.html",
        {
            "request": request,
            "payload_json": json.dumps(payload),
            "asset_version": ASSET_VERSION,
        },
    )


@app.post("/swap", response_class=HTMLResponse)
def swap(request: Request, payload: str = Form(...)) -> Any:
    """
    Swap source/template roles if user uploaded them incorrectly.
    Re-extract preview/headers and regenerate suggestions.
    """
    data = json.loads(payload)
    source_path = data["files"]["template_path"]
    template_path = data["files"]["source_path"]

    source_preview = read_source_preview_xlsx(str(source_path))
    template_headers = read_headers_xlsx(str(template_path))

    suggestions = _build_suggestions(source_preview.headers, template_headers)

    new_payload = {
        "source": {
            "headers": source_preview.headers,
            "preview_rows": source_preview.rows,
            "row_count": source_preview.row_count,
        },
        "template": {"headers": template_headers},
        "suggestions": suggestions,
        "files": {"source_path": str(source_path), "template_path": str(template_path)},
    }

    return templates.TemplateResponse(
        "map.html",
        {
            "request": request,
            "payload_json": json.dumps(new_payload),
            "asset_version": ASSET_VERSION,
        },
    )


@app.post("/export")
def export(
    payload: str = Form(...),
    mapping_json: str = Form(...),
) -> Response:
    data = json.loads(payload)
    mapping = json.loads(mapping_json)

    template_path = data["files"]["template_path"]
    source_path = data["files"]["source_path"]

    # Server-side validation gate (precision).
    template_headers = read_headers_xlsx(template_path)
    missing = [h for h in template_headers if h not in mapping]
    if missing:
        return Response(
            content=f"Missing mappings for: {missing}".encode("utf-8"),
            media_type="text/plain",
            status_code=400,
        )
    for h in template_headers:
        spec = mapping.get(h, {})
        if spec.get("type") not in {"source", "constant", "blank"}:
            return Response(
                content=f"Invalid mapping type for '{h}'".encode("utf-8"),
                media_type="text/plain",
                status_code=400,
            )

    out_bytes = write_filled_template(
        template_path=template_path,
        source_path=source_path,
        mapping=mapping,
        source_row_count_hint=int(data.get("source", {}).get("row_count") or 0),
    )

    filename = "filled_template.xlsx"
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Cache-Control": "no-store",
    }
    return Response(
        content=out_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )

