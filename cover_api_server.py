#!/usr/bin/env python3
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, Response
from pydantic import BaseModel

from cover_api import build_output_filename, generate_cover_pdf


app = FastAPI(title="Thapathali Cover API", version="1.0.0")

BASE_DIR = Path(__file__).resolve().parent
UI_HTML_PATH = BASE_DIR / "cover_ui.html"


def load_ui_html():
    if not UI_HTML_PATH.exists():
        raise RuntimeError(f"UI file not found: {UI_HTML_PATH}")
    return UI_HTML_PATH.read_text(encoding="utf-8")


class CoverRequest(BaseModel):
    id: str
    name: Optional[str] = None
    labnum: Optional[str] = None
    title: Optional[str] = None
    depart: Optional[str] = None
    date: Optional[str] = None
    npdate: bool = False
    nummeth: bool = False
    labdate: Optional[str] = None          # explicit manual date
    labdate_auto_src: Optional[str] = None  # 'today' | 'subdate' -> server computes -7 days


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/", response_class=HTMLResponse)
def ui():
    try:
        return load_ui_html()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to load UI HTML: {exc}")


@app.post("/generate")
def generate_cover(payload: CoverRequest):
    try:
        roll_no, pdf_bytes, meta = generate_cover_pdf(
            student_id=payload.id,
            name=payload.name,
            lab_num=payload.labnum,
            title=payload.title,
            department=payload.depart,
            date=payload.date,
            npdate=payload.npdate,
            nummeth=payload.nummeth,
            lab_date=payload.labdate,
            labdate_auto_src=payload.labdate_auto_src,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Unexpected error: {exc}")

    download_name = build_output_filename(
        mode=meta["mode"],
        lab_num=meta["lab_num"],
        title=meta["title"],
        submission_date=meta["submission_date"],
        student_id=roll_no,
    )

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{download_name}"'},
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("cover_api_server:app", host="127.0.0.1", port=8008, reload=False)
