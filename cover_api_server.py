#!/usr/bin/env python3
from functools import lru_cache
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, Response
from pydantic import BaseModel

from cover_api import generate_cover_pdf


app = FastAPI(title="Thapathali Cover API", version="1.0.0")

BASE_DIR = Path(__file__).resolve().parent
UI_HTML_PATH = BASE_DIR / "cover_ui.html"


@lru_cache(maxsize=1)
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
    color: bool = False
    nummeth: bool = False


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
        roll_no, pdf_bytes = generate_cover_pdf(
            student_id=payload.id,
            name=payload.name,
            lab_num=payload.labnum,
            title=payload.title,
            department=payload.depart,
            date=payload.date,
            npdate=payload.npdate,
            color=payload.color,
            nummeth=payload.nummeth,
        )
    except Exception as exc:
        # Keep API errors readable for frontend and quick manual testing.
        raise HTTPException(status_code=400, detail=str(exc))

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{roll_no}_cover.pdf"'},
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("cover_api_server:app", host="127.0.0.1", port=8008, reload=False)
