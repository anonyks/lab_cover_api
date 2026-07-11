# Thapathali Cover Generator

This project is a Python replica of a cover service of:
https://github.com/Jeeshan-Shrestha/bct-godfather

Generate lab cover PDFs from a DOCX template.

Main files:
- `cover_api.py` (CLI)
- `cover_api_server.py` (API + web UI)

Detailed CLI behavior, edge cases, and troubleshooting are in `info.md`.

## Quick Start

```bash
./cover_api.py --id="14"
```

## Requirements

- Python 3.9+
- LibreOffice

Install Python dependencies:

```bash
pip3 install -r requirements.txt
```

## API / UI

```bash
python3 cover_api_server.py
```

- `GET /health`
- `GET /` (web UI)
- `POST /generate`

Default URL: `http://127.0.0.1:8008`
