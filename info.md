# cover_api.py - Practical Guide

This project is a Python replica of a cover service of:
https://github.com/Jeeshan-Shrestha/bct-godfather

This is a quick, human-readable reference for how `cover_api.py` behaves in real use.

## Common Commands

Minimal run (ID only):

```bash
./cover_api.py --id="14"
```

Standard run:

```bash
./cover_api.py --labnum=1 --title="Digital Logic" --id="14"
```

Numerical methods run:

```bash
./cover_api.py --nummeth --labnum=3 --title="Bisection Method" --id="14" --date
```

Numerical methods with Nepali submission date:

```bash
./cover_api.py --nummeth --labnum=3 --title="Bisection Method" --id="14" --npdate
```

Use name from CSV (no `--name`):

```bash
./cover_api.py --labnum=2 --title="Circuit Analysis" --id="THA081BCT014"
```

Manual name fallback (if CSV has no match):

```bash
./cover_api.py --labnum=3 --title="Network Lab" --id="081BCT014" --name="john doe"
```

Department override:

```bash
./cover_api.py --id="14" --depart="Department of Computer Engineering"
```

## ID Rules

Accepted formats:
- `14`
- `081BCT014`
- `THA081BCT014` (or lowercase `tha081bct014`)

Normalization:
- `14` -> `THA081BCT014`
- `081BCT014` -> unchanged
- `tha081bct014` -> `THA081BCT014`

If the format is wrong (for example `081BCT14`), it is rejected.

## Name, Title, Lab, Department

- `--name` is title-cased automatically (`john doe` -> `John Doe`).
- If `--name` is not provided, the script looks it up in CSV by normalized ID.
- `--labnum` and `--title` only affect the template if you provide them.
- Long titles are wrapped safely (word-aware).
- `--depart` is optional; when omitted, template department text stays as-is.

## Date Modes (Optional, Mutually Exclusive)

You can choose one date mode, or skip date entirely:

- `--date` -> today's English date (`YYYY-MM-DD`)
- `--npdate` -> today's Nepali date (`YYYY-MM-DD`)
- `--date=YYYY-MM-DD` -> fixed date value

Rules:
- Only one of these is allowed at a time.
- `--npdate=...` is not supported.
- Invalid date formats (like `08-06-2026`) are rejected.

## Numerical Methods Mode

- `--nummeth` uses the `bct_info/nummeth_cover.docx` template.
- In this mode, if `--labnum` is omitted, it defaults to `..........`.
- `--depart` is ignored in this mode.
- Lab Date is optional and is mainly used by the API/UI flow.
- When lab date auto-subtract is used from the UI/API, the server subtracts 7 days and handles BS dates correctly.

## Output

- Output file name format:
  - `lab_<labnum>_<last_3_title_words>_<submission_date>_<normalized_id>_cover.pdf`
  - `nummeth_<labnum>_<last_3_title_words>_<submission_date>_<normalized_id>_cover.pdf`
- Examples:
  - `lab_lab2_stackapplications_2026-06-15_THA081BCT014_cover.pdf`
  - `nummeth_lab3_bisectionmethod_2026-06-15_THA081BCT014_cover.pdf`

Filename notes:
- `labnum` becomes `lab2`, `lab3`, etc. - not `lab_2`.
- The title part uses at most the last 3 words.
- The title part is compacted to lowercase letters/numbers only.
- The ID part always uses the normalized roll, for example `THA081BCT014`.

## Requirements and Limits

Requirements:
- LibreOffice (for DOCX -> PDF)
- `nepali-datetime` package (only for `--npdate`)

Limits:
- Unsupported ID formats are rejected.
- If CSV has no name match, you must pass `--name`.
- `--npdate=YYYY-MM-DD` is not available.

## Quick Troubleshooting

- `Invalid --id` -> check ID format exactly.
- `Student name not found` -> add `--name="Your Name"`.
- `LibreOffice not found` -> install LibreOffice.
- `Invalid --date format` -> use `YYYY-MM-DD`.
- `--npdate needs package 'nepali-datetime'` -> run `pip3 install nepali-datetime`.

## API/UI

`cover_api_server.py` provides:
- `GET /health`
- `GET /` (web UI)
- `POST /generate`

The web UI supports:
- Lab Report mode and Numerical Methods mode
- English or Nepali submission dates
- Numerical Methods lab-date auto calculation or manual entry
