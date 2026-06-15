#!/usr/bin/env python3
import argparse
import csv
import datetime as dt
import html
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
import zipfile
from pathlib import Path

try:
    import nepali_datetime as nep_dt
except ImportError:
    nep_dt = None


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "bct_info" / "Cover_template.docx"
NUMMETH_TEMPLATE_PATH = BASE_DIR / "bct_info" / "nummeth_cover.docx"
CSV_PATH = BASE_DIR / "bct_info" / "BCT081_students_data.csv"

# Three accepted ID shapes: bare number (14), short (081BCT014), full (THA081BCT014)
ID_NUMERIC_RE = re.compile(r"\d{1,3}")
ID_SHORT_RE = re.compile(r"\d{3}BCT\d{3}")
ID_THA_RE = re.compile(r"THA\d{3}BCT\d{3}")


def usage():
    return (
        "Usage:\n"
        "  cover_api.py --id=\"ID\" [--labnum=LAB_NO] [--title=\"TITLE\"] [--name=\"NAME\"] [--depart=\"DEPARTMENT\"] [--date[=YYYY-MM-DD] | --npdate] [--nummeth]\n\n"
        "Nummeth mode:\n"
        "  --nummeth        uses bct_info/nummeth_cover.docx template.\n"
        "  In this mode, if --labnum is omitted, it defaults to ..........\n\n"
        "Date flags:\n"
        "  --date           uses today's English date (YYYY-MM-DD).\n"
        "  --npdate         uses today's Nepali date (YYYY-MM-DD).\n"
        "  --date=VALUE     uses provided date in YYYY-MM-DD only.\n"
        "  Only one of these may be used at a time; all are optional.\n\n"
        "Allowed ID formats:\n"
        "  1) 14              -> THA081BCT014\n"
        "  2) 081BCT014       -> 081BCT014\n"
        "  3) tha081bct014    -> THA081BCT014\n\n"
        "Examples:\n"
        "  cover_api.py --id=\"14\"\n"
        "  cover_api.py --labnum=1 --title=\"Digital Logic\" --id=\"14\" --date\n"
        "  cover_api.py --labnum=1 --title=\"Digital Logic\" --id=\"14\" --npdate\n"
        "  cover_api.py --labnum=1 --title=\"Digital Logic\" --id=\"14\" --date=2026-06-08\n"
        "  cover_api.py --nummeth --id=\"14\" --labnum=1\n"
        "  cover_api.py --labnum=2 --title=\"Microprocessor\" --id=\"081BCT014\"\n"
        "  cover_api.py --labnum=3 --title=\"Signals\" --id=\"tha081bct014\" --name=\"john doe\"\n"
        "  cover_api.py --id=\"14\" --depart=\"Department of Computer Engineering\""
    )


def normalize_student_id(raw):
    value = raw.strip().upper()
    if ID_NUMERIC_RE.fullmatch(value):
        return f"THA081BCT{int(value):03d}"
    if ID_SHORT_RE.fullmatch(value):
        return value
    if ID_THA_RE.fullmatch(value):
        return value
    raise ValueError(
        "Invalid --id. Allowed formats: 14, 081BCT014, THA081BCT014 (case-insensitive for THA form)."
    )


def normalize_name(raw):
    return " ".join(part.capitalize() for part in raw.split())


def normalize_english_date(value):
    clean = value.strip()
    if not clean:
        raise ValueError("--date value cannot be empty. Use YYYY-MM-DD.")
    try:
        dt.date.fromisoformat(clean)
    except ValueError as exc:
        raise ValueError("Invalid --date format. Use YYYY-MM-DD.") from exc
    return clean


def subtract_7_days(date_str, npdate=False):
    """Subtract 7 days from a YYYY-MM-DD string, using the right calendar."""
    import datetime as _dt
    if npdate:
        if nep_dt is None:
            raise ValueError("nepali-datetime required for BS lab date calculation.")
        try:
            y, m, d = (int(part) for part in date_str.split("-"))
            bs_date = nep_dt.date(y, m, d)
            bs_date = bs_date - _dt.timedelta(days=7)
            return f"{bs_date.year:04d}-{bs_date.month:02d}-{bs_date.day:02d}"
        except Exception:
            # Some inputs may be AD dates even when npdate is set.
            # Fall back to Gregorian arithmetic instead of failing hard.
            try:
                ad_date = _dt.date.fromisoformat(date_str) - _dt.timedelta(days=7)
                return ad_date.isoformat()
            except Exception:
                raise ValueError(f"Cannot subtract 7 days from date: {date_str!r}")

    try:
        ad_date = _dt.date.fromisoformat(date_str) - _dt.timedelta(days=7)
        return ad_date.isoformat()
    except Exception:
        raise ValueError(f"Cannot subtract 7 days from date: {date_str!r}")


def resolve_submission_date(date_value, npdate=False):
    if date_value is not None and npdate:
        raise ValueError("Use only one date mode: --date, --npdate, or --date=YYYY-MM-DD")

    if npdate:
        if nep_dt is None:
            raise ValueError(
                "--npdate needs package 'nepali-datetime'. Install with: pip3 install nepali-datetime"
            )
        bs = nep_dt.date.today()
        return f"{bs.year:04d}-{bs.month:02d}-{bs.day:02d}"

    if date_value is not None:
        if date_value == "today":
            return dt.date.today().isoformat()
        return normalize_english_date(date_value)

    return None


def split_report_title(title, width=32):
    wrapped = textwrap.wrap(title.strip(), width=width, break_long_words=False, break_on_hyphens=False)
    if not wrapped:
        return "", ""
    if len(wrapped) == 1:
        return wrapped[0], ""
    return wrapped[0], " ".join(wrapped[1:])


def split_department_text(department, width=22):
    wrapped = textwrap.wrap(department.strip(), width=width, break_long_words=False, break_on_hyphens=False)
    if not wrapped:
        return "", "", ""
    return (
        wrapped[0],
        wrapped[1] if len(wrapped) > 1 else "",
        " ".join(wrapped[2:]) if len(wrapped) > 2 else "",
    )


def _sanitize_filename_value(value, *, compact_words=False):
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if compact_words:
        text = re.sub(r"\s+", "", text)
        text = re.sub(r"[^A-Za-z0-9]", "", text)
        return text.lower()
    # Keep alnum, underscore, hyphen, and dot; strip the rest.
    text = re.sub(r"[^A-Za-z0-9._-]", "", text)
    return text


def build_output_filename(*, mode, lab_num, title, submission_date, student_id):
    mode_part = "nummeth" if mode == "nummeth" else "lab"

    cleaned_lab = _sanitize_filename_value(lab_num)
    if cleaned_lab and not re.fullmatch(r"[.]+", cleaned_lab):
        lab_part = f"lab{cleaned_lab}"
    else:
        lab_part = ""

    title_words = str(title or "").strip().split()
    title_tail = " ".join(title_words[-3:]) if title_words else ""
    title_part = _sanitize_filename_value(title_tail, compact_words=True)
    date_part = _sanitize_filename_value(submission_date)
    id_part = _sanitize_filename_value(student_id)

    parts = [p for p in [mode_part, lab_part, title_part, date_part, id_part, "cover"] if p]
    filename = "_".join(parts) + ".pdf"
    # Keep filenames manageable across filesystems.
    return filename[:120]


def update_document_xml(xml, *, name, roll_no, lab_num, report_title, department, submission_date, lab_date=None):
    xml = replace_split_placeholder(xml, "student_name", name, " ")
    # New template uses {{roll_num}}; keep {{roll_no}} fallback for compatibility.
    xml = replace_split_placeholder(xml, "roll_num", roll_no, ": ")
    xml = replace_split_placeholder(xml, "roll_no", roll_no, ": ")

    if lab_num is not None:
        xml = replace_split_placeholder(xml, "lab_num", lab_num, "")
        escaped_lab_num = html.escape(lab_num, quote=True)
        xml = re.sub(
            r"Lab Report No:\s*[\.…]+",
            f"Lab Report No: {escaped_lab_num}",
            xml,
            count=1,
        )

    if report_title is not None:
        title_line_1, title_line_2 = split_report_title(report_title)
        escaped_title_1 = html.escape(title_line_1, quote=True)
        escaped_title_2 = html.escape(title_line_2, quote=True)
        # Line 1: replace dots after label (include optional trailing lone dot)
        xml = re.sub(
            r"A Report on:\s*[\.…]+\.?",
            f"A Report on: {escaped_title_1}",
            xml,
            count=1,
        )
        # Line 2: replace dots-only run with title_line_2 while preserving
        # template whitespace around the dots (nummeth uses leading spaces here).
        xml = re.sub(
            r">(\s*)[\.…]{6,}(\s*)<",
            lambda m: f">{m.group(1)}{escaped_title_2}{m.group(2)}<",
            xml,
            count=1,
        )
        # Clear the lone '.' run that immediately follows the dots run in the same paragraph
        xml = re.sub(r"(<w:t[^>]*>)\.(</w:t></w:r></w:p>)", r"\1\2", xml, count=1)

    if department is not None:
        dep_line_1, dep_line_2, dep_line_3 = split_department_text(department)
        dep_line_1 = html.escape(dep_line_1, quote=True)
        dep_line_2 = html.escape(dep_line_2, quote=True)
        dep_line_3 = html.escape(dep_line_3, quote=True)

        xml = xml.replace(">Department of<", f">{dep_line_1}<", 1)
        xml = xml.replace(">Computer and<", f">{dep_line_2}<", 1)
        xml = xml.replace(">Electronics Engineering<", f">{dep_line_3}<", 1)

    # Lab Date (nummeth only) must be explicit: manual/auto value only.
    # If lab_date is omitted, template dots are kept untouched.
    # The template has "Lab Date: ……<tab>...………" all in one run; replace both dot segments.
    if lab_date:
        escaped_lab_date = html.escape(lab_date, quote=True)
        xml = re.sub(
            r"Lab Date:\s*[\.…]+",
            f"Lab Date: {escaped_lab_date}",
            xml,
            count=1,
        )

    if submission_date:
        escaped_submission_date = html.escape(submission_date, quote=True)
        # "Submission Date:" and "Signature" share a run in nummeth.
        # Keep template spacing and add 8 spaces before Signature for alignment.
        xml = re.sub(
            r"Submission Date:\s*[\.…]+(\s*Signature)?",
            lambda m: f"Submission Date: {escaped_submission_date}"
            + (re.sub(r"Signature", "       Signature", m.group(1), count=1) if m.group(1) else ""),
            xml,
            count=1,
        )
        # In updated templates, Signature may be in a separate run; indent it explicitly.
        xml = re.sub(r"<w:t>Signature</w:t>", r"<w:t xml:space=\"preserve\">      Signature</w:t>", xml, count=1)

    return xml


def load_students(csv_path):
    students = {}
    with csv_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            roll = (row.get("Roll_No") or "").strip().upper()
            first = (row.get("first_name") or "").strip()
            last = (row.get("last_name") or "").strip()
            if roll:
                students[roll] = f"{first} {last}".strip()
    return students


def replace_split_placeholder(xml_text, key, value, prefix):
    # Word splits "...{{key}}" across XML runs with proofErr spellcheck markers.
    # The {{ may appear at the END of a run that already contains label text
    # (e.g. "Name: {{"), or at the start of its own run (" {{").
    # Capture whatever precedes {{ in that run and preserve it.
    pattern = (
        r"(<w:t[^>]*>[^<]*?)\{\{</w:t></w:r>"
        + r"<w:proofErr w:type=\"spellStart\"/>"
        + r"<w:r[^>]*><w:rPr>.*?</w:rPr><w:t>"
        + re.escape(key)
        + r"</w:t></w:r>"
        + r"<w:proofErr w:type=\"spellEnd\"/>"
        + r"<w:r[^>]*><w:rPr>.*?</w:rPr><w:t>\}\}</w:t></w:r>"
    )
    escaped_value = html.escape(value, quote=True)
    result = re.sub(
        pattern,
        lambda m: m.group(1) + escaped_value + "</w:t></w:r>",
        xml_text,
        flags=re.DOTALL,
    )
    # Fallback for templates where Word didn't split the placeholder at all.
    return result.replace("{{" + key + "}}", escaped_value)


def fill_template(template_path, name, roll_no, lab_num, report_title, department, submission_date, lab_date=None):
    # A .docx is just a zip - we repack it in-memory with only document.xml patched.
    with zipfile.ZipFile(template_path, "r") as zin:
        # delete=False because on Windows you can't open a file twice while it's open.
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_path = Path(tmp.name)

        try:
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "word/document.xml":
                        xml = data.decode("utf-8")
                        xml = update_document_xml(
                            xml,
                            name=name,
                            roll_no=roll_no,
                            lab_num=lab_num,
                            report_title=report_title,
                            department=department,
                            submission_date=submission_date,
                            lab_date=lab_date,
                        )
                        data = xml.encode("utf-8")
                    zout.writestr(item, data)
            return tmp_path.read_bytes()
        finally:
            tmp_path.unlink(missing_ok=True)


def _libreoffice_command():
    for cmd in ("libreoffice", "soffice", "soffice.exe"):
        if shutil.which(cmd):
            return cmd

    # Windows fallback: LibreOffice is often installed outside PATH.
    windows_candidates = (
        Path("C:/Program Files/LibreOffice/program/soffice.exe"),
        Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
    )
    for candidate in windows_candidates:
        if candidate.exists():
            return str(candidate)

    return None


def convert_docx_to_pdf(docx_bytes, roll_no):
    cmd = _libreoffice_command()
    if not cmd:
        raise RuntimeError("LibreOffice not found. Install LibreOffice to generate PDF.")

    with tempfile.TemporaryDirectory(prefix=f"cover_{roll_no}_") as td:
        tmp_dir = Path(td)
        docx_file = tmp_dir / f"{roll_no}.docx"
        pdf_file = tmp_dir / f"{roll_no}.pdf"
        docx_file.write_bytes(docx_bytes)

        proc = subprocess.run(
            [cmd, "--headless", "--convert-to", "pdf", "--outdir", str(tmp_dir), str(docx_file)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            check=False,
        )
        if proc.returncode != 0 or not pdf_file.exists():
            raise RuntimeError(f"PDF conversion failed.\n{proc.stdout.strip()}")

        return pdf_file.read_bytes()


def parse_args(argv):
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--id", dest="student_id")
    parser.add_argument("--name", dest="name")
    parser.add_argument("--depart", dest="department")
    parser.add_argument("--labnum", dest="lab_num")
    parser.add_argument("--title", dest="title")
    date_group = parser.add_mutually_exclusive_group()
    date_group.add_argument("--date", nargs="?", const="today", dest="date")
    date_group.add_argument("--npdate", action="store_true", dest="npdate")
    parser.add_argument("--nummeth", action="store_true", dest="nummeth")
    parser.add_argument("--labdate", dest="lab_date")
    parser.add_argument("-h", "--help", action="store_true", dest="help")
    return parser.parse_known_args(argv[1:])


def generate_cover_pdf(
    *,
    student_id,
    name=None,
    lab_num=None,
    title=None,
    department=None,
    date=None,
    npdate=False,
    nummeth=False,
    lab_date=None,
    labdate_auto_src=None,  # 'today' | 'subdate' -> compute -7 days server-side
):
    selected_template = NUMMETH_TEMPLATE_PATH if nummeth else TEMPLATE_PATH

    if not selected_template.exists() or not CSV_PATH.exists():
        raise ValueError(
            "Missing template or CSV in bct_info folder. "
            f"Expected: {selected_template} and {CSV_PATH}"
        )

    if nummeth and lab_num is None:
        lab_num = "........"

    submission_date = resolve_submission_date(date, npdate=npdate)

    # Auto-compute lab_date server-side so BS arithmetic works correctly.
    if nummeth and lab_date is None and labdate_auto_src in ('today', 'subdate'):
        if labdate_auto_src == 'today':
            if npdate:
                if nep_dt is None:
                    raise ValueError("nepali-datetime required for Nepali lab date.")
                bs = nep_dt.date.today()
                base = f"{bs.year:04d}-{bs.month:02d}-{bs.day:02d}"
            else:
                import datetime as _dt
                base = _dt.date.today().isoformat()
        else:  # subdate
            if not submission_date:
                raise ValueError("Set a submission date before using 'subtract from submission date'.")
            base = submission_date
        lab_date = subtract_7_days(base, npdate=npdate)
    roll_no = normalize_student_id(student_id)

    students = load_students(CSV_PATH)
    if name is None:
        resolved_name = students.get(roll_no)
        if not resolved_name:
            raise ValueError(f"Student name not found for ID: {roll_no}. Use --name/ name field.")
    else:
        resolved_name = name

    resolved_name = normalize_name(resolved_name)

    if lab_num is not None:
        lab_num = lab_num.strip()
        if not lab_num:
            if nummeth:
                lab_num = ".........."
            else:
                raise ValueError("--labnum cannot be empty.")

    if title is not None:
        title = title.strip()
        if not title:
            raise ValueError("--title cannot be empty.")

    if department is not None:
        department = department.strip()
        if not department:
            raise ValueError("--depart cannot be empty.")

    docx_bytes = fill_template(
        selected_template,
        resolved_name,
        roll_no,
        lab_num,
        title,
        department,
        submission_date,
        lab_date=lab_date if nummeth else None,
    )
    pdf_bytes = convert_docx_to_pdf(docx_bytes, roll_no)
    return (
        roll_no,
        pdf_bytes,
        {
            "submission_date": submission_date,
            "lab_num": lab_num,
            "title": title,
            "mode": "nummeth" if nummeth else "lab",
        },
    )


def main(argv):
    args, extras = parse_args(argv)

    if args.help:
        print(usage())
        return 0

    if extras:
        print(f"Unknown arguments: {' '.join(extras)}")
        print(usage())
        return 1

    if not args.student_id:
        print(usage())
        return 1

    try:
        roll_no, pdf_bytes, meta = generate_cover_pdf(
            student_id=args.student_id,
            name=args.name,
            lab_num=args.lab_num,
            title=args.title,
            department=args.department,
            date=args.date,
            npdate=args.npdate,
            nummeth=args.nummeth,
            lab_date=args.lab_date,
        )
    except Exception as exc:
        print(f"Failed to generate cover: {exc}")
        return 1

    out_file = BASE_DIR / build_output_filename(
        mode=meta["mode"],
        lab_num=meta["lab_num"],
        title=meta["title"],
        submission_date=meta["submission_date"],
        student_id=roll_no,
    )
    out_file.write_bytes(pdf_bytes)
    print(f"Generated: {out_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
