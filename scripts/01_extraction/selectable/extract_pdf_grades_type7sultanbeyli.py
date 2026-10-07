# -*- coding: utf-8 -*-
r"""
Extract the Sultanbeyli single-PDF grade list into Excel files.

This table type has one PDF containing many 3rd-grade branches. Each student row
already has separate columns for class/branch, school number, name, and grades.
The script creates one Excel file for each branch ("şube").

Install the required packages in the Python environment used by VS Code:

    pip install pdfplumber openpyxl

Recommended run:

    python extract_pdf_grades_sultanbeyli.py "C:\path\to\file.pdf"

You can also pass a folder path. In that case, every PDF in the folder is
processed.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path
from typing import Iterable

try:
    import pdfplumber
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    missing_name = getattr(exc, "name", "required package")
    raise SystemExit(
        f"Missing Python package: {missing_name}\n"
        "Install the requirements with:\n\n"
        "    pip install pdfplumber openpyxl\n"
    ) from exc


DEFAULT_INPUT_PATH = Path(r"C:\ERC-DATA\input.pdf")

COURSES = [
    "TÜRKÇE",
    "MATEMATİK",
    "HAYAT BİLGİSİ",
    "İNGİLİZCE",
    "GÖRSEL SANATLAR",
    "MÜZİK",
    "FEN BİLİMLERİ",
    "BEDEN EĞİTİMİ",
]

OUTPUT_COLUMNS = [
    "Sıra No",
    "İlçe",
    "Okul",
    "Dönem",
    "Sınıf",
    "Şube",
    "Okul No",
    "Ad Soyad",
    *COURSES,
]

GRADE_VALUES = [
    "ÇOK İYİ",
    "İYİ",
    "GELİŞTİRİLMELİ",
]

TABLE_SETTINGS_CANDIDATES = [
    {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "intersection_tolerance": 3,
        "edge_min_length": 3,
        "text_x_tolerance": 2,
        "text_y_tolerance": 3,
    },
    {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 5,
        "join_tolerance": 5,
        "intersection_tolerance": 5,
        "edge_min_length": 2,
        "text_x_tolerance": 2,
        "text_y_tolerance": 3,
    },
]

TURKISH_TO_ASCII = str.maketrans(
    {
        "ç": "c",
        "Ç": "C",
        "ğ": "g",
        "Ğ": "G",
        "ı": "i",
        "İ": "I",
        "ö": "o",
        "Ö": "O",
        "ş": "s",
        "Ş": "S",
        "ü": "u",
        "Ü": "U",
    }
)


def repair_mojibake(value: str) -> str:
    """Repair text that was decoded as Latin-1 even though it was UTF-8."""
    if not isinstance(value, str):
        return value

    suspicious = any(char in value for char in ("Ã", "Ä", "Å", "Â"))
    suspicious = suspicious or any("\x80" <= char <= "\x9f" for char in value)
    if not suspicious:
        return value

    try:
        return value.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return value


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", "\n").replace("\xa0", " ")
    text = repair_mojibake(text)
    text = unicodedata.normalize("NFC", text)
    return re.sub(r"\s+", " ", text).strip()


def fold_for_match(value: object) -> str:
    text = clean_text(value).translate(TURKISH_TO_ASCII)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.upper()
    return re.sub(r"[^A-Z0-9]+", " ", text).strip()


GRADE_ALIASES = {fold_for_match(grade): grade for grade in GRADE_VALUES}
GRADE_SCORES = {
    fold_for_match("GELİŞTİRİLMELİ"): 1,
    fold_for_match("İYİ"): 2,
    fold_for_match("ÇOK İYİ"): 3,
}
PARTIAL_GRADE_ALIASES = {
    "CO": 3,
    "COK": 3,
    "K": 3,
}
VALID_GRADE_OUTPUTS = set(GRADE_SCORES.values())


def row_text(row: Iterable[object]) -> str:
    return " ".join(clean_text(cell) for cell in row if clean_text(cell))


def normalize_int(value: object) -> str:
    text = clean_text(value)
    match = re.search(r"\d+", text)
    return match.group(0) if match else ""


def normalize_name(value: object) -> str:
    return clean_text(value).upper()


def normalize_grade(value: object) -> str:
    text = clean_text(value).upper()
    if not text:
        return ""

    folded = fold_for_match(text)
    if folded in GRADE_SCORES:
        return GRADE_SCORES[folded]
    if folded in PARTIAL_GRADE_ALIASES:
        return PARTIAL_GRADE_ALIASES[folded]

    for folded_grade, grade in GRADE_SCORES.items():
        if folded_grade in folded:
            return grade

    return text


def normalize_section_code(value: object) -> tuple[str, str]:
    """Return (class_no, section_letter) from values like '3-A' or '3-İ'."""
    text = clean_text(value).upper().replace(" ", "")
    match = re.search(r"(?P<class>\d+)\s*[-/]\s*(?P<section>[A-ZÇĞİÖŞÜIİ0-9]+)", text)
    if match:
        return match.group("class"), match.group("section")

    folded = fold_for_match(text)
    fallback = re.search(r"(?P<class>\d+)\s*[-/]\s*(?P<section>[A-Z0-9]+)", folded)
    if fallback:
        return fallback.group("class"), fallback.group("section")

    return "", ""


def normalize_row_to_12(row: list[object]) -> list[str]:
    """Keep the first three ID columns, one name column, and eight grade columns."""
    cells = [clean_text(cell) for cell in row]

    if len(cells) == 12:
        return cells

    if len(cells) > 12:
        first_three = cells[:3]
        grade_cells = cells[-len(COURSES) :]
        name = clean_text(" ".join(cells[3 : len(cells) - len(COURSES)]))
        return first_three + [name] + grade_cells

    return cells + [""] * (12 - len(cells))


def resolve_input_pdfs(raw_path: str | Path) -> tuple[list[Path], Path]:
    path = Path(raw_path).expanduser()

    if path.is_dir():
        pdf_files = sorted(item for item in path.glob("*.pdf") if item.is_file())
        if not pdf_files:
            raise FileNotFoundError(f"No PDF files were found in folder: {path}")
        return pdf_files, path

    if path.is_file():
        if path.suffix.lower() != ".pdf":
            raise ValueError(f"The input file is not a PDF: {path}")
        return [path], path.parent

    if path.suffix.lower() != ".pdf":
        pdf_version = path.with_suffix(".pdf")
        if pdf_version.is_file():
            return [pdf_version], pdf_version.parent

    raise FileNotFoundError(
        f"PDF file or folder was not found: {path}\n"
        "Tip: pass one PDF file path or a folder containing PDF files."
    )


def term_label_from_token(token: str) -> str:
    folded = fold_for_match(token)
    if folded in {"I", "1"}:
        return "1. Dönem"
    if folded in {"II", "2"}:
        return "2. Dönem"
    return clean_text(token)


def extract_metadata(pdf: pdfplumber.PDF) -> dict[str, str]:
    metadata = {
        "district": "",
        "school": "",
        "term": "",
        "class_no": "",
    }

    if not pdf.pages:
        return metadata

    first_text = pdf.pages[0].extract_text(x_tolerance=2, y_tolerance=3) or ""
    lines = [clean_text(line) for line in first_text.splitlines() if clean_text(line)]

    title = ""
    for line in lines[:10]:
        folded = fold_for_match(line)
        if "NOT LISTESI" in folded or ("SINIF" in folded and "DONEM" in folded):
            title = line
            break

    if not title and lines:
        title = lines[0]

    folded_title = fold_for_match(title)
    class_match_folded = re.search(r"\b(?P<class>\d+)\s+SINIF\b", folded_title)
    if class_match_folded:
        metadata["class_no"] = class_match_folded.group("class")

    raw_class_match = re.search(
        r"(?P<class>\d+)\.\s*S[ıiİI]n[ıiİI]f",
        title,
        flags=re.IGNORECASE,
    )

    school_prefix = title
    if raw_class_match:
        school_prefix = title[: raw_class_match.start()].strip()

    school_prefix = re.sub(r"\bT\.?\s*C\.?\b", "", school_prefix, flags=re.IGNORECASE)
    school_prefix = clean_text(school_prefix)
    prefix_parts = school_prefix.split()
    if prefix_parts:
        metadata["district"] = prefix_parts[0].upper()
        metadata["school"] = clean_text(" ".join(prefix_parts[1:])).upper()

    term_match = re.search(
        r"\b(?P<term>I{1,3}|IV|V|1|2)\.?\s*D[öoÖO]nem\b",
        title,
        flags=re.IGNORECASE,
    )
    if term_match:
        metadata["term"] = term_label_from_token(term_match.group("term"))
    elif "I DONEM" in folded_title or "1 DONEM" in folded_title:
        metadata["term"] = "1. Dönem"
    elif "II DONEM" in folded_title or "2 DONEM" in folded_title:
        metadata["term"] = "2. Dönem"

    return metadata


def extract_tables_from_page(page: pdfplumber.page.Page) -> list[list[list[object]]]:
    for settings in TABLE_SETTINGS_CANDIDATES:
        tables = page.extract_tables(table_settings=settings)
        usable = [
            table
            for table in tables
            if table and max((len(row) for row in table), default=0) >= 10
        ]
        if usable:
            return usable
    return []


def parse_student_row(
    raw_row: list[object],
    metadata: dict[str, str],
) -> tuple[dict[str, str] | None, list[str]]:
    warnings: list[str] = []
    cells = normalize_row_to_12(raw_row)

    serial_no = normalize_int(cells[0])
    class_no, section = normalize_section_code(cells[1])
    school_no = normalize_int(cells[2])
    student_name = normalize_name(cells[3])

    if not serial_no or not class_no or not section or not school_no or not student_name:
        return None, warnings

    student_row = {column: "" for column in OUTPUT_COLUMNS}
    student_row.update(
        {
            "Sıra No": serial_no,
            "İlçe": metadata.get("district", ""),
            "Okul": metadata.get("school", ""),
            "Dönem": metadata.get("term", ""),
            "Sınıf": class_no or metadata.get("class_no", ""),
            "Şube": section,
            "Okul No": school_no,
            "Ad Soyad": student_name,
        }
    )

    for course, grade_cell in zip(COURSES, cells[4 : 4 + len(COURSES)]):
        grade = normalize_grade(grade_cell)
        student_row[course] = grade
        if grade and grade not in VALID_GRADE_OUTPUTS:
            warnings.append(f"Unrecognized grade for {course}: {grade!r}")

    return student_row, warnings


def extract_pdf_rows(pdf_path: Path) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    rows: list[dict[str, str]] = []
    warnings: list[dict[str, str]] = []

    with pdfplumber.open(pdf_path) as pdf:
        metadata = extract_metadata(pdf)

        for page_index, page in enumerate(pdf.pages, start=1):
            tables = extract_tables_from_page(page)
            if not tables:
                warnings.append(
                    {
                        "pdf": pdf_path.name,
                        "page": str(page_index),
                        "row": "",
                        "reason": "No usable table detected on page.",
                        "raw_row": "",
                    }
                )
                continue

            for table in tables:
                for table_row_index, raw_row in enumerate(table, start=1):
                    parsed_row, row_warnings = parse_student_row(raw_row, metadata)

                    if parsed_row is None:
                        text = row_text(raw_row)
                        folded = fold_for_match(text)
                        first_cell = normalize_int(raw_row[0] if raw_row else "")
                        if first_cell and "OGRENCININ ADI SOYADI" not in folded:
                            warnings.append(
                                {
                                    "pdf": pdf_path.name,
                                    "page": str(page_index),
                                    "row": str(table_row_index),
                                    "reason": "Skipped row that looked partly numeric but not like a full student row.",
                                    "raw_row": repr([clean_text(cell) for cell in raw_row]),
                                }
                            )
                        continue

                    rows.append(parsed_row)
                    for warning in row_warnings:
                        warnings.append(
                            {
                                "pdf": pdf_path.name,
                                "page": str(page_index),
                                "row": str(table_row_index),
                                "reason": warning,
                                "raw_row": repr([clean_text(cell) for cell in raw_row]),
                            }
                        )

    return rows, warnings


def save_excel(rows: list[dict[str, str]], output_path: Path, sheet_title: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_title[:31] or "Grades"

    worksheet.append(OUTPUT_COLUMNS)
    for row in rows:
        worksheet.append([row.get(column, "") for column in OUTPUT_COLUMNS])

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    thin_gray = Side(style="thin", color="D9E2F3")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    text_columns = {"Sıra No", "Okul No"}
    for column_name in text_columns:
        column_index = OUTPUT_COLUMNS.index(column_name) + 1
        column_letter = get_column_letter(column_index)
        for row_number in range(2, worksheet.max_row + 1):
            worksheet[f"{column_letter}{row_number}"].number_format = "@"

    for column_index, column_cells in enumerate(worksheet.columns, start=1):
        header = clean_text(worksheet.cell(row=1, column=column_index).value)
        max_length = max(
            len(clean_text(cell.value))
            for cell in column_cells
            if clean_text(cell.value)
        )

        if header == "Ad Soyad":
            width = min(max(max_length + 2, 24), 42)
        elif header in COURSES:
            width = min(max(max_length + 2, 12), 22)
        elif header in {"İlçe", "Okul"}:
            width = min(max(max_length + 2, 14), 40)
        else:
            width = min(max(max_length + 2, 9), 18)

        worksheet.column_dimensions[get_column_letter(column_index)].width = width

    workbook.save(output_path)


def safe_file_part(value: str) -> str:
    value = clean_text(value)
    value = value.translate(TURKISH_TO_ASCII)
    value = unicodedata.normalize("NFKD", value)
    value = "".join(char for char in value if not unicodedata.combining(char))
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    return value.strip("._-") or "output"


def term_file_part(term: str) -> str:
    folded = fold_for_match(term)
    if folded.startswith("1"):
        return "1_donem"
    if folded.startswith("2"):
        return "2_donem"
    return safe_file_part(term) or "donem"


def group_rows_by_branch(
    rows: list[dict[str, str]],
) -> dict[tuple[str, str], list[dict[str, str]]]:
    grouped: dict[tuple[str, str], list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        grouped[(row.get("Sınıf", ""), row.get("Şube", ""))].append(row)
    return dict(grouped)


def output_path_for_group(
    pdf_path: Path,
    output_dir: Path,
    rows: list[dict[str, str]],
    class_no: str,
    section: str,
) -> Path:
    term = rows[0].get("Dönem", "") if rows else ""
    stem = safe_file_part(pdf_path.stem)
    class_part = safe_file_part(class_no or "sinif")
    section_part = safe_file_part(section or "sube")
    term_part = term_file_part(term)
    return output_dir / f"{stem}_{class_part}_{section_part}_{term_part}.xlsx"


def write_warning_log(warnings: list[dict[str, str]], output_dir: Path) -> Path | None:
    if not warnings:
        return None

    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / "sultanbeyli_warnings.csv"
    fieldnames = ["pdf", "page", "row", "reason", "raw_row"]

    with log_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(warnings)

    return log_path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert Sultanbeyli single-PDF grade lists into branch Excel files."
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        default=None,
        help="PDF file path or folder path. If omitted, the script asks for a path.",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Optional folder where Excel files and warnings will be saved.",
    )
    return parser.parse_args(argv)


def get_input_path(args: argparse.Namespace) -> str:
    if args.input_path:
        return args.input_path

    user_input = input(
        "Enter a PDF file path or a folder path containing PDFs "
        "[press Enter for the built-in example path]: "
    ).strip()
    return user_input or str(DEFAULT_INPUT_PATH)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        input_path = get_input_path(args)
        pdf_paths, input_base_dir = resolve_input_pdfs(input_path)
    except Exception as exc:
        print(f"Input error: {exc}", file=sys.stderr)
        return 1

    output_dir = (
        Path(args.output_dir).expanduser()
        if args.output_dir
        else input_base_dir / "sultanbeyli_output"
    )

    created_files: list[Path] = []
    all_warnings: list[dict[str, str]] = []
    converted_pdf_count = 0

    for pdf_index, pdf_path in enumerate(pdf_paths, start=1):
        try:
            rows, warnings = extract_pdf_rows(pdf_path)
            all_warnings.extend(warnings)

            if not rows:
                print(f"[{pdf_index}/{len(pdf_paths)}] No student rows found: {pdf_path.name}")
                continue

            grouped_rows = group_rows_by_branch(rows)
            for (class_no, section), branch_rows in sorted(
                grouped_rows.items(),
                key=lambda item: (normalize_int(item[0][0]), fold_for_match(item[0][1])),
            ):
                output_path = output_path_for_group(
                    pdf_path,
                    output_dir,
                    branch_rows,
                    class_no,
                    section,
                )
                sheet_title = f"{class_no}-{section}"
                save_excel(branch_rows, output_path, sheet_title)
                created_files.append(output_path)

            converted_pdf_count += 1
            print(
                f"[{pdf_index}/{len(pdf_paths)}] OK: {pdf_path.name} -> "
                f"{len(grouped_rows)} Excel file(s), {len(rows)} student row(s)"
            )

        except Exception as exc:
            all_warnings.append(
                {
                    "pdf": pdf_path.name,
                    "page": "",
                    "row": "",
                    "reason": f"PDF failed: {exc}",
                    "raw_row": "",
                }
            )
            print(f"[{pdf_index}/{len(pdf_paths)}] ERROR: {pdf_path.name}: {exc}", file=sys.stderr)

    warning_log = write_warning_log(all_warnings, output_dir)

    if not created_files:
        if warning_log:
            print(f"Warning log saved: {warning_log}")
        return 1

    print("\nCreated files:")
    for output_path in created_files:
        print(f"  {output_path}")
    if warning_log:
        print(f"\nWarning log saved: {warning_log}")

    return 0 if converted_pdf_count == len(pdf_paths) else 1


if __name__ == "__main__":
    raise SystemExit(main())
