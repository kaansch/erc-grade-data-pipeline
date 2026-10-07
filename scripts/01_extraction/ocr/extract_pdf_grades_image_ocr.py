# -*- coding: utf-8 -*-
r"""
OCR-only image version for grade-table PDFs.

Use this script when the PDF text layer is missing, broken, or unreliable.
The script does not use pdfplumber. It renders every PDF page as a high-
resolution image, detects the table grid, OCRs the required cells, and writes
Excel output with the same columns as the earlier selectable-text versions.

Install Python packages first:

    pip install pymupdf pillow pytesseract opencv-python numpy openpyxl

You must also install the Tesseract OCR program itself. On Windows, install
Tesseract and Turkish language data if possible. If Tesseract is installed in
a non-standard place, pass its path with --tesseract-cmd.

Run on a folder:

    python extract_pdf_grades_image_ocr.py "C:\path\to\folder_with_pdfs"

Run on one PDF:

    python extract_pdf_grades_image_ocr.py "C:\path\to\file.pdf"

Useful options:

    --dpi 400
    --ocr-lang tur+eng
    --tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe"
    --debug-images
"""

from __future__ import annotations

import argparse
import csv
import difflib
import io
import re
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import NamedTuple

try:
    import cv2
    import fitz  # PyMuPDF
    import numpy as np
    import pytesseract
    from PIL import Image, ImageEnhance, ImageOps
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
except ImportError as exc:
    missing_name = getattr(exc, "name", "required package")
    raise SystemExit(
        f"Missing Python package: {missing_name}\n"
        "Install the requirements with:\n\n"
        "    pip install pymupdf pillow pytesseract opencv-python numpy openpyxl\n"
    ) from exc


DEFAULT_INPUT_PATH = Path(r"C:\ERC-DATA\input.pdf")

COURSES = [
    "TÜRKÇE",
    "MATEMATİK",
    "HAYAT BİLGİSİ",
    "YABANCI DİL",
    "GÖRSEL SANATLAR",
    "MÜZİK",
    "FEN BİLİMLERİ",
    "BEDEN EĞİTİMİ VE OYUN",
]

OUTPUT_COLUMNS = [
    "İlçe",
    "Okul",
    "Sınıf",
    "Şube",
    "Okul No",
    "Ad Soyad",
    *COURSES,
]

GRADE_VALUES = [
    "GELİŞTİRİLMELİ",
    "ÇOK İYİ",
    "İYİ",
    "YETERLİ",
    "ORTA",
    "ZAYIF",
    "PEK İYİ",
    "PEKİYİ",
]

TURKISH_TO_ASCII = str.maketrans(
    {
        "ç": "c",
        "Ç": "C",
        "ğ": "g",
        "Ğ": "G",
        "ı": "i",
        "I": "I",
        "İ": "I",
        "ö": "o",
        "Ö": "O",
        "ş": "s",
        "Ş": "S",
        "ü": "u",
        "Ü": "U",
    }
)

ASCII_TO_COURSE = {
    "TURKCE": "TÜRKÇE",
    "MATEMATIK": "MATEMATİK",
    "HAYAT BILGISI": "HAYAT BİLGİSİ",
    "YABANCI DIL": "YABANCI DİL",
    "GORSEL SANATLAR": "GÖRSEL SANATLAR",
    "MUZIK": "MÜZİK",
    "FEN BILIMLERI": "FEN BİLİMLERİ",
    "BEDEN EGITIMI VE OYUN": "BEDEN EĞİTİMİ VE OYUN",
}

ASCII_TO_GRADE = {
    "GELISTIRILMELI": "GELİŞTİRİLMELİ",
    "COK IYI": "ÇOK İYİ",
    "IYI": "İYİ",
    "YETERLI": "YETERLİ",
    "ORTA": "ORTA",
    "ZAYIF": "ZAYIF",
    "PEK IYI": "PEK İYİ",
    "PEKIYI": "PEKİYİ",
}


class GroupKey(NamedTuple):
    district: str
    school: str
    class_no: str
    section: str


@dataclass
class PageResult:
    pdf_path: Path
    page_number: int
    header: dict[str, str]
    course: str
    rows: list[dict[str, str]]


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def fold_for_match(value: object) -> str:
    text = clean_text(value).translate(TURKISH_TO_ASCII)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    return re.sub(r"[^A-Z0-9]+", " ", text).strip()


def fold_compact(value: object) -> str:
    return fold_for_match(value).replace(" ", "")


def normalize_name(value: str) -> str:
    text = clean_text(value)
    text = text.replace("|", " ").replace("_", " ")
    text = re.sub(r"[^0-9A-Za-zÇĞİÖŞÜçğıöşüÂâÎîÛû\s.'-]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text.upper()


def normalize_school_no(value: str) -> str:
    return "".join(re.findall(r"\d+", value))


def normalize_grade(value: str) -> str:
    folded = fold_for_match(value)
    compact = folded.replace(" ", "")

    if "GELIST" in compact:
        return "GELİŞTİRİLMELİ"
    if "YETER" in compact:
        return "YETERLİ"
    if "ZAY" in compact:
        return "ZAYIF"
    if "ORTA" in compact:
        return "ORTA"
    if "COK" in compact and "IYI" in compact:
        return "ÇOK İYİ"
    if ("PEK" in compact and "IYI" in compact) or "PEKIYI" in compact:
        return "PEK İYİ"
    if "IYI" in compact or compact in {"IY", "1YI", "IY1"}:
        return "İYİ"

    known = list(ASCII_TO_GRADE)
    match = difflib.get_close_matches(folded, known, n=1, cutoff=0.68)
    if match:
        return ASCII_TO_GRADE[match[0]]

    compact_known = {key.replace(" ", ""): value for key, value in ASCII_TO_GRADE.items()}
    match = difflib.get_close_matches(compact, list(compact_known), n=1, cutoff=0.68)
    if match:
        return compact_known[match[0]]

    return clean_text(value).upper()


def resolve_input_pdfs(raw_path: str | Path) -> list[Path]:
    path = Path(raw_path).expanduser()

    if path.is_dir():
        pdf_files = sorted(item for item in path.glob("*.pdf") if item.is_file())
        if not pdf_files:
            raise FileNotFoundError(f"No PDF files were found in folder: {path}")
        return pdf_files

    if path.is_file():
        if path.suffix.lower() != ".pdf":
            raise ValueError(f"The input file is not a PDF: {path}")
        return [path]

    if path.suffix.lower() != ".pdf":
        pdf_version = path.with_suffix(".pdf")
        if pdf_version.is_file():
            return [pdf_version]

    raise FileNotFoundError(
        f"PDF file or folder was not found: {path}\n"
        "Tip: pass a folder containing PDF files, or pass one PDF file."
    )


def default_output_dir(input_arg: str | Path, pdf_paths: list[Path]) -> Path:
    input_path = Path(input_arg).expanduser()
    if input_path.is_dir():
        return input_path
    if len(pdf_paths) == 1:
        return pdf_paths[0].parent
    return DEFAULT_INPUT_PATH


def configure_tesseract(tesseract_cmd: str | None) -> None:
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
        return

    common_windows_path = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    if common_windows_path.exists():
        pytesseract.pytesseract.tesseract_cmd = str(common_windows_path)


def verify_tesseract_available() -> None:
    try:
        pytesseract.get_tesseract_version()
    except Exception as error:
        raise RuntimeError(
            "Tesseract OCR is not available. Install Tesseract OCR, then rerun. "
            "If it is already installed, pass --tesseract-cmd with the full "
            "path to tesseract.exe."
        ) from error


def render_pdf_page(page, dpi: int) -> Image.Image:
    zoom = dpi / 72
    pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")


def preprocess_for_ocr(image: Image.Image) -> Image.Image:
    grayscale = ImageOps.grayscale(image)
    grayscale = ImageEnhance.Contrast(grayscale).enhance(1.8)
    array = np.array(grayscale)
    threshold = cv2.adaptiveThreshold(
        array,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        35,
        11,
    )
    return Image.fromarray(threshold)


def pil_to_gray_array(image: Image.Image) -> np.ndarray:
    if image.mode != "L":
        image = ImageOps.grayscale(image)
    return np.array(image)


def ocr_image(
    image: Image.Image,
    lang: str,
    psm: int,
    whitelist: str | None = None,
    preserve_spaces: bool = False,
) -> str:
    config_parts = [f"--psm {psm}", "--oem 3"]
    if whitelist:
        config_parts.append(f"-c tessedit_char_whitelist={whitelist}")
    if preserve_spaces:
        config_parts.append("-c preserve_interword_spaces=1")

    config = " ".join(config_parts)
    try:
        return pytesseract.image_to_string(image, lang=lang, config=config)
    except pytesseract.TesseractError as error:
        if "Failed loading language" in str(error) and lang != "eng":
            return pytesseract.image_to_string(image, lang="eng", config=config)
        raise


def extract_header_info(header_text: str) -> dict[str, str]:
    district = ""
    school = ""
    class_no = "3"
    section = ""

    lines = [clean_text(line) for line in header_text.splitlines() if clean_text(line)]

    for line in lines:
        folded = fold_for_match(line)
        if "/" not in line:
            continue
        if "MUDURLUGU" not in folded and "MUDURLUG" not in folded:
            continue

        district_part, school_part = line.split("/", 1)
        district = clean_text(district_part)
        school_part = clean_text(school_part)
        school = re.split(
            r"\s+M[üu]d[üu]rl[üu][ğg][üu]\b",
            school_part,
            maxsplit=1,
            flags=re.IGNORECASE,
        )[0].strip()
        if school == school_part:
            folded_school = fold_for_match(school_part)
            school = re.sub(r"\s*MUDURLUGU?.*$", "", folded_school).title()
        break

    folded_header = fold_for_match(header_text)
    section_match = re.search(
        r"\b(?P<class>\d+)\s*SINIF\s*(?:/)?\s*(?P<section>[A-Z0-9]+)\s*SUBESI\b",
        folded_header,
    )
    if section_match:
        class_no = clean_text(section_match.group("class")) or "3"
        section = clean_text(section_match.group("section"))

    return {
        "district": district,
        "school": school,
        "class_no": class_no,
        "section": section,
    }


def extract_course_name(header_text: str, pdf_name: str = "") -> str:
    folded_header = fold_for_match(header_text)

    title_match = re.search(r"\bDERS YILI (?P<course>.+?) DERSI\b", folded_header)
    if title_match:
        candidate = clean_text(title_match.group("course"))
        for ascii_course, course in ASCII_TO_COURSE.items():
            if candidate == ascii_course:
                return course
        for ascii_course, course in sorted(
            ASCII_TO_COURSE.items(),
            key=lambda item: len(item[0]),
            reverse=True,
        ):
            if ascii_course in candidate or candidate in ascii_course:
                return course

    for ascii_course, course in sorted(
        ASCII_TO_COURSE.items(),
        key=lambda item: len(item[0]),
        reverse=True,
    ):
        if ascii_course in folded_header:
            return course

    folded_file_name = fold_for_match(pdf_name)
    for ascii_course, course in sorted(
        ASCII_TO_COURSE.items(),
        key=lambda item: len(item[0]),
        reverse=True,
    ):
        if ascii_course in folded_file_name:
            return course

    return ""


def merge_positions(indices: np.ndarray, max_gap: int = 3) -> list[int]:
    if len(indices) == 0:
        return []

    groups: list[list[int]] = []
    current = [int(indices[0])]
    for index in map(int, indices[1:]):
        if index - current[-1] <= max_gap:
            current.append(index)
        else:
            groups.append(current)
            current = [index]
    groups.append(current)

    return [int(round(sum(group) / len(group))) for group in groups]


def detect_table_lines(image: Image.Image) -> tuple[list[int], list[int]]:
    gray = pil_to_gray_array(image)
    binary_inverse = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    height, width = binary_inverse.shape
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(width // 25, 40), 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(height // 20, 40)))

    horizontal = cv2.morphologyEx(binary_inverse, cv2.MORPH_OPEN, horizontal_kernel)
    vertical = cv2.morphologyEx(binary_inverse, cv2.MORPH_OPEN, vertical_kernel)

    horizontal_projection = np.count_nonzero(horizontal, axis=1)
    vertical_projection = np.count_nonzero(vertical, axis=0)

    y_indices = np.where(horizontal_projection > width * 0.35)[0]
    x_indices = np.where(vertical_projection > height * 0.18)[0]

    horizontal_lines = merge_positions(y_indices, max_gap=6)
    vertical_lines = merge_positions(x_indices, max_gap=6)

    horizontal_lines = [y for y in horizontal_lines if 0 <= y < height]
    vertical_lines = [x for x in vertical_lines if 0 <= x < width]
    return horizontal_lines, vertical_lines


def choose_table_grid(
    horizontal_lines: list[int],
    vertical_lines: list[int],
    image_height: int,
) -> tuple[list[int], list[int]]:
    if len(horizontal_lines) < 4:
        raise ValueError("table horizontal lines could not be detected")
    if len(vertical_lines) < 6:
        raise ValueError("table vertical lines could not be detected")

    horizontal_lines = sorted(horizontal_lines)
    vertical_lines = sorted(vertical_lines)

    # The table starts below the centered header. Keep the longest lower cluster.
    filtered_horizontal = [y for y in horizontal_lines if y > image_height * 0.10]
    if len(filtered_horizontal) >= 4:
        horizontal_lines = filtered_horizontal

    vertical_lines = vertical_lines[:6]
    return horizontal_lines, vertical_lines


def crop_cell(
    image: Image.Image,
    left: int,
    top: int,
    right: int,
    bottom: int,
    pad_x: int = 8,
    pad_y: int = 3,
) -> Image.Image:
    width, height = image.size
    left = max(0, left + pad_x)
    top = max(0, top + pad_y)
    right = min(width, right - pad_x)
    bottom = min(height, bottom - pad_y)
    if right <= left or bottom <= top:
        return image.crop((0, 0, 1, 1))
    return image.crop((left, top, right, bottom))


def enhance_cell_for_ocr(cell: Image.Image, scale: float = 2.0) -> Image.Image:
    grayscale = ImageOps.grayscale(cell)
    grayscale = ImageEnhance.Contrast(grayscale).enhance(2.0)
    if scale != 1:
        new_size = (max(1, int(grayscale.width * scale)), max(1, int(grayscale.height * scale)))
        grayscale = grayscale.resize(new_size, Image.Resampling.LANCZOS)
    array = np.array(grayscale)
    threshold = cv2.threshold(array, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    return Image.fromarray(threshold)


def extract_rows_from_table(
    page_image: Image.Image,
    processed_image: Image.Image,
    lang: str,
    pdf_path: Path,
    page_number: int,
    errors: list[dict[str, str]],
) -> list[dict[str, str]]:
    horizontal_lines, vertical_lines = detect_table_lines(processed_image)
    horizontal_lines, vertical_lines = choose_table_grid(
        horizontal_lines,
        vertical_lines,
        processed_image.height,
    )

    # Header row is between the first two horizontal lines.
    student_row_bands = list(zip(horizontal_lines[1:-1], horizontal_lines[2:]))

    rows: list[dict[str, str]] = []
    for row_index, (row_top, row_bottom) in enumerate(student_row_bands, start=1):
        if row_bottom - row_top < 8:
            continue

        school_no_cell = crop_cell(page_image, vertical_lines[1], row_top, vertical_lines[2], row_bottom)
        name_cell = crop_cell(page_image, vertical_lines[2], row_top, vertical_lines[3], row_bottom)
        grade_cell = crop_cell(page_image, vertical_lines[3], row_top, vertical_lines[4], row_bottom)

        school_no_text = ocr_image(
            enhance_cell_for_ocr(school_no_cell, scale=2.6),
            lang=lang,
            psm=7,
            whitelist="0123456789",
        )
        name_text = ocr_image(
            enhance_cell_for_ocr(name_cell, scale=2.2),
            lang=lang,
            psm=7,
            preserve_spaces=True,
        )
        grade_text = ocr_image(
            enhance_cell_for_ocr(grade_cell, scale=2.4),
            lang=lang,
            psm=7,
            preserve_spaces=True,
        )

        school_no = normalize_school_no(school_no_text)
        name = normalize_name(name_text)
        grade = normalize_grade(grade_text)

        if not school_no and not name and not grade:
            continue

        if not school_no:
            errors.append(
                {
                    "pdf": str(pdf_path),
                    "page": str(page_number),
                    "row": str(row_index),
                    "reason": "missing school number",
                    "raw_text": f"name={name_text!r}; grade={grade_text!r}",
                }
            )
            continue

        if not name:
            errors.append(
                {
                    "pdf": str(pdf_path),
                    "page": str(page_number),
                    "row": str(row_index),
                    "reason": "missing student name",
                    "raw_text": f"school_no={school_no_text!r}; grade={grade_text!r}",
                }
            )
            continue

        if not grade:
            errors.append(
                {
                    "pdf": str(pdf_path),
                    "page": str(page_number),
                    "row": str(row_index),
                    "reason": "missing grade",
                    "raw_text": f"school_no={school_no_text!r}; name={name_text!r}",
                }
            )
            continue

        rows.append({"school_no": school_no, "name": name, "grade": grade})

    return rows


def empty_student_record(header: dict[str, str], school_no: str, name: str) -> dict[str, str]:
    record = {column: "" for column in OUTPUT_COLUMNS}
    record.update(
        {
            "İlçe": header["district"],
            "Okul": header["school"],
            "Sınıf": header["class_no"] or "3",
            "Şube": header["section"],
            "Okul No": school_no,
            "Ad Soyad": name,
        }
    )
    return record


def add_page_to_groups(
    groups: dict[GroupKey, dict[str, object]],
    page_result: PageResult,
) -> None:
    header = page_result.header
    key = GroupKey(
        district=header["district"] or "Bilinmeyen İlçe",
        school=header["school"] or "Bilinmeyen Okul",
        class_no=header["class_no"] or "3",
        section=header["section"] or "Bilinmeyen Şube",
    )

    if key not in groups:
        groups[key] = {"students": {}, "order": []}

    students = groups[key]["students"]
    student_order = groups[key]["order"]
    assert isinstance(students, dict)
    assert isinstance(student_order, list)

    for row in page_result.rows:
        school_no = row["school_no"]

        if school_no not in students:
            students[school_no] = empty_student_record(header, school_no, row["name"])
            student_order.append(school_no)
        else:
            student = students[school_no]
            assert isinstance(student, dict)
            for column, value in (
                ("İlçe", header["district"]),
                ("Okul", header["school"]),
                ("Sınıf", header["class_no"] or "3"),
                ("Şube", header["section"]),
                ("Ad Soyad", row["name"]),
            ):
                if value and not student.get(column):
                    student[column] = value

        student = students[school_no]
        assert isinstance(student, dict)
        student[page_result.course] = row["grade"]


def save_debug_images(
    debug_dir: Path,
    pdf_path: Path,
    page_number: int,
    page_image: Image.Image,
    processed_image: Image.Image,
) -> None:
    debug_dir.mkdir(parents=True, exist_ok=True)
    safe_stem = safe_file_part(pdf_path.stem)
    page_image.save(debug_dir / f"{safe_stem}_page{page_number}_render.png")
    processed_image.save(debug_dir / f"{safe_stem}_page{page_number}_processed.png")


def extract_pdf_pages(
    pdf_path: Path,
    dpi: int,
    lang: str,
    errors: list[dict[str, str]],
    debug_dir: Path | None = None,
) -> list[PageResult]:
    page_results: list[PageResult] = []
    previous_header: dict[str, str] | None = None
    previous_course = ""

    with fitz.open(str(pdf_path)) as document:
        for page_index in range(document.page_count):
            page_number = page_index + 1
            page = document.load_page(page_index)
            page_image = render_pdf_page(page, dpi=dpi)
            processed_image = preprocess_for_ocr(page_image)

            if debug_dir:
                save_debug_images(debug_dir, pdf_path, page_number, page_image, processed_image)

            header_height = int(page_image.height * 0.18)
            header_crop = page_image.crop((0, 0, page_image.width, header_height))
            header_ocr_image = enhance_cell_for_ocr(header_crop, scale=1.5)
            header_text = ocr_image(header_ocr_image, lang=lang, psm=6, preserve_spaces=True)

            header = extract_header_info(header_text)
            course = extract_course_name(header_text, pdf_name=pdf_path.name)

            if previous_header:
                for key, value in previous_header.items():
                    if not header.get(key) and value:
                        header[key] = value
            if not course and previous_course:
                course = previous_course

            if not course:
                errors.append(
                    {
                        "pdf": str(pdf_path),
                        "page": str(page_number),
                        "row": "",
                        "reason": "course not detected",
                        "raw_text": header_text,
                    }
                )
                continue

            page_rows = extract_rows_from_table(
                page_image=page_image,
                processed_image=processed_image,
                lang=lang,
                pdf_path=pdf_path,
                page_number=page_number,
                errors=errors,
            )

            if not page_rows:
                errors.append(
                    {
                        "pdf": str(pdf_path),
                        "page": str(page_number),
                        "row": "",
                        "reason": "no student rows extracted",
                        "raw_text": header_text,
                    }
                )
                continue

            previous_header = header
            previous_course = course
            page_results.append(
                PageResult(
                    pdf_path=pdf_path,
                    page_number=page_number,
                    header=header,
                    course=course,
                    rows=page_rows,
                )
            )

    return page_results


def rows_for_group(group_data: dict[str, object]) -> list[dict[str, str]]:
    students = group_data["students"]
    student_order = group_data["order"]
    assert isinstance(students, dict)
    assert isinstance(student_order, list)
    return [students[school_no] for school_no in student_order]


def save_excel(rows: list[dict[str, str]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Notlar"

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
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            cell.border = border

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for row_number in range(2, worksheet.max_row + 1):
        worksheet[f"E{row_number}"].number_format = "@"

    for column_cells in worksheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        max_length = max(len(str(cell.value or "")) for cell in column_cells)
        if column_letter == "F":
            width = min(max(max_length + 2, 18), 36)
        elif column_letter in {"A", "B"}:
            width = min(max(max_length + 2, 14), 34)
        elif column_letter in {"G", "H", "I", "J", "K", "L", "M", "N"}:
            width = min(max(max_length + 2, 14), 28)
        else:
            width = min(max(max_length + 2, 10), 18)
        worksheet.column_dimensions[column_letter].width = width

    worksheet.row_dimensions[1].height = 42
    workbook.save(output_path)


def safe_file_part(value: str) -> str:
    value = clean_text(value)
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value)
    value = re.sub(r"\s+", " ", value).strip(" ._")
    return value or "bilinmeyen"


def build_group_output_path(
    key: GroupKey,
    output_dir: Path,
    output_file: str | None,
    single_group: bool,
) -> Path:
    if output_file:
        if not single_group:
            raise ValueError(
                "--output-file can only be used when the input creates one Excel file."
            )
        output_path = Path(output_file).expanduser()
        if not output_path.suffix:
            output_path = output_path.with_suffix(".xlsx")
        if not output_path.is_absolute():
            output_path = output_dir / output_path
        return output_path

    file_name = (
        f"{safe_file_part(key.district)}_"
        f"{safe_file_part(key.school)}_"
        f"{safe_file_part(key.class_no)}"
        f"{safe_file_part(key.section)}_notlar.xlsx"
    )
    return output_dir / file_name


def write_error_log(errors: list[dict[str, str]], output_dir: Path) -> Path | None:
    if not errors:
        return None

    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / "ocr_extraction_errors.csv"
    fieldnames = ["pdf", "page", "row", "reason", "raw_text"]
    with log_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for error in errors:
            writer.writerow({field: error.get(field, "") for field in fieldnames})
    return log_path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract grade tables from image-like PDFs using OCR only."
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        default=str(DEFAULT_INPUT_PATH),
        help="Folder containing PDFs, one PDF path, or a PDF path without .pdf.",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help=(
            "Folder where Excel files will be saved. If omitted, folder input "
            "saves Excel files in that same folder; one-PDF input saves beside "
            "the PDF."
        ),
    )
    parser.add_argument(
        "--output-file",
        default=None,
        help="Optional Excel file name or full output path. Use only for one output workbook.",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=350,
        help="Rendering DPI. Use 300-400 for most cases. Default: 350.",
    )
    parser.add_argument(
        "--ocr-lang",
        default="tur+eng",
        help="Tesseract language code. Default: tur+eng.",
    )
    parser.add_argument(
        "--tesseract-cmd",
        default=None,
        help="Full path to tesseract.exe if it is not on PATH.",
    )
    parser.add_argument(
        "--debug-images",
        action="store_true",
        help="Save rendered and preprocessed page images for troubleshooting.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    configure_tesseract(args.tesseract_cmd)
    verify_tesseract_available()

    pdf_paths = resolve_input_pdfs(args.input_path)
    output_dir = (
        Path(args.output_dir).expanduser()
        if args.output_dir
        else default_output_dir(args.input_path, pdf_paths)
    )

    debug_dir = output_dir / "_ocr_debug_images" if args.debug_images else None
    groups: dict[GroupKey, dict[str, object]] = {}
    errors: list[dict[str, str]] = []
    failures: list[tuple[Path, Exception]] = []

    print(f"PDF files found: {len(pdf_paths)}")
    print(f"Excel output folder: {output_dir}")
    print(f"OCR language: {args.ocr_lang}")
    print(f"Render DPI: {args.dpi}")

    for index, pdf_path in enumerate(pdf_paths, start=1):
        try:
            page_results = extract_pdf_pages(
                pdf_path=pdf_path,
                dpi=args.dpi,
                lang=args.ocr_lang,
                errors=errors,
                debug_dir=debug_dir,
            )
            for page_result in page_results:
                add_page_to_groups(groups, page_result)
            row_entries = sum(len(page_result.rows) for page_result in page_results)
            print(f"[{index}/{len(pdf_paths)}] OK: {pdf_path.name} ({row_entries} row entries)")
        except Exception as error:
            failures.append((pdf_path, error))
            print(f"[{index}/{len(pdf_paths)}] ERROR: {pdf_path.name}: {error}", file=sys.stderr)

    if failures:
        for pdf_path, error in failures:
            errors.append(
                {
                    "pdf": str(pdf_path),
                    "page": "",
                    "row": "",
                    "reason": "pdf processing failed",
                    "raw_text": str(error),
                }
            )

    log_path = write_error_log(errors, output_dir)

    if not groups:
        if log_path:
            print(f"Error log saved: {log_path}", file=sys.stderr)
        print("No student data was extracted.", file=sys.stderr)
        return 1

    single_group = len(groups) == 1
    for key, group_data in sorted(groups.items(), key=lambda item: item[0]):
        rows = rows_for_group(group_data)
        output_path = build_group_output_path(key, output_dir, args.output_file, single_group)
        save_excel(rows, output_path)
        print(
            "Excel saved: "
            f"{output_path} ({len(rows)} students, "
            f"{key.district} / {key.school} / {key.class_no}-{key.section})"
        )

    if log_path:
        print(f"Error log saved: {log_path}")

    if failures:
        return 1

    print("OCR extraction completed.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
