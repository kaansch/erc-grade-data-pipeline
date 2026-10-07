# -*- coding: utf-8 -*-
r"""
OCR extractor for type-3 image PDFs where each PDF is one course for all şubes.

This is a variant of extract_pdf_grades_type3_image_ocr.py.

Use this version when:
- Each PDF contains only one course.
- The pages inside that PDF are different class sections/şubes.
- The table is image/scanned, not selectable text.
- "SERBEST ETKİNLİKLER" should be skipped.

If you pass a folder containing several course PDFs, the script combines them
into one Excel file, matching students by İlçe/Okul/Sınıf/Şube/Okul No.

Install requirements in the Python environment used by VS Code:

    pip install pymupdf pillow pytesseract numpy openpyxl

You must also install the Tesseract OCR program itself. On Windows, if Tesseract
is not on PATH, pass its full path with --tesseract-cmd.

Run:

    python extract_pdf_grades_type3_course_all_subes_image_ocr.py "C:\path\to\pdf_or_folder"
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

try:
    import extract_pdf_grades_type3_image_ocr as base
except ImportError as exc:
    raise SystemExit(
        "This script must stay in the same folder as "
        "extract_pdf_grades_type3_image_ocr.py."
    ) from exc


SECTION_FALLBACK = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")


def rotated_candidates(image):
    """Try all practical orientations; some source PDFs are sideways, some are upright."""
    return [
        ("original", image),
        ("rot_minus_90", image.rotate(-90, expand=True, fillcolor="white")),
        ("rot_plus_90", image.rotate(90, expand=True, fillcolor="white")),
        ("rot_180", image.rotate(180, expand=True, fillcolor="white")),
    ]


def score_header_text(header_text: str) -> int:
    folded = base.fold_for_match(header_text)
    score = 0
    for token in ("ISTANBUL", "VALILIGI", "DERS", "SINIF", "SUBESI", "OGRETMEN", "CIZELGESI"):
        if token in folded:
            score += 2
    if base.canonical_course_from_text(header_text):
        score += 5
    header = base.extract_header_info([header_text])
    if header.get("section"):
        score += 4
    if header.get("class_no"):
        score += 2
    if header.get("district"):
        score += 2
    return score


def score_grid_geometry(grid) -> int:
    """Score whether the detected grid looks like the upright type-3 table."""
    intervals = list(zip(grid.x_lines, grid.x_lines[1:]))
    widths = [right - left for left, right in intervals]
    if not widths or grid.name_col >= len(widths):
        return 0

    score = 0
    name_width = widths[grid.name_col]
    median_width = sorted(widths)[len(widths) // 2]

    # In the real upright table the columns are: S.No, Okul No, wide name.
    if grid.school_no_col == 1:
        score += 8
    if grid.name_col == 2:
        score += 12
    elif grid.name_col in {1, 3}:
        score += 2

    if median_width and name_width >= median_width * 4:
        score += 10
    elif median_width and name_width >= median_width * 2:
        score += 4

    if 20 <= len(grid.row_intervals) <= 45:
        score += 4
    if 18 <= len(grid.x_lines) <= 45:
        score += 3

    return score


def try_oriented_page(orientation_name: str, image, lang: str):
    grid = base.detect_grid(image)
    header_text = ""
    try:
        header_text = base.ocr_header(image, grid, lang=lang)
    except Exception:
        # Keep the grid candidate. A failed header OCR should not make us
        # discard an otherwise valid table orientation.
        header_text = ""

    score = score_grid_geometry(grid) + score_header_text(header_text)
    return score, orientation_name, image, grid, header_text


def choose_oriented_page(raw_image, lang: str, preferred_orientation: str | None = None):
    candidates = rotated_candidates(raw_image)

    if preferred_orientation:
        preferred = [item for item in candidates if item[0] == preferred_orientation]
        if preferred:
            orientation_name, image = preferred[0]
            try:
                score, orientation_name, image, grid, header_text = try_oriented_page(
                    orientation_name,
                    image,
                    lang,
                )
                if score >= 20:
                    return orientation_name, image, grid, header_text
            except Exception:
                pass

    scored = []
    for orientation_name, image in candidates:
        try:
            scored.append(try_oriented_page(orientation_name, image, lang))
        except Exception:
            continue

    if not scored:
        raise ValueError("Could not orient page and detect its grade grid.")

    scored.sort(key=lambda item: item[0], reverse=True)
    _score, orientation_name, image, grid, header_text = scored[0]
    return orientation_name, image, grid, header_text


def merge_page_header(
    page_header: dict[str, str],
    common_header: dict[str, str],
    page_index: int,
) -> dict[str, str]:
    merged = {
        "district": page_header.get("district") or common_header.get("district", ""),
        "school": page_header.get("school") or common_header.get("school", ""),
        "class_no": page_header.get("class_no") or common_header.get("class_no", "") or "3",
        "section": page_header.get("section") or common_header.get("section", ""),
    }

    if not merged["section"] and 0 <= page_index < len(SECTION_FALLBACK):
        merged["section"] = SECTION_FALLBACK[page_index]

    return merged


def detect_pdf_course(header_texts: list[str], pdf_path: Path) -> str:
    for header_text in header_texts:
        course = base.canonical_course_from_text(header_text)
        if course:
            return course

    # Last-resort fallback for file names that include a course name.
    course = base.canonical_course_from_text(pdf_path.stem)
    return course


def extract_pdf(
    pdf_path: Path,
    dpi: int,
    lang: str,
    errors: list[dict[str, str]],
    debug_dir: Path | None,
    course_override: str = "",
) -> base.ParsedPdf:
    with base.fitz.open(str(pdf_path)) as document:
        page_images = []
        page_grids = []
        header_texts: list[str] = []
        orientations: list[str] = []
        preferred_orientation: str | None = None

        for page_index in range(document.page_count):
            page = document.load_page(page_index)
            raw_image = base.render_pdf_page(page, dpi=dpi)
            orientation_name, image, grid, header_text = choose_oriented_page(
                raw_image,
                lang=lang,
                preferred_orientation=preferred_orientation,
            )

            preferred_orientation = orientation_name
            orientations.append(orientation_name)
            page_images.append(image)
            page_grids.append(grid)
            header_texts.append(header_text)

        common_header = base.extract_header_info(header_texts)
        pdf_course = course_override or detect_pdf_course(header_texts, pdf_path)

        if not pdf_course:
            errors.append(
                {
                    "pdf": str(pdf_path),
                    "page": "",
                    "reason": "Could not detect course name from OCR header or file name.",
                    "raw_text": "\n\n".join(header_texts)[:800],
                }
            )
            return base.ParsedPdf(pdf_path=pdf_path, courses=[], header=common_header, rows=[])

        if base.is_excluded_course(pdf_course):
            return base.ParsedPdf(pdf_path=pdf_path, courses=[], header=common_header, rows=[])

        rows: list[dict[str, object]] = []
        for page_index, (image, grid, header_text, orientation_name) in enumerate(
            zip(page_images, page_grids, header_texts, orientations),
            start=0,
        ):
            page_header = base.extract_header_info([header_text])
            header = merge_page_header(page_header, common_header, page_index)

            page_rows, _identity_slots = base.extract_page_rows_from_grid(
                image=image,
                grid=grid,
                header=header,
                course=pdf_course,
                lang=lang,
                identity_slots=None,
            )
            rows.extend(page_rows)

            if not page_rows:
                errors.append(
                    {
                        "pdf": str(pdf_path),
                        "page": str(page_index + 1),
                        "reason": "No student rows parsed from detected grid.",
                        "raw_text": header_text[:500],
                    }
                )

            if debug_dir is not None:
                base.save_debug_files(
                    debug_dir=debug_dir,
                    pdf_path=pdf_path,
                    page_number=page_index + 1,
                    image=image,
                    grid=grid,
                    header_text=f"Orientation: {orientation_name}\n\n{header_text}",
                    course=pdf_course,
                    rows=page_rows,
                )

        return base.ParsedPdf(
            pdf_path=pdf_path,
            courses=[pdf_course],
            header=common_header,
            rows=rows,
        )


def build_output_path(
    pdf_paths: list[Path],
    input_is_folder: bool,
    input_base_dir: Path,
    output_dir: Path | None,
) -> Path:
    if output_dir is None:
        output_dir = input_base_dir / "type3_course_all_subes_ocr_output"

    if input_is_folder:
        return output_dir / "combined_grades.xlsx"

    return output_dir / f"{base.safe_file_part(pdf_paths[0].stem)}.xlsx"


def write_error_log(errors: list[dict[str, str]], output_dir: Path) -> Path | None:
    if not errors:
        return None

    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / "type3_course_all_subes_ocr_errors.csv"
    fieldnames = ["pdf", "page", "reason", "raw_text"]

    with log_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for error in errors:
            writer.writerow({field: error.get(field, "") for field in fieldnames})

    return log_path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Convert type-3 image/scanned one-course/all-şubes PDFs into Excel "
            "using grid-based OCR."
        )
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        default=None,
        help="PDF file path or folder path. If omitted, the script asks for it.",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Optional folder where the Excel file and error log will be saved.",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=350,
        help="Rendering DPI. Try 300, 350, or 400. Default: 350.",
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
        "--course",
        default="",
        help=(
            "Optional course name override, for example \"GÖRSEL SANATLAR\". "
            "Use only when the course title cannot be read by OCR."
        ),
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Save OCR header text and grid-overlay images for troubleshooting.",
    )
    return parser.parse_args(argv)


def get_input_path(args: argparse.Namespace) -> str:
    if args.input_path:
        return args.input_path

    user_input = input("Enter a PDF file path or a folder path containing PDF files: ").strip()
    if not user_input:
        raise ValueError("No input path was entered.")
    return user_input.strip('"')


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    base.configure_tesseract(args.tesseract_cmd)
    base.verify_tesseract_available()
    ocr_lang = base.resolve_ocr_language(args.ocr_lang)

    input_path = get_input_path(args)
    pdf_paths, input_is_folder, input_base_dir = base.resolve_input_pdfs(input_path)
    output_dir = Path(args.output_dir).expanduser() if args.output_dir else None
    output_path = build_output_path(pdf_paths, input_is_folder, input_base_dir, output_dir)
    log_output_dir = output_path.parent
    debug_dir = log_output_dir / "_type3_course_all_subes_ocr_debug" if args.debug else None

    parsed_pdfs: list[base.ParsedPdf] = []
    errors: list[dict[str, str]] = []

    print(f"PDF files found: {len(pdf_paths)}")
    print(f"OCR language: {ocr_lang}")
    print(f"Render DPI: {args.dpi}")

    for index, pdf_path in enumerate(pdf_paths, start=1):
        try:
            parsed = extract_pdf(
                pdf_path=pdf_path,
                dpi=args.dpi,
                lang=ocr_lang,
                errors=errors,
                debug_dir=debug_dir,
                course_override=base.clean_text(args.course),
            )

            if parsed.rows:
                parsed_pdfs.append(parsed)
                course_text = ", ".join(parsed.courses) if parsed.courses else "no included course"
                print(
                    f"[{index}/{len(pdf_paths)}] OK: {pdf_path.name} "
                    f"-> course={course_text}, rows={len(parsed.rows)}"
                )
            else:
                print(
                    f"[{index}/{len(pdf_paths)}] SKIPPED: {pdf_path.name} "
                    "(no included course rows found)"
                )

        except Exception as error:
            errors.append(
                {
                    "pdf": str(pdf_path),
                    "page": "",
                    "reason": str(error),
                    "raw_text": "",
                }
            )
            print(f"[{index}/{len(pdf_paths)}] ERROR: {pdf_path.name}: {error}", file=sys.stderr)

    if not parsed_pdfs:
        log_path = write_error_log(errors, log_output_dir)
        if log_path:
            print(f"Error log saved: {log_path}")
        raise RuntimeError("No PDFs could be converted.")

    rows, course_columns = base.combine_parsed_pdfs(parsed_pdfs)
    if not rows:
        raise RuntimeError("No student rows were extracted from the processed PDFs.")

    base.save_excel(rows, course_columns, output_path)
    log_path = write_error_log(errors, log_output_dir)

    print(f"Converted PDFs: {len(parsed_pdfs)}")
    print(f"Created: {output_path}")
    if log_path:
        print(f"Error log saved: {log_path}")
    if debug_dir:
        print(f"Debug files saved: {debug_dir}")

    return 1 if errors else 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
