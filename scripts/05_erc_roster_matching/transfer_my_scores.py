#!/usr/bin/env python3
"""Transfer manual ``My score`` values into a newly generated report copy.

The old report must contain the manually entered ``My score`` column on its
``Mismatched Students`` sheet. The new report is the clean workbook produced
after rerunning ``student_match_report.py``.

The clean report must have ``ERC District`` in column F. ``My score`` is
inserted immediately after it in column G, shifting the clean report's
original G-and-later columns one place to the right.

Every nonblank, non-formula score is transferred only when:

* Source Sheet, district, school, section, and student are exactly identical
  in one old row and one new row; and
* the ERC location, candidates, sections, and similarity values displayed for
  review are also exactly identical between the two reports.

There is no data-field normalization, fuzzy matching, or row-number matching.
Even a capitalization, Turkish-character, or whitespace difference is
reported and skipped. Header labels are located independently of their column
positions, and unrelated extra columns in the old report are ignored. Both
inputs are hash-verified as unchanged. Only a separately named output is
written.

Interactive use:

    python transfer_my_scores.py

Paths may also be supplied on the command line:

    python transfer_my_scores.py old_graded.xlsx new_report.xlsx output.xlsx
"""

from __future__ import annotations

import argparse
import hashlib
import math
import os
import re
import sys
import tempfile
import unicodedata
from collections import defaultdict
from copy import copy
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Sequence

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.worksheet import Worksheet


SUPPORTED_EXTENSION = ".xlsx"
MISMATCH_SHEET_TITLE = "Mismatched Students"
MY_SCORE_HEADER = "My score"
ERC_DISTRICT_COLUMN = 6  # Column F in the generated mismatch report.
MY_SCORE_COLUMN = 7  # Immediately after ERC District, in column G.

IDENTITY_HEADERS = (
    "Source Sheet",
    "Source District",
    "Source School",
    "Source Section",
    "Source Student Name",
)

BASE_MISMATCH_HEADERS = (
    "Source Sheet",
    "Source District",
    "Source School",
    "Source Section",
    "Source Student Name",
    "ERC District",
    "District Similarity",
    "ERC School",
    "School Similarity",
    "Best Candidate in Source Section",
    "Source-Section Similarity",
    "Source-Section ERC Section",
    "Best Candidate Across All Sections",
    "All-Sections Similarity",
    "All-Sections ERC Section",
    "Reason Not Matched",
)

SCORED_MISMATCH_HEADERS = (
    *BASE_MISMATCH_HEADERS[:6],
    MY_SCORE_HEADER,
    *BASE_MISMATCH_HEADERS[6:],
)

TEXT_CONTEXT_HEADERS = (
    "ERC District",
    "ERC School",
    "Best Candidate in Source Section",
    "Source-Section ERC Section",
    "Best Candidate Across All Sections",
    "All-Sections ERC Section",
)

NUMERIC_CONTEXT_HEADERS = (
    "District Similarity",
    "School Similarity",
    "Source-Section Similarity",
    "All-Sections Similarity",
)


ExactValue = tuple[str, object]
Identity = tuple[object, object, object, object, object]
IdentityKey = tuple[ExactValue, ExactValue, ExactValue, ExactValue, ExactValue]


@dataclass
class TransferStats:
    old_scored_rows: int = 0
    transferred: int = 0
    already_present: int = 0
    no_exact_match: int = 0
    ambiguous_old: int = 0
    ambiguous_new: int = 0
    changed_context: int = 0
    target_conflicts: int = 0
    unsafe_rows: int = 0
    formula_scores: int = 0
    inserted_score_column: bool = False
    warnings: list[str] = field(default_factory=list)
    old_sha256: str = ""
    new_sha256: str = ""

    @property
    def skipped(self) -> int:
        return (
            self.no_exact_match
            + self.ambiguous_old
            + self.ambiguous_new
            + self.changed_context
            + self.target_conflicts
            + self.unsafe_rows
            + self.formula_scores
        )


@dataclass(frozen=True)
class ScoreInstruction:
    old_row: int
    identity: Identity
    identity_key: IdentityKey
    score_value: object


@dataclass(frozen=True)
class PendingTransfer:
    instruction: ScoreInstruction
    new_row: int


def cell_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalized_text(value: object) -> str:
    """Normalize Turkish case, diacritics, and whitespace without fuzziness."""
    text = unicodedata.normalize("NFKC", cell_text(value))
    text = text.replace("I", "ı").replace("İ", "i").casefold().replace("ı", "i")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(text.split())


def normalized_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", normalized_text(value))


def is_blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def is_formula(value: object) -> bool:
    return isinstance(value, str) and value.lstrip().startswith("=")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def clean_user_path(value: str) -> str:
    cleaned = value.strip()
    if (
        len(cleaned) >= 2
        and cleaned[0] == cleaned[-1]
        and cleaned[0] in {'"', "'"}
    ):
        cleaned = cleaned[1:-1].strip()
    return cleaned


def resolved_input_path(value: str, label: str) -> Path:
    path = Path(value).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(f"{label} does not exist or is not a file: {path}")
    if path.suffix.lower() != SUPPORTED_EXTENSION:
        raise ValueError(f"{label} must be an .xlsx file.")
    return path


def resolved_output_path(value: str) -> Path:
    path = Path(value).expanduser().resolve()
    if path.suffix.lower() != SUPPORTED_EXTENSION:
        raise ValueError("The output workbook must use the .xlsx extension.")
    return path


def request_input_workbook(
    supplied_value: str | None,
    *,
    label: str,
    prompt: str,
) -> Path:
    if supplied_value is not None:
        return resolved_input_path(clean_user_path(supplied_value), label)

    while True:
        try:
            entered_value = input(prompt)
        except EOFError as error:
            raise ValueError(f"No path was supplied for {label.lower()}.") from error

        try:
            return resolved_input_path(clean_user_path(entered_value), label)
        except (FileNotFoundError, ValueError) as error:
            print(f"Error: {error}", file=sys.stderr)
            print("Please enter the full path again.", file=sys.stderr)


def default_output_path(new_report_path: Path) -> Path:
    base_name = f"{new_report_path.stem}_with_my_scores"
    candidate = new_report_path.with_name(f"{base_name}.xlsx")
    counter = 2
    while candidate.exists():
        candidate = new_report_path.with_name(f"{base_name}_{counter}.xlsx")
        counter += 1
    return candidate.resolve()


def paths_refer_to_same_file(first: Path, second: Path) -> bool:
    if first == second:
        return True
    try:
        return first.exists() and second.exists() and first.samefile(second)
    except OSError:
        return False


def save_workbook_atomically(workbook, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{output_path.stem}_",
            suffix=".xlsx",
            dir=output_path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        workbook.save(temporary_path)
        os.replace(temporary_path, output_path)
        temporary_path = None
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


def header_positions(worksheet: Worksheet) -> dict[str, list[int]]:
    positions: dict[str, list[int]] = defaultdict(list)
    for cell in worksheet[1]:
        key = normalized_header(cell.value)
        if key:
            positions[key].append(cell.column)
    return positions


def locate_headers(
    worksheet: Worksheet,
    required_headers: Sequence[str],
) -> dict[str, int]:
    positions = header_positions(worksheet)
    result: dict[str, int] = {}
    missing: list[str] = []
    duplicates: list[str] = []

    for header in required_headers:
        columns = positions.get(normalized_header(header), [])
        if not columns:
            missing.append(header)
        elif len(columns) > 1:
            duplicates.append(header)
        else:
            result[header] = columns[0]

    problems: list[str] = []
    if missing:
        problems.append(f"missing header(s): {', '.join(missing)}")
    if duplicates:
        problems.append(f"duplicate header(s): {', '.join(duplicates)}")
    if problems:
        raise ValueError(
            f"Sheet '{worksheet.title}' has " + "; ".join(problems) + "."
        )
    return result


def optional_header(worksheet: Worksheet, header: str) -> int | None:
    columns = header_positions(worksheet).get(normalized_header(header), [])
    if len(columns) > 1:
        raise ValueError(
            f"Sheet '{worksheet.title}' has duplicate '{header}' headers."
        )
    return columns[0] if columns else None


def header_prefix_differences(
    worksheet: Worksheet,
    expected_headers: Sequence[str],
) -> list[str]:
    """Compare the required header prefix and ignore later extra columns."""
    differences: list[str] = []
    for column, expected_header in enumerate(expected_headers, start=1):
        actual_value = worksheet.cell(1, column).value
        if exact_value(actual_value) != exact_value(expected_header):
            differences.append(
                f"{get_column_letter(column)}={actual_value!r}; "
                f"expected {expected_header!r}"
            )
    return differences


def raise_header_prefix_error(
    *,
    report_label: str,
    report_path: Path,
    worksheet: Worksheet,
    expected_headers: Sequence[str],
    detected_erc_column: int,
) -> None:
    differences = header_prefix_differences(worksheet, expected_headers)
    if not differences:
        return

    preview = "; ".join(differences[:8])
    if len(differences) > 8:
        preview += f"; plus {len(differences) - 8} more"
    detected_letter = get_column_letter(detected_erc_column)
    raise ValueError(
        f"The {report_label} required Sheet-4 headers are not in the exact "
        f"expected order in '{report_path}' / '{worksheet.title}'. "
        f"'ERC District' was detected at column {detected_letter}. {preview}"
    )


def locate_mismatch_sheet(
    workbook,
    required_headers: Sequence[str],
) -> tuple[Worksheet, dict[str, int]]:
    if MISMATCH_SHEET_TITLE in workbook.sheetnames:
        worksheet = workbook[MISMATCH_SHEET_TITLE]
        return worksheet, locate_headers(worksheet, required_headers)

    if len(workbook.worksheets) < 4:
        raise ValueError(
            "The workbook has fewer than four sheets and no "
            f"'{MISMATCH_SHEET_TITLE}' sheet."
        )

    worksheet = workbook.worksheets[3]
    try:
        headers = locate_headers(worksheet, required_headers)
    except ValueError as error:
        raise ValueError(
            f"Could not identify '{MISMATCH_SHEET_TITLE}'. "
            "The fourth sheet does not have the required headers."
        ) from error
    return worksheet, headers


def identity_from_row(
    worksheet: Worksheet,
    row_number: int,
    headers: dict[str, int],
) -> Identity:
    return tuple(
        worksheet.cell(row_number, headers[header]).value
        for header in IDENTITY_HEADERS
    )  # type: ignore[return-value]


def exact_value(value: object) -> ExactValue:
    """Create a type-aware exact comparison key for an Excel cell value."""
    if value is None:
        return "blank", None
    if isinstance(value, bool):
        return "boolean", value
    if isinstance(value, (int, float, Decimal)):
        if isinstance(value, float) and not math.isfinite(value):
            return "float", repr(value)
        try:
            return "number", Decimal(str(value))
        except InvalidOperation:
            return "other", value
    if isinstance(value, str):
        return "text", value
    return type(value).__name__, value


def identity_key(identity: Identity) -> IdentityKey:
    return tuple(exact_value(value) for value in identity)  # type: ignore[return-value]


def identity_description(identity: Identity) -> str:
    # repr() keeps leading/trailing spaces visible in exact-match warnings.
    return (
        f"source sheet={identity[0]!r}, "
        f"district={identity[1]!r}, "
        f"school={identity[2]!r}, "
        f"section={identity[3]!r}, "
        f"student={identity[4]!r}"
    )


def shift_cell_range_for_insert(range_ref: str, insert_column: int) -> str:
    cell_range = CellRange(range_ref)
    if cell_range.max_col < insert_column:
        return str(cell_range)
    if cell_range.min_col >= insert_column:
        cell_range.shift(col_shift=1)
    else:
        cell_range.max_col += 1
    return str(cell_range)


def shifted_coordinate_for_insert(
    coordinate: str | None,
    insert_column: int,
) -> str | None:
    if not coordinate:
        return coordinate
    cell_range = CellRange(f"{coordinate}:{coordinate}")
    if cell_range.min_col >= insert_column:
        cell_range.shift(col_shift=1)
    return f"{get_column_letter(cell_range.min_col)}{cell_range.min_row}"


def clone_column_dimension(
    source_dimension,
    worksheet: Worksheet,
    target_index: int,
) -> None:
    target_letter = get_column_letter(target_index)
    cloned = copy(source_dimension)
    cloned.index = target_letter
    cloned.min = target_index
    cloned.max = target_index
    cloned.worksheet = worksheet
    worksheet.column_dimensions[target_letter] = cloned


def insert_my_score_column(
    worksheet: Worksheet,
    *,
    insertion_column: int,
    old_score_width: float | None,
) -> int:
    """Insert My score while preserving shifted cells, widths, and ranges."""
    merged_ranges = [str(item) for item in worksheet.merged_cells.ranges]
    for merged_range in merged_ranges:
        worksheet.unmerge_cells(merged_range)

    original_dimensions: list[tuple[int, object]] = []
    for key, dimension in worksheet.column_dimensions.items():
        try:
            column_index = column_index_from_string(key)
        except ValueError:
            continue
        original_dimensions.append((column_index, copy(dimension)))

    original_filter_ref = worksheet.auto_filter.ref
    original_freeze_panes = (
        worksheet.freeze_panes.coordinate
        if hasattr(worksheet.freeze_panes, "coordinate")
        else worksheet.freeze_panes
    )
    original_table_refs = {
        table.name: table.ref for table in worksheet.tables.values()
    }

    worksheet.insert_cols(insertion_column, 1)

    for column_index, dimension in sorted(
        original_dimensions, reverse=True, key=lambda item: item[0]
    ):
        if column_index >= insertion_column:
            clone_column_dimension(dimension, worksheet, column_index + 1)

    score_letter = get_column_letter(insertion_column)
    worksheet.column_dimensions[score_letter].width = old_score_width or 14.0

    header_cell = worksheet.cell(1, insertion_column)
    template_header = worksheet.cell(1, max(1, insertion_column - 1))
    header_cell._style = copy(template_header._style)
    header_cell.value = MY_SCORE_HEADER

    for row_number in range(2, worksheet.max_row + 1):
        score_cell = worksheet.cell(row_number, insertion_column)
        template_cell = worksheet.cell(row_number, max(1, insertion_column - 1))
        score_cell._style = copy(template_cell._style)
        score_cell.value = None
        score_cell.number_format = "General"
        score_alignment = copy(score_cell.alignment)
        score_alignment.horizontal = "center"
        score_cell.alignment = score_alignment

    for merged_range in merged_ranges:
        worksheet.merge_cells(
            shift_cell_range_for_insert(merged_range, insertion_column)
        )

    if original_filter_ref:
        worksheet.auto_filter.ref = shift_cell_range_for_insert(
            original_filter_ref, insertion_column
        )
    worksheet.freeze_panes = shifted_coordinate_for_insert(
        original_freeze_panes, insertion_column
    )
    for table_name, table_ref in original_table_refs.items():
        worksheet.tables[table_name].ref = shift_cell_range_for_insert(
            table_ref, insertion_column
        )
    return insertion_column


def context_differences(
    old_sheet: Worksheet,
    old_row: int,
    old_headers: dict[str, int],
    new_sheet: Worksheet,
    new_row: int,
    new_headers: dict[str, int],
) -> list[str]:
    """Return every ERC review field that is not exactly equal."""
    differences: list[str] = []
    for header in (*TEXT_CONTEXT_HEADERS, *NUMERIC_CONTEXT_HEADERS):
        old_value = old_sheet.cell(old_row, old_headers[header]).value
        new_value = new_sheet.cell(new_row, new_headers[header]).value
        if exact_value(old_value) != exact_value(new_value):
            differences.append(header)
    return differences


def values_equivalent(left: object, right: object) -> bool:
    return exact_value(left) == exact_value(right)


def transfer_my_scores(
    old_report_path: Path,
    new_report_path: Path,
    output_path: Path,
) -> TransferStats:
    """Transfer scores into a new copy and return detailed statistics."""
    old_report_path = old_report_path.expanduser().resolve()
    new_report_path = new_report_path.expanduser().resolve()
    output_path = output_path.expanduser().resolve()

    for path, label in (
        (old_report_path, "Old graded report"),
        (new_report_path, "New report"),
    ):
        if not path.is_file():
            raise FileNotFoundError(f"{label} does not exist: {path}")
        if path.suffix.lower() != SUPPORTED_EXTENSION:
            raise ValueError(f"{label} must be an .xlsx file.")

    if old_report_path == new_report_path or paths_refer_to_same_file(
        old_report_path, new_report_path
    ):
        raise ValueError("The old graded report and new report must be different files.")
    if output_path.suffix.lower() != SUPPORTED_EXTENSION:
        raise ValueError("The output workbook must use the .xlsx extension.")
    if paths_refer_to_same_file(output_path, old_report_path) or paths_refer_to_same_file(
        output_path, new_report_path
    ):
        raise ValueError("The output path must differ from both input reports.")
    if output_path.exists():
        raise FileExistsError(
            f"Output already exists and will not be overwritten: {output_path}"
        )

    stats = TransferStats(
        old_sha256=sha256_file(old_report_path),
        new_sha256=sha256_file(new_report_path),
    )
    old_workbook = load_workbook(old_report_path, data_only=False)
    new_workbook = load_workbook(new_report_path, data_only=False)
    try:
        old_sheet, old_headers = locate_mismatch_sheet(
            old_workbook,
            (*BASE_MISMATCH_HEADERS, MY_SCORE_HEADER),
        )
        new_sheet, new_headers = locate_mismatch_sheet(
            new_workbook,
            BASE_MISMATCH_HEADERS,
        )

        old_score_column = old_headers[MY_SCORE_HEADER]
        if old_score_column != MY_SCORE_COLUMN:
            raise ValueError(
                "The old graded report must have the score-transfer column "
                "'My score' in column G. No transfer was performed."
            )

        raise_header_prefix_error(
            report_label="old graded report",
            report_path=old_report_path,
            worksheet=old_sheet,
            expected_headers=SCORED_MISMATCH_HEADERS,
            detected_erc_column=old_headers["ERC District"],
        )

        new_erc_district_column = new_headers["ERC District"]
        if new_erc_district_column != ERC_DISTRICT_COLUMN:
            raise ValueError(
                "The new report has 'ERC District' outside column F. "
                "No transfer was performed."
            )
        expected_score_column = new_erc_district_column + 1
        if expected_score_column != MY_SCORE_COLUMN:
            raise ValueError(
                "The new report layout does not place column G immediately "
                "after 'ERC District'. No transfer was performed."
            )

        new_score_column = optional_header(new_sheet, MY_SCORE_HEADER)
        if new_score_column is None:
            raise_header_prefix_error(
                report_label="new report",
                report_path=new_report_path,
                worksheet=new_sheet,
                expected_headers=BASE_MISMATCH_HEADERS,
                detected_erc_column=new_erc_district_column,
            )
            old_score_letter = get_column_letter(old_score_column)
            old_score_width = old_sheet.column_dimensions[old_score_letter].width
            new_score_column = insert_my_score_column(
                new_sheet,
                insertion_column=expected_score_column,
                old_score_width=old_score_width,
            )
            stats.inserted_score_column = True
            new_headers = locate_headers(
                new_sheet,
                (*BASE_MISMATCH_HEADERS, MY_SCORE_HEADER),
            )
        else:
            if (
                new_score_column != new_erc_district_column + 1
                or new_score_column != MY_SCORE_COLUMN
            ):
                raise ValueError(
                    "The new report must have 'My score' immediately after "
                    "'ERC District', in column G. No transfer was performed."
            )
            new_headers[MY_SCORE_HEADER] = new_score_column

        if (
            new_headers["ERC District"] != ERC_DISTRICT_COLUMN
            or new_headers[MY_SCORE_HEADER] != MY_SCORE_COLUMN
        ):
            raise ValueError(
                "The resulting Sheet-4 layout is invalid: 'ERC District' must "
                "be in F and 'My score' must be immediately after it in G."
            )

        raise_header_prefix_error(
            report_label="new report after adding My score",
            report_path=new_report_path,
            worksheet=new_sheet,
            expected_headers=SCORED_MISMATCH_HEADERS,
            detected_erc_column=new_headers["ERC District"],
        )

        old_exact_rows: dict[IdentityKey, list[int]] = defaultdict(list)
        for row_number in range(2, old_sheet.max_row + 1):
            identity = identity_from_row(old_sheet, row_number, old_headers)
            if any(not is_blank(value) for value in identity):
                old_exact_rows[identity_key(identity)].append(row_number)

        new_exact_rows: dict[IdentityKey, list[int]] = defaultdict(list)
        for row_number in range(2, new_sheet.max_row + 1):
            identity = identity_from_row(new_sheet, row_number, new_headers)
            if any(not is_blank(value) for value in identity):
                new_exact_rows[identity_key(identity)].append(row_number)

        instructions: list[ScoreInstruction] = []
        for old_row in range(2, old_sheet.max_row + 1):
            score_value = old_sheet.cell(old_row, old_score_column).value
            if is_blank(score_value):
                continue
            stats.old_scored_rows += 1
            identity = identity_from_row(old_sheet, old_row, old_headers)
            description = identity_description(identity)

            if is_formula(score_value):
                stats.formula_scores += 1
                stats.warnings.append(
                    f"Old Sheet 4 row {old_row} ({description}) has a formula "
                    "in My score; formulas are not transferred."
                )
                continue
            if any(is_blank(value) for value in identity):
                stats.unsafe_rows += 1
                stats.warnings.append(
                    f"Old Sheet 4 row {old_row} ({description}) has a blank "
                    "identity field and was skipped."
                )
                continue

            key = identity_key(identity)
            if len(old_exact_rows[key]) > 1:
                stats.ambiguous_old += 1
                stats.warnings.append(
                    f"Old Sheet 4 row {old_row} ({description}) shares its full "
                    f"identity with {len(old_exact_rows[key])} old rows and was "
                    "skipped."
                )
                continue

            instructions.append(
                ScoreInstruction(
                    old_row=old_row,
                    identity=identity,
                    identity_key=key,
                    score_value=score_value,
                )
            )

        pending: list[PendingTransfer] = []
        for instruction in instructions:
            target_rows = new_exact_rows.get(instruction.identity_key, [])
            if not target_rows:
                stats.no_exact_match += 1
                stats.warnings.append(
                    f"Old Sheet 4 row {instruction.old_row} "
                    f"({identity_description(instruction.identity)}) has no "
                    "100%-identical Source Sheet/district/school/section/student "
                    "row in the new mismatch sheet. It was not transferred."
                )
                continue

            if len(target_rows) > 1:
                stats.ambiguous_new += 1
                stats.warnings.append(
                    f"Old Sheet 4 row {instruction.old_row} "
                    f"({identity_description(instruction.identity)}) matched "
                    f"{len(target_rows)} exactly identical new rows and "
                    "was skipped."
                )
                continue

            new_row = target_rows[0]
            changed_fields = context_differences(
                old_sheet,
                instruction.old_row,
                old_headers,
                new_sheet,
                new_row,
                new_headers,
            )
            if changed_fields:
                stats.changed_context += 1
                stats.warnings.append(
                    f"Old Sheet 4 row {instruction.old_row} "
                    f"({identity_description(instruction.identity)}) has "
                    "non-identical ERC field(s) in the new report: "
                    f"{', '.join(changed_fields)}. It was not transferred."
                )
                continue

            pending.append(
                PendingTransfer(instruction=instruction, new_row=new_row)
            )

        pending_by_new_row: dict[int, list[PendingTransfer]] = defaultdict(list)
        for transfer in pending:
            pending_by_new_row[transfer.new_row].append(transfer)

        for new_row, transfers in pending_by_new_row.items():
            if len(transfers) > 1:
                for transfer in transfers:
                    stats.ambiguous_new += 1
                    stats.warnings.append(
                        f"Old Sheet 4 row {transfer.instruction.old_row} "
                        f"({identity_description(transfer.instruction.identity)}) "
                        "collides with another score instruction for the same new "
                        "row and was skipped."
                    )
                continue

            transfer = transfers[0]
            target_cell = new_sheet.cell(new_row, new_score_column)
            if not is_blank(target_cell.value):
                if values_equivalent(
                    target_cell.value, transfer.instruction.score_value
                ):
                    stats.already_present += 1
                else:
                    stats.target_conflicts += 1
                    stats.warnings.append(
                        f"New Sheet 4 row {new_row} "
                        f"({identity_description(transfer.instruction.identity)}) "
                        "already has a different My score value and was not "
                        "overwritten."
                    )
                continue

            target_cell.value = transfer.instruction.score_value
            stats.transferred += 1

        if sha256_file(old_report_path) != stats.old_sha256:
            raise RuntimeError(
                "The old graded report changed during processing; the output "
                "was not written."
            )
        if sha256_file(new_report_path) != stats.new_sha256:
            raise RuntimeError(
                "The new report changed during processing; the output was not "
                "written."
            )

        save_workbook_atomically(new_workbook, output_path)
    finally:
        old_workbook.close()
        new_workbook.close()

    if sha256_file(old_report_path) != stats.old_sha256:
        raise RuntimeError("The old graded report changed while saving the output.")
    if sha256_file(new_report_path) != stats.new_sha256:
        raise RuntimeError("The new report changed while saving the output.")
    return stats


def parse_arguments(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Transfer My score values from an old manually graded report into "
            "a new regenerated report copy."
        )
    )
    parser.add_argument(
        "old_report",
        nargs="?",
        help="Full path of the old manually graded .xlsx report",
    )
    parser.add_argument(
        "new_report",
        nargs="?",
        help="Full path of the newly regenerated .xlsx report",
    )
    parser.add_argument(
        "output_workbook",
        nargs="?",
        help="Optional full path for the new score-transferred workbook",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_arguments(argv)
    try:
        old_report_path = request_input_workbook(
            args.old_report,
            label="Old manually graded report",
            prompt="Enter the full path of the OLD manually graded report: ",
        )
        new_report_path = request_input_workbook(
            args.new_report,
            label="New regenerated report",
            prompt="Enter the full path of the NEW regenerated report: ",
        )
        output_path = (
            resolved_output_path(clean_user_path(args.output_workbook))
            if args.output_workbook is not None
            else default_output_path(new_report_path)
        )
        if args.output_workbook is None:
            print(f"The score-transferred copy will be created at: {output_path}")

        stats = transfer_my_scores(
            old_report_path,
            new_report_path,
            output_path,
        )
    except (
        FileNotFoundError,
        FileExistsError,
        PermissionError,
        RuntimeError,
        ValueError,
    ) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 2

    for warning in stats.warnings:
        print(f"WARNING: {warning}", file=sys.stderr)

    print(
        f"Old nonblank score rows: {stats.old_scored_rows} | "
        f"transferred: {stats.transferred} | "
        f"already present: {stats.already_present} | "
        f"no 100%-identical source identity: {stats.no_exact_match} | "
        f"non-identical ERC context: {stats.changed_context} | "
        f"ambiguous skipped: {stats.ambiguous_old + stats.ambiguous_new} | "
        f"other skipped: "
        f"{stats.target_conflicts + stats.unsafe_rows + stats.formula_scores}"
    )
    print("Both input workbooks are unchanged (SHA-256 verified).")
    print(f"Score-transferred workbook written to: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
