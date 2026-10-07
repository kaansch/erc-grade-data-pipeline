#!/usr/bin/env python3
"""Create a reference-based student-roster reconciliation report.

The source workbook is opened with ``read_only=True`` and is never saved.  The
program calculates its SHA-256 digest before and after processing and refuses
to use the source path as the output path.

Expected source layout
----------------------

* First sheet:  A=district, B=school, C=grade, D=section, F=student name.
  Column E is intentionally ignored.
* Second sheet: A=district, B=school, C=grade, D=section, E=student name.
* Third sheet:  A=district, B=school, C=grade, D=section, E=student name.

The first two sheets must contain grade 3 and the third sheet grade 4.  Section
letters are strict matching boundaries: students are never matched across
different district-school-section tuples.

Where a district-school-section exists in Sheet 1, its students form the
authoritative baseline and are matched independently to Sheets 2 and 3. Where
Sheet 1 has no such section, Sheets 2 and 3 are compared directly. Differences
are reported as review exceptions; they are not automatically classified as
administrative errors because legitimate student movements may also exist.

Example
-------

Run without a path to receive an interactive prompt:

    python student_tracking.py

Or supply the path directly on the command line:

    python student_tracking.py rosters.xlsx -o student_tracking_output.xlsx

The first three worksheets are used by default.  Use ``--sheet-1``,
``--sheet-2``, and ``--sheet-3`` when explicit worksheet names are preferable.
Fuzzy thresholds are configurable; the defaults are deliberately conservative.
"""

from __future__ import annotations

import argparse
import hashlib
import math
import os
import re
import tempfile
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Callable, Iterable, Sequence

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
except ImportError as exc:  # pragma: no cover - environment-specific error
    raise SystemExit(
        "This program requires openpyxl. Install it with: pip install openpyxl"
    ) from exc


PERIOD_LABELS = {
    1: "1st Semester 3rd Grade",
    2: "2nd Semester 3rd Grade",
    3: "4th Grade",
}

PERIOD_DEFINITIONS = {
    1: {"expected_grade": 3, "student_column": 6},
    2: {"expected_grade": 3, "student_column": 5},
    3: {"expected_grade": 4, "student_column": 5},
}

OUTPUT_SHEETS = (
    "Overview",
    "Section Reconciliation",
    "Student Exceptions",
    "Unmatched Later Records",
    "Duplicate Names",
    "Match Review",
)

TURKISH_UPPER_TRANSLATION = str.maketrans({"i": "İ", "ı": "I", "ş": "Ş", "ğ": "Ğ", "ü": "Ü", "ö": "Ö", "ç": "Ç"})
TURKISH_ASCII_TRANSLATION = str.maketrans(
    {
        "Ç": "C",
        "Ğ": "G",
        "İ": "I",
        "I": "I",
        "Ö": "O",
        "Ş": "S",
        "Ü": "U",
        "ç": "c",
        "ğ": "g",
        "ı": "i",
        "i": "i",
        "ö": "o",
        "ş": "s",
        "ü": "u",
    }
)

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
GREEN_FONT = Font(color="006100", bold=True)
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
RED_FONT = Font(color="9C0006", bold=True)
ALT_FILL = PatternFill("solid", fgColor="D9EAF7")
THIN_GRAY = Side(style="thin", color="D9E1F2")
HEADER_BOTTOM = Side(style="medium", color="9EADBA")
# A light horizontal separator keeps long reports readable without boxing every cell.
CELL_BORDER = Border(bottom=THIN_GRAY)
HEADER_BORDER = Border(bottom=HEADER_BOTTOM)


@dataclass
class ReviewIssue:
    status: str
    issue_type: str
    district: str = ""
    school: str = ""
    section: str = ""
    period: str = ""
    sheet: str = ""
    source_row: int | str = ""
    original_value: str = ""
    candidate_value: str = ""
    similarity_score: float | None = None
    second_best_score: float | None = None
    reason: str = ""

    def as_excel_row(self) -> list[object]:
        return [
            self.status,
            self.issue_type,
            self.district,
            self.school,
            self.section,
            self.period,
            self.sheet,
            self.source_row,
            self.original_value,
            self.candidate_value,
            round(self.similarity_score, 2)
            if self.similarity_score is not None
            else "",
            round(self.second_best_score, 2)
            if self.second_best_score is not None
            else "",
            self.reason,
        ]


@dataclass
class SourceRow:
    period: int
    sheet_name: str
    row_number: int
    district_raw: str
    school_raw: str
    grade_raw: object
    section_raw: str
    student_raw: str
    district_norm: str
    school_norm: str
    section: str
    student_strict: str
    student_loose: str
    district_id: str = ""
    school_id: str = ""

    @property
    def has_student(self) -> bool:
        return bool(self.student_raw and self.student_strict)

    @property
    def tuple_key(self) -> tuple[str, str, str]:
        if not self.district_id or not self.school_id:
            raise RuntimeError("Canonical district/school IDs have not been assigned.")
        return (self.district_id, self.school_id, self.section)

    @property
    def source_key(self) -> tuple[int, str, int]:
        return (self.period, self.sheet_name, self.row_number)


@dataclass(frozen=True)
class PairMatch:
    left: SourceRow
    right: SourceRow
    score: float
    method: str


@dataclass(frozen=True)
class CandidateInfo:
    record: SourceRow
    score: float


@dataclass
class PairwiseResult:
    left_period: int
    right_period: int
    matches: list[PairMatch]
    unmatched_left: list[SourceRow]
    unmatched_right: list[SourceRow]
    best_for_left: dict[tuple[int, str, int], CandidateInfo]
    best_for_right: dict[tuple[int, str, int], CandidateInfo]

    @property
    def right_by_left(self) -> dict[tuple[int, str, int], PairMatch]:
        return {match.left.source_key: match for match in self.matches}

    @property
    def left_by_right(self) -> dict[tuple[int, str, int], PairMatch]:
        return {match.right.source_key: match for match in self.matches}


@dataclass
class TupleComparison:
    tuple_key: tuple[str, str, str]
    reference_mode: str
    section_periods: set[int]
    records_by_period: dict[int, list[SourceRow]]
    sheet1_to_sheet2: PairwiseResult | None = None
    sheet1_to_sheet3: PairwiseResult | None = None
    sheet2_to_sheet3: PairwiseResult | None = None


@dataclass
class DuplicateGroup:
    period: int
    tuple_key: tuple[str, str, str]
    normalized_name: str
    records: list[SourceRow]


@dataclass
class IdFactory:
    prefix: str
    next_number: int = 1

    def new(self) -> str:
        value = f"{self.prefix}{self.next_number:05d}"
        self.next_number += 1
        return value


def cell_text(value: object) -> str:
    """Convert a cell value to trimmed text without changing its meaning."""

    if value is None:
        return ""
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def turkish_upper(value: str) -> str:
    return value.translate(TURKISH_UPPER_TRANSLATION).upper()


def normalize_strict(value: object) -> str:
    """Normalize harmless formatting while retaining Turkish letters."""

    text = unicodedata.normalize("NFKC", cell_text(value))
    text = turkish_upper(text)
    text = re.sub(r"[^0-9A-ZÇĞİÖŞÜ]+", " ", text)
    return " ".join(text.split())


def normalize_loose(value: object) -> str:
    """Create a secondary comparison form that folds Turkish diacritics."""

    text = normalize_strict(value).translate(TURKISH_ASCII_TRANSLATION)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"[^0-9A-Z]+", " ", text.upper())
    return " ".join(text.split())


def label_loose_form(kind: str, value: str) -> str:
    """Return a loose institutional comparison form."""

    text = normalize_loose(value)
    tokens = text.split()
    if kind == "district":
        tokens = [token for token in tokens if token not in {"ILCE", "ILCESI"}]
    return " ".join(tokens)


def sequence_ratio(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    return SequenceMatcher(None, left, right, autojunk=False).ratio() * 100.0


def token_sort(value: str) -> str:
    return " ".join(sorted(value.split()))


def text_similarity(left: str, right: str, *, kind: str) -> float:
    """Conservative score combining sequence and token-order comparisons."""

    if not left or not right:
        return 0.0
    if left == right:
        return 100.0

    left_loose = (
        normalize_loose(left) if kind == "student" else label_loose_form(kind, left)
    )
    right_loose = (
        normalize_loose(right) if kind == "student" else label_loose_form(kind, right)
    )
    scores = (
        sequence_ratio(left, right),
        sequence_ratio(left_loose, right_loose),
        sequence_ratio(token_sort(left_loose), token_sort(right_loose)),
    )
    return max(scores)


def grade_matches(value: object, expected: int) -> bool:
    if value is None or cell_text(value) == "":
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        try:
            return float(value) == float(expected)
        except (TypeError, ValueError):
            return False
    text = cell_text(value).replace(",", ".")
    try:
        return float(text) == float(expected)
    except ValueError:
        return normalize_strict(text) == str(expected)


def normalize_section(value: object) -> str:
    section = normalize_strict(value).replace(" ", "")
    if not re.fullmatch(r"[A-ZÇĞİÖŞÜ]", section):
        return ""
    return section


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def row_value(row: Sequence[object], one_based_column: int) -> object:
    index = one_based_column - 1
    return row[index] if index < len(row) else None


def resolve_sheet_names(
    available: Sequence[str], requested: Sequence[str | None]
) -> tuple[str, str, str]:
    if len(available) < 3:
        raise ValueError("The source workbook must contain at least three worksheets.")

    resolved: list[str] = []
    for index, requested_name in enumerate(requested):
        name = requested_name if requested_name else available[index]
        if name not in available:
            raise ValueError(
                f"Worksheet {name!r} was not found. Available sheets: {', '.join(available)}"
            )
        resolved.append(name)

    if len(set(resolved)) != 3:
        raise ValueError("The three period worksheet selections must be different.")
    return (resolved[0], resolved[1], resolved[2])


def read_source_rows(
    input_path: Path,
    requested_sheets: Sequence[str | None],
    header_row: int,
    issues: list[ReviewIssue],
) -> tuple[list[SourceRow], tuple[str, str, str]]:
    """Read source rows without creating a writable source workbook."""

    workbook = load_workbook(
        filename=input_path,
        read_only=True,
        data_only=True,
        keep_links=False,
    )
    try:
        selected_sheets = resolve_sheet_names(workbook.sheetnames, requested_sheets)
        rows: list[SourceRow] = []

        for period, sheet_name in enumerate(selected_sheets, start=1):
            worksheet = workbook[sheet_name]
            definition = PERIOD_DEFINITIONS[period]
            expected_grade = int(definition["expected_grade"])
            student_column = int(definition["student_column"])

            for row_number, values in enumerate(
                worksheet.iter_rows(min_row=header_row + 1, values_only=True),
                start=header_row + 1,
            ):
                district_raw = cell_text(row_value(values, 1))
                school_raw = cell_text(row_value(values, 2))
                grade_raw = row_value(values, 3)
                section_raw = cell_text(row_value(values, 4))
                student_raw = cell_text(row_value(values, student_column))

                if not any(
                    (district_raw, school_raw, cell_text(grade_raw), section_raw, student_raw)
                ):
                    continue

                if not district_raw or not school_raw or not section_raw:
                    issues.append(
                        ReviewIssue(
                            status="Data issue",
                            issue_type="Incomplete location data",
                            district=district_raw,
                            school=school_raw,
                            section=section_raw,
                            period=PERIOD_LABELS[period],
                            sheet=sheet_name,
                            source_row=row_number,
                            original_value=(
                                f"District={district_raw!r}; School={school_raw!r}; "
                                f"Section={section_raw!r}"
                            ),
                            reason=(
                                "District, school, and section are required for strict "
                                "district-school-section matching."
                            ),
                        )
                    )
                    continue

                if not grade_matches(grade_raw, expected_grade):
                    issues.append(
                        ReviewIssue(
                            status="Data issue",
                            issue_type="Unexpected grade",
                            district=district_raw,
                            school=school_raw,
                            section=section_raw,
                            period=PERIOD_LABELS[period],
                            sheet=sheet_name,
                            source_row=row_number,
                            original_value=cell_text(grade_raw),
                            candidate_value=str(expected_grade),
                            reason="The row was excluded because its grade did not match the period.",
                        )
                    )
                    continue

                section = normalize_section(section_raw)
                if not section:
                    issues.append(
                        ReviewIssue(
                            status="Data issue",
                            issue_type="Invalid section",
                            district=district_raw,
                            school=school_raw,
                            section=section_raw,
                            period=PERIOD_LABELS[period],
                            sheet=sheet_name,
                            source_row=row_number,
                            original_value=section_raw,
                            reason="The section must contain exactly one letter.",
                        )
                    )
                    continue

                if not student_raw:
                    issues.append(
                        ReviewIssue(
                            status="Data issue",
                            issue_type="Blank student name",
                            district=district_raw,
                            school=school_raw,
                            section=section,
                            period=PERIOD_LABELS[period],
                            sheet=sheet_name,
                            source_row=row_number,
                            reason=(
                                "The section is treated as present, but this row is not "
                                "counted as a student."
                            ),
                        )
                    )

                rows.append(
                    SourceRow(
                        period=period,
                        sheet_name=sheet_name,
                        row_number=row_number,
                        district_raw=district_raw,
                        school_raw=school_raw,
                        grade_raw=grade_raw,
                        section_raw=section_raw,
                        student_raw=student_raw,
                        district_norm=normalize_strict(district_raw),
                        school_norm=normalize_strict(school_raw),
                        section=section,
                        student_strict=normalize_strict(student_raw),
                        student_loose=normalize_loose(student_raw),
                    )
                )
        return rows, selected_sheets
    finally:
        workbook.close()


def earliest_occurrence(rows: Iterable[SourceRow]) -> SourceRow:
    return min(rows, key=lambda item: (item.period, item.row_number))


def cluster_similarity(left: set[str], right: set[str], *, kind: str) -> float:
    """Use complete-link similarity so every spelling in a merge is compatible."""

    return min(text_similarity(a, b, kind=kind) for a in left for b in right)


def resolve_label_variants(
    rows: Sequence[SourceRow],
    *,
    kind: str,
    norm_getter: Callable[[SourceRow], str],
    raw_getter: Callable[[SourceRow], str],
    threshold: float,
    margin: float,
    id_factory: IdFactory,
    issues: list[ReviewIssue],
    context_district: str = "",
) -> tuple[dict[str, str], dict[str, str]]:
    """Cluster district or school variants conservatively.

    Exact normalized values are one node.  Loose-identical values are merged,
    then mutually-best fuzzy cluster pairs may merge when both their threshold
    and confidence-margin tests pass.  Uncertain candidates stay separate.
    """

    occurrences: dict[str, list[SourceRow]] = defaultdict(list)
    for row in rows:
        occurrences[norm_getter(row)].append(row)

    if not occurrences:
        return {}, {}

    # Loose-identical spellings are safe normalization variants.
    loose_groups: dict[str, set[str]] = defaultdict(set)
    for normalized_value in occurrences:
        loose_groups[label_loose_form(kind, normalized_value)].add(normalized_value)
    clusters = [set(group) for group in loose_groups.values()]

    for cluster in clusters:
        if len(cluster) <= 1:
            continue
        ordered = sorted(cluster)
        representative = ordered[0]
        rep_row = earliest_occurrence(occurrences[representative])
        for variant in ordered[1:]:
            variant_row = earliest_occurrence(occurrences[variant])
            issues.append(
                ReviewIssue(
                    status="Auto-accepted",
                    issue_type=f"Normalized {kind} match",
                    district=(
                        context_district
                        if kind == "school"
                        else raw_getter(rep_row)
                    ),
                    school=raw_getter(rep_row) if kind == "school" else "",
                    period=(
                        f"{PERIOD_LABELS[rep_row.period]} vs "
                        f"{PERIOD_LABELS[variant_row.period]}"
                    ),
                    sheet=f"{rep_row.sheet_name} / {variant_row.sheet_name}",
                    source_row=f"{rep_row.row_number} / {variant_row.row_number}",
                    original_value=raw_getter(rep_row),
                    candidate_value=raw_getter(variant_row),
                    similarity_score=100.0,
                    reason="Loose normalized forms are identical.",
                )
            )

    while len(clusters) > 1:
        candidate_lists: dict[int, list[tuple[float, int]]] = defaultdict(list)
        for left_index in range(len(clusters)):
            for right_index in range(left_index + 1, len(clusters)):
                score = cluster_similarity(
                    clusters[left_index], clusters[right_index], kind=kind
                )
                candidate_lists[left_index].append((score, right_index))
                candidate_lists[right_index].append((score, left_index))

        best: dict[int, tuple[float, int, float]] = {}
        for index, candidates in candidate_lists.items():
            candidates.sort(key=lambda item: (-item[0], item[1]))
            best_score, best_index = candidates[0]
            second_score = candidates[1][0] if len(candidates) > 1 else 0.0
            best[index] = (best_score, best_index, second_score)

        proposals: list[tuple[float, int, int, float]] = []
        for left_index, (score, right_index, left_second) in best.items():
            if left_index >= right_index or right_index not in best:
                continue
            right_score, right_best, right_second = best[right_index]
            if right_best != left_index:
                continue
            if score < threshold or right_score < threshold:
                continue
            if score - left_second < margin or right_score - right_second < margin:
                continue
            proposals.append(
                (score, left_index, right_index, max(left_second, right_second))
            )

        if not proposals:
            break

        score, left_index, right_index, second_score = max(
            proposals, key=lambda item: (item[0], -item[1], -item[2])
        )
        left_cluster = clusters[left_index]
        right_cluster = clusters[right_index]
        left_node = min(
            left_cluster,
            key=lambda node: (
                earliest_occurrence(occurrences[node]).period,
                earliest_occurrence(occurrences[node]).row_number,
            ),
        )
        right_node = min(
            right_cluster,
            key=lambda node: (
                earliest_occurrence(occurrences[node]).period,
                earliest_occurrence(occurrences[node]).row_number,
            ),
        )
        left_row = earliest_occurrence(occurrences[left_node])
        right_row = earliest_occurrence(occurrences[right_node])
        issues.append(
            ReviewIssue(
                status="Auto-accepted",
                issue_type=f"Fuzzy {kind} match",
                district=(
                    context_district if kind == "school" else raw_getter(left_row)
                ),
                school=raw_getter(left_row) if kind == "school" else "",
                period=(
                    f"{PERIOD_LABELS[left_row.period]} vs "
                    f"{PERIOD_LABELS[right_row.period]}"
                ),
                sheet=f"{left_row.sheet_name} / {right_row.sheet_name}",
                source_row=f"{left_row.row_number} / {right_row.row_number}",
                original_value=raw_getter(left_row),
                candidate_value=raw_getter(right_row),
                similarity_score=score,
                second_best_score=second_score,
                reason=(
                    f"Mutually best {kind} candidates above threshold {threshold:.1f} "
                    f"with margin at least {margin:.1f}."
                ),
            )
        )
        merged = left_cluster | right_cluster
        clusters = [
            cluster
            for index, cluster in enumerate(clusters)
            if index not in {left_index, right_index}
        ]
        clusters.append(merged)

    # Record the best remaining plausible pair for each unresolved cluster.
    review_floor = max(0.0, threshold - 10.0)
    seen_pairs: set[tuple[int, int]] = set()
    for left_index, left_cluster in enumerate(clusters):
        candidates: list[tuple[float, int]] = []
        for right_index, right_cluster in enumerate(clusters):
            if left_index == right_index:
                continue
            candidates.append(
                (
                    cluster_similarity(left_cluster, right_cluster, kind=kind),
                    right_index,
                )
            )
        if not candidates:
            continue
        candidates.sort(key=lambda item: (-item[0], item[1]))
        score, right_index = candidates[0]
        if score < review_floor:
            continue
        pair = tuple(sorted((left_index, right_index)))
        if pair in seen_pairs:
            continue
        seen_pairs.add(pair)
        second_score = candidates[1][0] if len(candidates) > 1 else 0.0
        left_node = min(
            left_cluster,
            key=lambda node: (
                earliest_occurrence(occurrences[node]).period,
                earliest_occurrence(occurrences[node]).row_number,
            ),
        )
        right_cluster = clusters[right_index]
        right_node = min(
            right_cluster,
            key=lambda node: (
                earliest_occurrence(occurrences[node]).period,
                earliest_occurrence(occurrences[node]).row_number,
            ),
        )
        left_row = earliest_occurrence(occurrences[left_node])
        right_row = earliest_occurrence(occurrences[right_node])
        if score < threshold:
            reason = f"Best score is below automatic threshold {threshold:.1f}."
        elif score - second_score < margin:
            reason = f"Competing candidates are within the required margin {margin:.1f}."
        else:
            reason = "The candidate relationship was not mutually best."
        issues.append(
            ReviewIssue(
                status="Needs review",
                issue_type=f"Ambiguous {kind} match",
                district=(
                    context_district if kind == "school" else raw_getter(left_row)
                ),
                school=raw_getter(left_row) if kind == "school" else "",
                period=(
                    f"{PERIOD_LABELS[left_row.period]} vs "
                    f"{PERIOD_LABELS[right_row.period]}"
                ),
                sheet=f"{left_row.sheet_name} / {right_row.sheet_name}",
                source_row=f"{left_row.row_number} / {right_row.row_number}",
                original_value=raw_getter(left_row),
                candidate_value=raw_getter(right_row),
                similarity_score=score,
                second_best_score=second_score,
                reason=reason + " Values were kept separate.",
            )
        )

    def cluster_order(cluster: set[str]) -> tuple[int, int, str]:
        record = min(
            (row for node in cluster for row in occurrences[node]),
            key=lambda item: (item.period, item.row_number),
        )
        return (record.period, record.row_number, raw_getter(record))

    normalized_to_id: dict[str, str] = {}
    display_by_id: dict[str, str] = {}
    for cluster in sorted(clusters, key=cluster_order):
        canonical_id = id_factory.new()
        canonical_row = min(
            (row for node in cluster for row in occurrences[node]),
            key=lambda item: (item.period, item.row_number),
        )
        display_by_id[canonical_id] = raw_getter(canonical_row)
        for node in cluster:
            normalized_to_id[node] = canonical_id

    return normalized_to_id, display_by_id


def resolve_institutions(
    rows: Sequence[SourceRow],
    *,
    district_threshold: float,
    district_margin: float,
    school_threshold: float,
    school_margin: float,
    issues: list[ReviewIssue],
) -> tuple[dict[str, str], dict[str, str]]:
    district_factory = IdFactory("D")
    district_mapping, district_display = resolve_label_variants(
        rows,
        kind="district",
        norm_getter=lambda item: item.district_norm,
        raw_getter=lambda item: item.district_raw,
        threshold=district_threshold,
        margin=district_margin,
        id_factory=district_factory,
        issues=issues,
    )
    for row in rows:
        row.district_id = district_mapping[row.district_norm]

    school_factory = IdFactory("S")
    school_display: dict[str, str] = {}
    rows_by_district: dict[str, list[SourceRow]] = defaultdict(list)
    for row in rows:
        rows_by_district[row.district_id].append(row)

    for district_id in sorted(
        rows_by_district, key=lambda key: normalize_loose(district_display[key])
    ):
        district_rows = rows_by_district[district_id]
        school_mapping, group_display = resolve_label_variants(
            district_rows,
            kind="school",
            norm_getter=lambda item: item.school_norm,
            raw_getter=lambda item: item.school_raw,
            threshold=school_threshold,
            margin=school_margin,
            id_factory=school_factory,
            issues=issues,
            context_district=district_display[district_id],
        )
        school_display.update(group_display)
        for row in district_rows:
            row.school_id = school_mapping[row.school_norm]

    return district_display, school_display


def detect_duplicate_names(records: Sequence[SourceRow]) -> list[DuplicateGroup]:
    groups: dict[
        tuple[int, tuple[str, str, str], str], list[SourceRow]
    ] = defaultdict(list)
    for record in records:
        groups[(record.period, record.tuple_key, record.student_strict)].append(record)

    duplicates = [
        DuplicateGroup(
            period=period,
            tuple_key=tuple_key,
            normalized_name=normalized_name,
            records=sorted(group, key=lambda item: item.row_number),
        )
        for (period, tuple_key, normalized_name), group in groups.items()
        if len(group) > 1
    ]
    return duplicates


def match_student_lists(
    left_records: Sequence[SourceRow],
    right_records: Sequence[SourceRow],
    *,
    threshold: float,
    margin: float,
    district_display: dict[str, str],
    school_display: dict[str, str],
    issues: list[ReviewIssue],
) -> PairwiseResult:
    """Match two rosters one-to-one within one strict location tuple.

    Strict normalized names are paired first while preserving duplicate
    multiplicity. Remaining rows are matched only when they are mutually best,
    exceed the configured threshold, and have a sufficient alternative-candidate
    margin. Ambiguous rows remain unmatched and are reported for review.
    """

    if left_records:
        left_period = left_records[0].period
    elif right_records:
        # Callers always compare adjacent named periods and replace this below.
        left_period = max(1, right_records[0].period - 1)
    else:
        raise ValueError("At least one comparison side must identify a period.")
    right_period = right_records[0].period if right_records else left_period + 1

    if any(record.period != left_period for record in left_records):
        raise RuntimeError("Left comparison list contains more than one period.")
    if any(record.period != right_period for record in right_records):
        raise RuntimeError("Right comparison list contains more than one period.")

    matches: list[PairMatch] = []
    left_by_name: dict[str, list[SourceRow]] = defaultdict(list)
    right_by_name: dict[str, list[SourceRow]] = defaultdict(list)
    for record in left_records:
        left_by_name[record.student_strict].append(record)
    for record in right_records:
        right_by_name[record.student_strict].append(record)

    matched_left: set[tuple[int, str, int]] = set()
    matched_right: set[tuple[int, str, int]] = set()
    for normalized_name in sorted(set(left_by_name) & set(right_by_name)):
        left_group = sorted(left_by_name[normalized_name], key=lambda item: item.row_number)
        right_group = sorted(right_by_name[normalized_name], key=lambda item: item.row_number)
        for left, right in zip(left_group, right_group):
            matches.append(PairMatch(left=left, right=right, score=100.0, method="Exact"))
            matched_left.add(left.source_key)
            matched_right.add(right.source_key)

    remaining_left = {
        record.source_key: record
        for record in left_records
        if record.source_key not in matched_left
    }
    remaining_right = {
        record.source_key: record
        for record in right_records
        if record.source_key not in matched_right
    }

    while remaining_left and remaining_right:
        left_candidates: dict[
            tuple[int, str, int], list[tuple[float, tuple[int, str, int]]]
        ] = defaultdict(list)
        right_candidates: dict[
            tuple[int, str, int], list[tuple[float, tuple[int, str, int]]]
        ] = defaultdict(list)
        for left_key, left in remaining_left.items():
            for right_key, right in remaining_right.items():
                score = text_similarity(
                    left.student_strict, right.student_strict, kind="student"
                )
                left_candidates[left_key].append((score, right_key))
                right_candidates[right_key].append((score, left_key))

        left_best: dict[
            tuple[int, str, int],
            tuple[float, tuple[int, str, int], float],
        ] = {}
        right_best: dict[
            tuple[int, str, int],
            tuple[float, tuple[int, str, int], float],
        ] = {}
        for key, candidates in left_candidates.items():
            candidates.sort(key=lambda item: (-item[0], item[1]))
            left_best[key] = (
                candidates[0][0],
                candidates[0][1],
                candidates[1][0] if len(candidates) > 1 else 0.0,
            )
        for key, candidates in right_candidates.items():
            candidates.sort(key=lambda item: (-item[0], item[1]))
            right_best[key] = (
                candidates[0][0],
                candidates[0][1],
                candidates[1][0] if len(candidates) > 1 else 0.0,
            )

        proposals: list[
            tuple[
                float,
                tuple[int, str, int],
                tuple[int, str, int],
                float,
            ]
        ] = []
        for left_key, (score, right_key, left_second) in left_best.items():
            right_score, right_choice, right_second = right_best[right_key]
            if right_choice != left_key:
                continue
            if score < threshold or right_score < threshold:
                continue
            if score - left_second < margin or right_score - right_second < margin:
                continue
            proposals.append(
                (score, left_key, right_key, max(left_second, right_second))
            )

        if not proposals:
            break

        score, left_key, right_key, second_best = max(
            proposals, key=lambda item: (item[0], item[1], item[2])
        )
        left = remaining_left.pop(left_key)
        right = remaining_right.pop(right_key)
        method = (
            "Normalized"
            if left.student_loose == right.student_loose
            else "Fuzzy"
        )
        matches.append(PairMatch(left=left, right=right, score=score, method=method))
        district_id, school_id, section = left.tuple_key
        issues.append(
            ReviewIssue(
                status="Auto-accepted",
                issue_type=f"{method} student match",
                district=district_display[district_id],
                school=school_display[school_id],
                section=section,
                period=f"{PERIOD_LABELS[left.period]} vs {PERIOD_LABELS[right.period]}",
                sheet=f"{left.sheet_name} / {right.sheet_name}",
                source_row=f"{left.row_number} / {right.row_number}",
                original_value=left.student_raw,
                candidate_value=right.student_raw,
                similarity_score=score,
                second_best_score=second_best,
                reason=(
                    f"Mutually best one-to-one candidates above threshold {threshold:.1f} "
                    f"with margin at least {margin:.1f}."
                ),
            )
        )

    best_for_left: dict[tuple[int, str, int], CandidateInfo] = {}
    best_for_right: dict[tuple[int, str, int], CandidateInfo] = {}
    left_rankings: dict[
        tuple[int, str, int], list[tuple[float, tuple[int, str, int]]]
    ] = defaultdict(list)
    right_rankings: dict[
        tuple[int, str, int], list[tuple[float, tuple[int, str, int]]]
    ] = defaultdict(list)
    for left_key, left in remaining_left.items():
        for right_key, right in remaining_right.items():
            score = text_similarity(left.student_strict, right.student_strict, kind="student")
            left_rankings[left_key].append((score, right_key))
            right_rankings[right_key].append((score, left_key))

    for left_key, candidates in left_rankings.items():
        candidates.sort(key=lambda item: (-item[0], item[1]))
        score, right_key = candidates[0]
        best_for_left[left_key] = CandidateInfo(remaining_right[right_key], score)
    for right_key, candidates in right_rankings.items():
        candidates.sort(key=lambda item: (-item[0], item[1]))
        score, left_key = candidates[0]
        best_for_right[right_key] = CandidateInfo(remaining_left[left_key], score)

    review_floor = max(0.0, threshold - 10.0)
    seen_pairs: set[
        tuple[tuple[int, str, int], tuple[int, str, int]]
    ] = set()
    for left_key, candidate in best_for_left.items():
        if candidate.score < review_floor:
            continue
        right_key = candidate.record.source_key
        pair_key = (left_key, right_key)
        if pair_key in seen_pairs:
            continue
        seen_pairs.add(pair_key)
        left = remaining_left[left_key]
        right = remaining_right[right_key]
        left_scores = sorted(
            (score for score, _ in left_rankings[left_key]), reverse=True
        )
        right_scores = sorted(
            (score for score, _ in right_rankings[right_key]), reverse=True
        )
        second_best = max(
            left_scores[1] if len(left_scores) > 1 else 0.0,
            right_scores[1] if len(right_scores) > 1 else 0.0,
        )
        if candidate.score < threshold:
            reason = f"Best score is below automatic threshold {threshold:.1f}."
        elif candidate.score - second_best < margin:
            reason = f"Competing candidates are within margin {margin:.1f}."
        else:
            reason = "The candidate relationship is not mutually best."
        district_id, school_id, section = left.tuple_key
        issues.append(
            ReviewIssue(
                status="Needs review",
                issue_type="Ambiguous student match",
                district=district_display[district_id],
                school=school_display[school_id],
                section=section,
                period=f"{PERIOD_LABELS[left.period]} vs {PERIOD_LABELS[right.period]}",
                sheet=f"{left.sheet_name} / {right.sheet_name}",
                source_row=f"{left.row_number} / {right.row_number}",
                original_value=left.student_raw,
                candidate_value=right.student_raw,
                similarity_score=candidate.score,
                second_best_score=second_best,
                reason=reason + " The rows were kept unmatched.",
            )
        )

    return PairwiseResult(
        left_period=left_period,
        right_period=right_period,
        matches=sorted(matches, key=lambda match: (match.left.row_number, match.right.row_number)),
        unmatched_left=sorted(remaining_left.values(), key=lambda item: item.row_number),
        unmatched_right=sorted(remaining_right.values(), key=lambda item: item.row_number),
        best_for_left=best_for_left,
        best_for_right=best_for_right,
    )


def empty_pairwise_result(left_period: int, right_period: int) -> PairwiseResult:
    return PairwiseResult(
        left_period=left_period,
        right_period=right_period,
        matches=[],
        unmatched_left=[],
        unmatched_right=[],
        best_for_left={},
        best_for_right={},
    )


def compare_periods(
    left_records: Sequence[SourceRow],
    right_records: Sequence[SourceRow],
    *,
    left_period: int,
    right_period: int,
    threshold: float,
    margin: float,
    district_display: dict[str, str],
    school_display: dict[str, str],
    issues: list[ReviewIssue],
) -> PairwiseResult:
    if not left_records and not right_records:
        return empty_pairwise_result(left_period, right_period)
    result = match_student_lists(
        left_records,
        right_records,
        threshold=threshold,
        margin=margin,
        district_display=district_display,
        school_display=school_display,
        issues=issues,
    )
    result.left_period = left_period
    result.right_period = right_period
    return result


def create_tuple_comparisons(
    rows: Sequence[SourceRow],
    student_records: Sequence[SourceRow],
    *,
    student_threshold: float,
    student_margin: float,
    district_display: dict[str, str],
    school_display: dict[str, str],
    issues: list[ReviewIssue],
) -> list[TupleComparison]:
    section_periods: dict[tuple[str, str, str], set[int]] = defaultdict(set)
    records_by_tuple: dict[
        tuple[str, str, str], dict[int, list[SourceRow]]
    ] = defaultdict(lambda: defaultdict(list))
    for row in rows:
        section_periods[row.tuple_key].add(row.period)
    for record in student_records:
        records_by_tuple[record.tuple_key][record.period].append(record)

    comparisons: list[TupleComparison] = []
    for tuple_key in sorted(
        section_periods,
        key=lambda key: (
            normalize_loose(district_display[key[0]]),
            normalize_loose(school_display[key[1]]),
            key[2],
        ),
    ):
        period_records = {
            period: sorted(records_by_tuple[tuple_key].get(period, []), key=lambda r: r.row_number)
            for period in (1, 2, 3)
        }
        if 1 in section_periods[tuple_key]:
            comparison = TupleComparison(
                tuple_key=tuple_key,
                reference_mode="Sheet 1 baseline",
                section_periods=set(section_periods[tuple_key]),
                records_by_period=period_records,
            )
            comparison.sheet1_to_sheet2 = compare_periods(
                period_records[1],
                period_records[2],
                left_period=1,
                right_period=2,
                threshold=student_threshold,
                margin=student_margin,
                district_display=district_display,
                school_display=school_display,
                issues=issues,
            )
            comparison.sheet1_to_sheet3 = compare_periods(
                period_records[1],
                period_records[3],
                left_period=1,
                right_period=3,
                threshold=student_threshold,
                margin=student_margin,
                district_display=district_display,
                school_display=school_display,
                issues=issues,
            )
        else:
            comparison = TupleComparison(
                tuple_key=tuple_key,
                reference_mode="Sheets 2-3 comparison (no Sheet 1 baseline)",
                section_periods=set(section_periods[tuple_key]),
                records_by_period=period_records,
            )
            comparison.sheet2_to_sheet3 = compare_periods(
                period_records[2],
                period_records[3],
                left_period=2,
                right_period=3,
                threshold=student_threshold,
                margin=student_margin,
                district_display=district_display,
                school_display=school_display,
                issues=issues,
            )
        comparisons.append(comparison)
    return comparisons


def validate_tuple_comparisons(comparisons: Sequence[TupleComparison]) -> None:
    def validate_pair(
        result: PairwiseResult,
        expected_left: Sequence[SourceRow],
        expected_right: Sequence[SourceRow],
    ) -> None:
        used_left = [match.left.source_key for match in result.matches] + [
            record.source_key for record in result.unmatched_left
        ]
        used_right = [match.right.source_key for match in result.matches] + [
            record.source_key for record in result.unmatched_right
        ]
        expected_left_keys = [record.source_key for record in expected_left]
        expected_right_keys = [record.source_key for record in expected_right]
        if len(used_left) != len(set(used_left)) or set(used_left) != set(expected_left_keys):
            raise RuntimeError("Left-side pairwise reconciliation failed.")
        if len(used_right) != len(set(used_right)) or set(used_right) != set(expected_right_keys):
            raise RuntimeError("Right-side pairwise reconciliation failed.")
        if len(expected_left) != len(result.matches) + len(result.unmatched_left):
            raise RuntimeError("Left-side matched/missing equation failed.")
        if len(expected_right) != len(result.matches) + len(result.unmatched_right):
            raise RuntimeError("Right-side matched/extra equation failed.")

    for comparison in comparisons:
        records = comparison.records_by_period
        if comparison.reference_mode == "Sheet 1 baseline":
            if comparison.sheet1_to_sheet2 is None or comparison.sheet1_to_sheet3 is None:
                raise RuntimeError("A baseline tuple lacks required pairwise comparisons.")
            validate_pair(comparison.sheet1_to_sheet2, records[1], records[2])
            validate_pair(comparison.sheet1_to_sheet3, records[1], records[3])
        else:
            if comparison.sheet2_to_sheet3 is None:
                raise RuntimeError("A no-baseline tuple lacks its Sheet 2-3 comparison.")
            validate_pair(comparison.sheet2_to_sheet3, records[2], records[3])


def write_header(worksheet, headers: Sequence[str]) -> None:
    worksheet.append(list(headers))
    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = HEADER_BORDER
    worksheet.row_dimensions[1].height = 34
    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False


def style_body(worksheet, numeric_columns: set[int] | None = None) -> None:
    numeric_columns = numeric_columns or set()
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.column in numeric_columns and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"


def fit_columns(
    worksheet,
    *,
    minimum: int = 10,
    maximum: int = 46,
    overrides: dict[int, int] | None = None,
) -> None:
    overrides = overrides or {}
    for column_index in range(1, worksheet.max_column + 1):
        if column_index in overrides:
            worksheet.column_dimensions[
                worksheet.cell(row=1, column=column_index).column_letter
            ].width = overrides[column_index]
            continue
        longest = 0
        for cell in worksheet.iter_cols(
            min_col=column_index,
            max_col=column_index,
            min_row=1,
            max_row=min(worksheet.max_row, 5000),
        ):
            for item in cell:
                if item.value is None:
                    continue
                lines = str(item.value).splitlines() or [""]
                longest = max(longest, *(len(line) for line in lines))
        width = min(max(longest + 2, minimum), maximum)
        worksheet.column_dimensions[
            worksheet.cell(row=1, column=column_index).column_letter
        ].width = width


def write_empty_message(worksheet, message: str, column_count: int) -> None:
    worksheet.cell(row=2, column=1, value=message)
    if column_count > 1:
        worksheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=column_count)
    cell = worksheet.cell(row=2, column=1)
    cell.fill = ALT_FILL
    cell.font = Font(italic=True, color="44546A")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = CELL_BORDER


def build_reconciliation_workbook(
    *,
    comparisons: Sequence[TupleComparison],
    duplicates: Sequence[DuplicateGroup],
    issues: Sequence[ReviewIssue],
    district_display: dict[str, str],
    school_display: dict[str, str],
) -> Workbook:
    """Build the reference-based reconciliation report."""

    workbook = Workbook()
    workbook.properties.title = "Student Roster Reconciliation"
    workbook.properties.subject = "Sheet 1 baseline validation and Sheet 2-3 fallback comparison"
    workbook.properties.creator = "Student roster reconciliation script"

    overview_sheet = workbook.active
    overview_sheet.title = "Overview"
    section_sheet = workbook.create_sheet("Section Reconciliation")
    exceptions_sheet = workbook.create_sheet("Student Exceptions")
    unmatched_sheet = workbook.create_sheet("Unmatched Later Records")
    duplicates_sheet = workbook.create_sheet("Duplicate Names")
    review_sheet = workbook.create_sheet("Match Review")

    metrics: dict[str, int] = defaultdict(int)
    student_exception_rows: list[list[object]] = []
    unmatched_rows: list[list[object]] = []

    section_headers = (
        "District",
        "School",
        "Section",
        "Comparison Basis",
        "Sheet 1 Section Available",
        "Sheet 2 Section Available",
        "Sheet 3 Section Available",
        "Sheet 1 Student Count",
        "Sheet 2 Student Count",
        "Sheet 3 Student Count",
        "Sheet 1-2 Matched",
        "Baseline Students Missing from Sheet 2",
        "Unmatched/Extra Sheet 2 Rows",
        "Sheet 2 Exact Baseline Roster",
        "Sheet 1-3 Matched",
        "Baseline Students Missing from Sheet 3",
        "Unmatched/Extra Sheet 3 Rows",
        "Sheet 3 Exact Baseline Roster",
        "Baseline Students Matched in Both Later Sheets",
        "Baseline Students Missing in At Least One Later Sheet",
        "Sheet 2-3 Matched (No Baseline)",
        "Only in Sheet 2 (No Baseline)",
        "Only in Sheet 3 (No Baseline)",
        "Sheet 2-3 Exact Roster Match (No Baseline)",
    )
    write_header(section_sheet, section_headers)

    for comparison in comparisons:
        district_id, school_id, section = comparison.tuple_key
        district = district_display[district_id]
        school = school_display[school_id]
        records = comparison.records_by_period
        periods = comparison.section_periods
        counts = {period: len(records[period]) for period in (1, 2, 3)}
        metrics["Total district-school-section combinations"] += 1

        if comparison.reference_mode == "Sheet 1 baseline":
            metrics["Sections with Sheet 1 baseline"] += 1
            metrics["Sheet 1 baseline student rows"] += counts[1]
            if 2 not in periods:
                metrics["Baseline sections absent from Sheet 2"] += 1
            if 3 not in periods:
                metrics["Baseline sections absent from Sheet 3"] += 1

            result_12 = comparison.sheet1_to_sheet2
            result_13 = comparison.sheet1_to_sheet3
            if result_12 is None or result_13 is None:
                raise RuntimeError("Missing baseline comparison result.")

            matched_12 = len(result_12.matches)
            missing_2 = len(result_12.unmatched_left)
            extra_2 = len(result_12.unmatched_right)
            matched_13 = len(result_13.matches)
            missing_3 = len(result_13.unmatched_left)
            extra_3 = len(result_13.unmatched_right)
            matched_both_keys = set(result_12.right_by_left) & set(result_13.right_by_left)
            matched_both = len(matched_both_keys)
            missing_either = counts[1] - matched_both
            exact_2 = 2 in periods and missing_2 == 0 and extra_2 == 0
            exact_3 = 3 in periods and missing_3 == 0 and extra_3 == 0

            metrics["Baseline students matched in Sheet 2"] += matched_12
            metrics["Baseline students missing from Sheet 2"] += missing_2
            metrics["Unmatched/extra Sheet 2 rows in baseline sections"] += extra_2
            metrics["Baseline students matched in Sheet 3"] += matched_13
            metrics["Baseline students missing from Sheet 3"] += missing_3
            metrics["Unmatched/extra Sheet 3 rows in baseline sections"] += extra_3
            metrics["Baseline students matched in both later sheets"] += matched_both
            metrics["Baseline students missing in at least one later sheet"] += missing_either

            section_sheet.append(
                [
                    district,
                    school,
                    section,
                    comparison.reference_mode,
                    "Yes",
                    "Yes" if 2 in periods else "No",
                    "Yes" if 3 in periods else "No",
                    counts[1],
                    counts[2],
                    counts[3],
                    matched_12,
                    missing_2,
                    extra_2,
                    "Yes" if exact_2 else "No",
                    matched_13,
                    missing_3,
                    extra_3,
                    "Yes" if exact_3 else "No",
                    matched_both,
                    missing_either,
                    "",
                    "",
                    "",
                    "",
                ]
            )

            match_12_by_left = result_12.right_by_left
            match_13_by_left = result_13.right_by_left
            for baseline_record in records[1]:
                match_12 = match_12_by_left.get(baseline_record.source_key)
                match_13 = match_13_by_left.get(baseline_record.source_key)
                if match_12 is not None and match_13 is not None:
                    continue
                if match_12 is None and match_13 is None:
                    issue_type = "Missing from Sheets 2 and 3"
                elif match_12 is None:
                    issue_type = "Missing from Sheet 2"
                else:
                    issue_type = "Missing from Sheet 3"
                student_exception_rows.append(
                    [
                        district,
                        school,
                        section,
                        baseline_record.student_raw,
                        "Available" if match_12 is not None else "Missing",
                        match_12.right.student_raw if match_12 else "",
                        match_12.method if match_12 else "",
                        round(match_12.score, 2) if match_12 else "",
                        "Available" if match_13 is not None else "Missing",
                        match_13.right.student_raw if match_13 else "",
                        match_13.method if match_13 else "",
                        round(match_13.score, 2) if match_13 else "",
                        issue_type,
                        baseline_record.sheet_name,
                        baseline_record.row_number,
                    ]
                )

            for period, result in ((2, result_12), (3, result_13)):
                for record in result.unmatched_right:
                    candidate = result.best_for_right.get(record.source_key)
                    unmatched_rows.append(
                        [
                            PERIOD_LABELS[period],
                            district,
                            school,
                            section,
                            record.student_raw,
                            "Sheet 1 baseline available",
                            f"Unmatched/extra relative to Sheet 1",
                            candidate.record.student_raw if candidate else "",
                            round(candidate.score, 2) if candidate else "",
                            record.sheet_name,
                            record.row_number,
                        ]
                    )
        else:
            metrics["Sections without Sheet 1 baseline"] += 1
            metrics["No-baseline Sheet 2 student rows"] += counts[2]
            metrics["No-baseline Sheet 3 student rows"] += counts[3]
            result_23 = comparison.sheet2_to_sheet3
            if result_23 is None:
                raise RuntimeError("Missing Sheet 2-3 no-baseline comparison result.")
            matched_23 = len(result_23.matches)
            only_2 = len(result_23.unmatched_left)
            only_3 = len(result_23.unmatched_right)
            exact_23 = 2 in periods and 3 in periods and only_2 == 0 and only_3 == 0
            metrics["No-baseline students matched between Sheets 2 and 3"] += matched_23
            metrics["No-baseline rows only in Sheet 2"] += only_2
            metrics["No-baseline rows only in Sheet 3"] += only_3

            section_sheet.append(
                [
                    district,
                    school,
                    section,
                    comparison.reference_mode,
                    "No",
                    "Yes" if 2 in periods else "No",
                    "Yes" if 3 in periods else "No",
                    counts[1],
                    counts[2],
                    counts[3],
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    matched_23,
                    only_2,
                    only_3,
                    "Yes" if exact_23 else "No",
                ]
            )

            for record in result_23.unmatched_left:
                candidate = result_23.best_for_left.get(record.source_key)
                unmatched_rows.append(
                    [
                        PERIOD_LABELS[2],
                        district,
                        school,
                        section,
                        record.student_raw,
                        "No Sheet 1 baseline; compared with Sheet 3",
                        "Present only/unmatched in Sheet 2",
                        candidate.record.student_raw if candidate else "",
                        round(candidate.score, 2) if candidate else "",
                        record.sheet_name,
                        record.row_number,
                    ]
                )
            for record in result_23.unmatched_right:
                candidate = result_23.best_for_right.get(record.source_key)
                unmatched_rows.append(
                    [
                        PERIOD_LABELS[3],
                        district,
                        school,
                        section,
                        record.student_raw,
                        "No Sheet 1 baseline; compared with Sheet 2",
                        "Present only/unmatched in Sheet 3",
                        candidate.record.student_raw if candidate else "",
                        round(candidate.score, 2) if candidate else "",
                        record.sheet_name,
                        record.row_number,
                    ]
                )

    metrics["Student exception rows"] = len(student_exception_rows)
    metrics["Unmatched later-record rows"] = len(unmatched_rows)
    metrics["Duplicate-name groups"] = len(duplicates)
    metrics["Matches needing review"] = sum(
        issue.status == "Needs review" for issue in issues
    )
    metrics["Data-quality issues"] = sum(issue.status == "Data issue" for issue in issues)

    overview_headers = ("Metric", "Value", "Meaning")
    write_header(overview_sheet, overview_headers)
    overview_order = (
        ("Total district-school-section combinations", "Union of section combinations found in any source sheet."),
        ("Sections with Sheet 1 baseline", "Sections validated against the authoritative Sheet 1 roster."),
        ("Sections without Sheet 1 baseline", "Sections compared directly between Sheets 2 and 3."),
        ("Baseline sections absent from Sheet 2", "Sheet 1 sections with no corresponding Sheet 2 submission."),
        ("Baseline sections absent from Sheet 3", "Sheet 1 sections with no corresponding Sheet 3 submission."),
        ("Sheet 1 baseline student rows", "Authoritative student-row population in baseline sections."),
        ("Baseline students matched in Sheet 2", "Sheet 1 students reproduced in Sheet 2."),
        ("Baseline students missing from Sheet 2", "Sheet 1 students without a Sheet 2 match."),
        ("Unmatched/extra Sheet 2 rows in baseline sections", "Sheet 2 rows without a Sheet 1 match."),
        ("Baseline students matched in Sheet 3", "Sheet 1 students reproduced in Sheet 3."),
        ("Baseline students missing from Sheet 3", "Sheet 1 students without a Sheet 3 match."),
        ("Unmatched/extra Sheet 3 rows in baseline sections", "Sheet 3 rows without a Sheet 1 match."),
        ("Baseline students matched in both later sheets", "Sheet 1 students matched independently in both later sheets."),
        ("Baseline students missing in at least one later sheet", "Sheet 1 students requiring review in Sheet 2, Sheet 3, or both."),
        ("No-baseline Sheet 2 student rows", "Sheet 2 student rows in sections unavailable from Sheet 1."),
        ("No-baseline Sheet 3 student rows", "Sheet 3 student rows in sections unavailable from Sheet 1."),
        ("No-baseline students matched between Sheets 2 and 3", "Direct matches where Sheet 1 has no section baseline."),
        ("No-baseline rows only in Sheet 2", "Sheet 2 rows without a Sheet 3 match in no-baseline sections."),
        ("No-baseline rows only in Sheet 3", "Sheet 3 rows without a Sheet 2 match in no-baseline sections."),
        ("Student exception rows", "Baseline students missing from at least one later sheet."),
        ("Unmatched later-record rows", "Later-sheet records not matched to the applicable reference roster."),
        ("Duplicate-name groups", "Repeated normalized names inside one period and strict location tuple."),
        ("Matches needing review", "Possible fuzzy matches intentionally left unresolved."),
        ("Data-quality issues", "Invalid or incomplete source rows recorded for review."),
    )
    for metric, meaning in overview_order:
        overview_sheet.append([metric, metrics[metric], meaning])
    style_body(overview_sheet, numeric_columns={2})
    fit_columns(overview_sheet, overrides={1: 48, 2: 16, 3: 70})

    for row_number in range(2, section_sheet.max_row + 1):
        for column in (5, 6, 7, 14, 18, 24):
            cell = section_sheet.cell(row_number, column)
            if cell.value == "Yes":
                cell.fill = GREEN_FILL
                cell.font = GREEN_FONT
            elif cell.value == "No":
                cell.fill = RED_FILL
                cell.font = RED_FONT
            if cell.value:
                cell.alignment = Alignment(horizontal="center", vertical="center")
    style_body(
        section_sheet,
        numeric_columns=set(range(8, 14)) | set(range(15, 24)),
    )
    fit_columns(
        section_sheet,
        maximum=34,
        overrides={1: 22, 2: 36, 3: 10, 4: 36, 5: 18, 6: 18, 7: 18},
    )

    exception_headers = (
        "District",
        "School",
        "Section",
        "Sheet 1 Baseline Student Name",
        "Sheet 2 Status",
        "Sheet 2 Matched Name",
        "Sheet 2 Match Method",
        "Sheet 2 Match Score",
        "Sheet 3 Status",
        "Sheet 3 Matched Name",
        "Sheet 3 Match Method",
        "Sheet 3 Match Score",
        "Issue Type",
        "Baseline Source Sheet",
        "Baseline Source Row",
    )
    write_header(exceptions_sheet, exception_headers)
    student_exception_rows.sort(
        key=lambda row: (
            normalize_loose(row[0]),
            normalize_loose(row[1]),
            row[2],
            normalize_loose(row[3]),
            row[14],
        )
    )
    for output_row in student_exception_rows:
        exceptions_sheet.append(output_row)
        row_number = exceptions_sheet.max_row
        for column in (5, 9):
            cell = exceptions_sheet.cell(row_number, column)
            present = cell.value == "Available"
            cell.fill = GREEN_FILL if present else RED_FILL
            cell.font = GREEN_FONT if present else RED_FONT
            cell.alignment = Alignment(horizontal="center")
    if not student_exception_rows:
        write_empty_message(
            exceptions_sheet,
            "No baseline students are missing from Sheets 2 or 3.",
            len(exception_headers),
        )
    style_body(exceptions_sheet, numeric_columns={8, 12, 15})
    fit_columns(
        exceptions_sheet,
        maximum=40,
        overrides={1: 22, 2: 36, 3: 10, 4: 30, 5: 16, 6: 30, 9: 16, 10: 30, 13: 30},
    )

    unmatched_headers = (
        "Period",
        "District",
        "School",
        "Section",
        "Unmatched Student Name",
        "Comparison Basis",
        "Issue Type",
        "Best Candidate",
        "Best Candidate Score",
        "Source Sheet",
        "Source Row",
    )
    write_header(unmatched_sheet, unmatched_headers)
    unmatched_rows.sort(
        key=lambda row: (
            normalize_loose(row[1]),
            normalize_loose(row[2]),
            row[3],
            row[0],
            normalize_loose(row[4]),
            row[10],
        )
    )
    for output_row in unmatched_rows:
        unmatched_sheet.append(output_row)
    if not unmatched_rows:
        write_empty_message(
            unmatched_sheet,
            "No unmatched later-sheet student records were detected.",
            len(unmatched_headers),
        )
    style_body(unmatched_sheet, numeric_columns={9, 11})
    fit_columns(
        unmatched_sheet,
        maximum=44,
        overrides={1: 24, 2: 22, 3: 36, 4: 10, 5: 30, 6: 38, 7: 36, 8: 30},
    )

    duplicate_headers = (
        "Period",
        "District",
        "School",
        "Section",
        "Student Name",
        "Occurrence Count",
        "Original Spellings",
        "Source Row Numbers",
        "Explanation",
    )
    write_header(duplicates_sheet, duplicate_headers)
    sorted_duplicates = sorted(
        duplicates,
        key=lambda group: (
            normalize_loose(district_display[group.tuple_key[0]]),
            normalize_loose(school_display[group.tuple_key[1]]),
            group.tuple_key[2],
            group.period,
            group.normalized_name,
        ),
    )
    for group in sorted_duplicates:
        district_id, school_id, section = group.tuple_key
        duplicates_sheet.append(
            [
                PERIOD_LABELS[group.period],
                district_display[district_id],
                school_display[school_id],
                section,
                group.records[0].student_raw,
                len(group.records),
                " | ".join(record.student_raw for record in group.records),
                ", ".join(str(record.row_number) for record in group.records),
                "Rows remain separate because later sheets do not contain a common student ID.",
            ]
        )
    if not sorted_duplicates:
        write_empty_message(
            duplicates_sheet, "No duplicate-name groups detected.", len(duplicate_headers)
        )
    style_body(duplicates_sheet, numeric_columns={6})
    fit_columns(
        duplicates_sheet,
        overrides={1: 24, 2: 22, 3: 36, 4: 10, 5: 28, 6: 16, 7: 36, 8: 20, 9: 52},
    )

    review_headers = (
        "Status",
        "Issue Type",
        "District",
        "School",
        "Section",
        "Period",
        "Sheet",
        "Source Row",
        "Original Value",
        "Candidate Value",
        "Similarity Score",
        "Second-Best Score",
        "Reason",
    )
    write_header(review_sheet, review_headers)
    status_order = {"Needs review": 0, "Data issue": 1, "Auto-accepted": 2}
    sorted_issues = sorted(
        issues,
        key=lambda issue: (
            status_order.get(issue.status, 9),
            normalize_loose(issue.district),
            normalize_loose(issue.school),
            issue.section,
            issue.issue_type,
            str(issue.source_row),
        ),
    )
    for issue in sorted_issues:
        review_sheet.append(issue.as_excel_row())
        status_cell = review_sheet.cell(review_sheet.max_row, 1)
        if issue.status == "Auto-accepted":
            status_cell.fill = GREEN_FILL
            status_cell.font = GREEN_FONT
        elif issue.status == "Needs review":
            status_cell.fill = PatternFill("solid", fgColor="FFF2CC")
            status_cell.font = Font(color="9C6500", bold=True)
        else:
            status_cell.fill = RED_FILL
            status_cell.font = RED_FONT
    if not sorted_issues:
        write_empty_message(review_sheet, "No review issues detected.", len(review_headers))
    style_body(review_sheet, numeric_columns={8, 11, 12})
    fit_columns(
        review_sheet,
        overrides={1: 16, 2: 26, 3: 22, 4: 34, 5: 10, 6: 34, 7: 28, 8: 16, 9: 42, 10: 42, 11: 16, 12: 18, 13: 54},
    )

    for worksheet in workbook.worksheets:
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0
        worksheet.page_layout_view = False

    return workbook


def verify_output_workbook(path: Path) -> None:
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        if tuple(workbook.sheetnames) != OUTPUT_SHEETS:
            raise RuntimeError(
                f"Output sheet verification failed: found {workbook.sheetnames!r}."
            )
        expected_first_headers = {
            "Overview": "Metric",
            "Section Reconciliation": "District",
            "Student Exceptions": "District",
            "Unmatched Later Records": "Period",
            "Duplicate Names": "Period",
            "Match Review": "Status",
        }
        for sheet_name, expected_header in expected_first_headers.items():
            value = workbook[sheet_name].cell(row=1, column=1).value
            if value != expected_header:
                raise RuntimeError(
                    f"Output header verification failed in {sheet_name!r}: {value!r}."
                )
    finally:
        workbook.close()


def validate_score(value: float, name: str) -> None:
    if not 0.0 <= value <= 100.0:
        raise ValueError(f"{name} must be between 0 and 100.")


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Reconcile two third-grade rosters and one fourth-grade roster using "
            "Sheet 1 as the baseline when available, without modifying the source."
        )
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        help=(
            "Path to the source .xlsx/.xlsm workbook. If omitted, the program "
            "asks you to enter it."
        ),
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help=(
            "Path for the new .xlsx report. Defaults to "
            "<input_stem>_student_tracking.xlsx beside the input."
        ),
    )
    parser.add_argument("--sheet-1", help="First-semester third-grade worksheet name")
    parser.add_argument("--sheet-2", help="Second-semester third-grade worksheet name")
    parser.add_argument("--sheet-3", help="Fourth-grade worksheet name")
    parser.add_argument(
        "--header-row",
        type=int,
        default=1,
        help="One-based header row number; data begins on the following row (default: 1)",
    )
    parser.add_argument("--district-threshold", type=float, default=90.0)
    parser.add_argument("--district-margin", type=float, default=5.0)
    parser.add_argument("--school-threshold", type=float, default=88.0)
    parser.add_argument("--school-margin", type=float, default=5.0)
    parser.add_argument("--student-threshold", type=float, default=90.0)
    parser.add_argument("--student-margin", type=float, default=5.0)
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Allow replacement of an existing output file; never affects the source file",
    )
    return parser.parse_args(argv)


def prompt_for_input_path() -> Path:
    """Ask for the source path when it was not supplied on the command line."""

    try:
        entered_path = input(
            "Enter or paste the full path of the input Excel workbook: "
        ).strip()
    except (EOFError, KeyboardInterrupt) as exc:
        raise SystemExit("\nNo input workbook path was provided.") from exc

    # Dragging a file into a Windows terminal commonly surrounds its path with
    # matching single or double quotation marks.
    if (
        len(entered_path) >= 2
        and entered_path[0] == entered_path[-1]
        and entered_path[0] in {"'", '"'}
    ):
        entered_path = entered_path[1:-1].strip()

    if not entered_path:
        raise SystemExit("No input workbook path was provided.")
    return Path(entered_path)


def run(args: argparse.Namespace) -> Path:
    input_path = args.input.expanduser().resolve()
    if not input_path.is_file():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")
    if input_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("The input must be an .xlsx or .xlsm workbook.")
    if args.header_row < 1:
        raise ValueError("--header-row must be at least 1.")

    output_path = (
        args.output.expanduser().resolve()
        if args.output
        else input_path.with_name(f"{input_path.stem}_student_tracking.xlsx")
    )
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("The output path must end with .xlsx.")
    if input_path == output_path:
        raise ValueError("The output path must be different from the source workbook path.")
    if output_path.exists() and not args.overwrite:
        raise FileExistsError(
            f"Output already exists: {output_path}. Use --overwrite to replace it."
        )

    for value, name in (
        (args.district_threshold, "--district-threshold"),
        (args.district_margin, "--district-margin"),
        (args.school_threshold, "--school-threshold"),
        (args.school_margin, "--school-margin"),
        (args.student_threshold, "--student-threshold"),
        (args.student_margin, "--student-margin"),
    ):
        validate_score(value, name)

    source_hash_before = sha256_file(input_path)
    issues: list[ReviewIssue] = []
    rows, selected_sheets = read_source_rows(
        input_path,
        (args.sheet_1, args.sheet_2, args.sheet_3),
        args.header_row,
        issues,
    )
    if sha256_file(input_path) != source_hash_before:
        raise RuntimeError("The source workbook changed while it was being read; aborting.")
    if not rows:
        raise ValueError("No valid district-school-section rows were found.")

    district_display, school_display = resolve_institutions(
        rows,
        district_threshold=args.district_threshold,
        district_margin=args.district_margin,
        school_threshold=args.school_threshold,
        school_margin=args.school_margin,
        issues=issues,
    )
    student_records = [row for row in rows if row.has_student]
    duplicates = detect_duplicate_names(student_records)
    comparisons = create_tuple_comparisons(
        rows,
        student_records,
        student_threshold=args.student_threshold,
        student_margin=args.student_margin,
        district_display=district_display,
        school_display=school_display,
        issues=issues,
    )
    validate_tuple_comparisons(comparisons)

    output_workbook = build_reconciliation_workbook(
        comparisons=comparisons,
        duplicates=duplicates,
        issues=issues,
        district_display=district_display,
        school_display=school_display,
    )

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
        output_workbook.save(temporary_path)
        output_workbook.close()
        verify_output_workbook(temporary_path)

        if sha256_file(input_path) != source_hash_before:
            raise RuntimeError(
                "The source workbook changed during processing; output was not committed."
            )
        if output_path.exists() and not args.overwrite:
            raise FileExistsError(
                f"Output appeared during processing and was not overwritten: {output_path}"
            )
        os.replace(temporary_path, output_path)
        temporary_path = None
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()

    source_hash_after = sha256_file(input_path)
    if source_hash_after != source_hash_before:
        raise RuntimeError(
            "Source hash changed unexpectedly after output creation. The report exists, "
            "but the source should be investigated."
        )

    baseline_sections = sum(
        comparison.reference_mode == "Sheet 1 baseline"
        for comparison in comparisons
    )
    no_baseline_sections = len(comparisons) - baseline_sections
    baseline_exception_count = 0
    unmatched_later_count = 0
    for comparison in comparisons:
        if comparison.reference_mode == "Sheet 1 baseline":
            result_12 = comparison.sheet1_to_sheet2
            result_13 = comparison.sheet1_to_sheet3
            if result_12 is None or result_13 is None:
                raise RuntimeError("Missing baseline comparison result during reporting.")
            matched_both = set(result_12.right_by_left) & set(result_13.right_by_left)
            baseline_exception_count += len(comparison.records_by_period[1]) - len(matched_both)
            unmatched_later_count += len(result_12.unmatched_right) + len(result_13.unmatched_right)
        else:
            result_23 = comparison.sheet2_to_sheet3
            if result_23 is None:
                raise RuntimeError("Missing no-baseline comparison result during reporting.")
            unmatched_later_count += len(result_23.unmatched_left) + len(result_23.unmatched_right)
    needs_review_count = sum(issue.status == "Needs review" for issue in issues)
    data_issue_count = sum(issue.status == "Data issue" for issue in issues)
    print(f"Created: {output_path}")
    print(f"Sheets read: {', '.join(selected_sheets)}")
    print(f"Valid student rows: {len(student_records):,}")
    print(f"Sections with Sheet 1 baseline: {baseline_sections:,}")
    print(f"Sections compared only between Sheets 2 and 3: {no_baseline_sections:,}")
    print(f"Baseline students missing from at least one later sheet: {baseline_exception_count:,}")
    print(f"Unmatched later-sheet records: {unmatched_later_count:,}")
    print(f"Duplicate-name groups: {len(duplicates):,}")
    print(f"Matches needing review: {needs_review_count:,}")
    print(f"Data issues: {data_issue_count:,}")
    print(f"Source SHA-256 unchanged: {source_hash_after}")
    return output_path


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    if args.input is None:
        args.input = prompt_for_input_path()
    try:
        run(args)
    except (FileNotFoundError, FileExistsError, ValueError, RuntimeError) as exc:
        raise SystemExit(f"Error: {exc}") from exc
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
