"""Conservative, auditable additions to the existing name similarity score.

Inputs use student_match_report.normalize_for_match. This module never changes
that normalizer or its order-sensitive SequenceMatcher baseline. Scores are
heuristic ranking scores, not identity probabilities. Context and competition
must be checked by the caller before accepting a boosted pair.
"""
from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache

from student_match_report import BaseMatcher, BaseRecord, MatchOutcome, string_similarity

FULL_PART_SIMILARITY_THRESHOLD = 90.0


@dataclass(frozen=True)
class NameEvidence:
    pattern: str = ""
    rule_score: float | None = None
    initials: int = 0
    omitted: int = 0
    anchors: int = 0
    fuzzy_parts: int = 0
    full_parts_similarity: float | None = None
    reason: str = ""

    def adjusted(self, baseline: float) -> float:
        return max(baseline, self.rule_score) if self.rule_score is not None else baseline


@lru_cache(maxsize=100_000)
def name_evidence(left: str, right: str) -> NameEvidence:
    """Explain only positional initials and contiguous missing trailing tokens.

    For initials, the aligned full-token similarities must average at least
    90%; a short spelling difference may be offset by other matching parts.
    For omissions without initials, each retained full token must reach 90%.
    Token permutations, interior omissions, arbitrary substrings and fuzzy
    initial expansion are excluded. One-letter tokens match only the first
    letter of their aligned full token.
    Rules are symmetric: either source may contain the shorter form.
    """
    if not left or not right or left == right:
        return NameEvidence()
    a, b = left.split(), right.split()
    if not all(token.isalpha() and token.isascii() for token in a + b):
        return NameEvidence()
    short, long = (a, b) if len(a) <= len(b) else (b, a)
    initials = 0
    anchors = []
    fuzzy_parts = 0
    full_part_scores = []
    for u, v in zip(short, long):
        if u == v:
            if len(u) > 1:
                anchors.append((u, v))
                full_part_scores.append(100.0)
        elif min(len(u), len(v)) == 1 and u[0] == v[0]:
            initials += 1
        elif len(u) > 1 and len(v) > 1:
            anchors.append((u, v))
            full_part_scores.append(string_similarity(u, v))
            fuzzy_parts += 1
        else:
            return NameEvidence()
    anchor_count = min(len({u for u, _ in anchors}),
                       len({v for _, v in anchors}))
    full_parts_similarity = (sum(full_part_scores) / len(full_part_scores)
                             if full_part_scores else None)
    omitted = len(long) - len(short)
    if not initials and not omitted:
        return NameEvidence()
    if initials and (full_parts_similarity is None or
                     full_parts_similarity < FULL_PART_SIMILARITY_THRESHOLD):
        return NameEvidence()
    if not initials and any(score < FULL_PART_SIMILARITY_THRESHOLD
                            for score in full_part_scores):
        return NameEvidence()
    pattern = ("Initials and trailing omission" if initials and omitted else
               "Initial abbreviation" if initials else "Trailing omission")
    blocks = []
    if initials and len(short) < 3:
        blocks.append("Names with initials require at least three parts")
    if omitted:
        if len(short) < 3 or anchor_count < 3:
            blocks.append("Fewer than three distinct full-word anchors")
        if omitted > 2:
            blocks.append("More than two trailing name parts are missing")
        # An appended copy of earlier text is an extraction issue, not evidence
        # for the requested missing-tail rule. Do not silently clean it away.
        if any(token in long[:len(short)] for token in long[len(short):]):
            blocks.append("Omitted tail repeats earlier text; possible extraction duplication")
        if any(len(token) == 1 for token in long[len(short):]):
            blocks.append("Omitted tail contains unresolved initials")
    else:
        if anchor_count < 2:
            blocks.append("Fewer than two distinct full-word anchors")
        if len(a[-1]) <= 1 or len(b[-1]) <= 1:
            blocks.append("Last name token is not a full word on both sides")
    if initials > 2 or initials * 2 > len(short):
        blocks.append("Too many abbreviated parts")
    if any(len(u) == len(v) == 1 for u, v in zip(short, long)):
        blocks.append("An initial is unresolved on both sides")
    # Charge every abbreviated, omitted, or fuzzy full part; never award 100.
    floor = 100.0 - 3.0 * (initials + omitted + fuzzy_parts)
    if floor <= 90:
        blocks.append("Combined information loss exceeds the automatic-boost limit")
    return NameEvidence(
        pattern=pattern,
        rule_score=None if blocks else floor,
        initials=initials,
        omitted=omitted,
        anchors=anchor_count,
        fuzzy_parts=fuzzy_parts,
        full_parts_similarity=full_parts_similarity,
        reason="; ".join(blocks) if blocks else
               ("Aligned full name parts average at least 90%; initials agree" if initials else
                "Retained full name parts each meet 90% similarity"),
    )


def name_similarity(left: str, right: str) -> tuple[float, float, NameEvidence]:
    """Return the untouched baseline, proposed score and explanation."""
    baseline = string_similarity(left, right)
    evidence = name_evidence(left, right)
    return baseline, evidence.adjusted(baseline), evidence


@dataclass(frozen=True)
class PairAssessment:
    baseline: float
    adjusted: float
    evidence: NameEvidence
    decision: str
    reason: str
    baseline_status: str = ""
    baseline_candidates: int | None = None
    augmented_candidates: int | None = None
    rival_score: float | None = None
    rivals: tuple[tuple[BaseRecord, float, float, str], ...] = ()


def assess_existing_pair(
    matcher: BaseMatcher, district: str, school: str, section: str,
    source_name: str, erc_record: BaseRecord | None, erc_name: str,
) -> PairAssessment:
    """Boost an existing mapping only after checking the entire resolved school.

    No candidate is silently substituted. All rows with a detected pattern are
    reviewable, including rejected boosts and already-high scores. Unrelated
    pairs are returned unchanged. Duplicate roster rows remain competitors.
    """
    baseline, proposed, evidence = name_similarity(source_name, erc_name)
    if not evidence.pattern:
        return PairAssessment(baseline, baseline, evidence, "Unchanged", "")
    location = matcher.resolve_location(district, school)
    if (not location.accepted or location.location is None or erc_record is None
            or (location.location.district, location.location.school) !=
               (erc_record.district, erc_record.school)):
        return PairAssessment(baseline, baseline, evidence, "Review only",
                              "ERC mapping is missing/duplicated or source location does not resolve to it")
    candidates = matcher.records_by_location[(erc_record.district, erc_record.school)]
    scores = []
    for record in candidates:
        old, new, ev = name_similarity(source_name, record.student)
        scores.append((record, old, new, ev))
    old_candidates = [x for x in scores if x[1] > matcher.student_threshold]
    new_candidates = [x for x in scores if x[2] > matcher.student_threshold]
    rivals = [x for x in scores if x[0].row_number != erc_record.row_number]
    # Weaker prefix matches still disclose competition if >=2 full anchors
    # exist, even when that weaker pair cannot earn its own score boost.
    plausible_rivals = [x for x in rivals if x[2] > matcher.student_threshold or
                        (x[3].pattern and x[3].anchors >= 2 and
                         "duplication" not in x[3].reason)]
    rival_score = max((x[2] for x in rivals), default=None)
    baseline_status = ("Red" if not old_candidates else "Orange" if len(old_candidates)>1
                       else "Green" if section and old_candidates[0][0].section == section else "Yellow")
    details = dict(
        baseline_status=baseline_status,
        baseline_candidates=len(old_candidates),
        augmented_candidates=len(new_candidates),
        rival_score=rival_score,
        rivals=tuple((r, old, new, ev.pattern) for r,old,new,ev in
                     sorted(plausible_rivals, key=lambda x: (-x[2],x[0].row_number))),
    )
    reasons = []
    if evidence.rule_score is None:
        reasons.append(evidence.reason)
    if plausible_rivals:
        reasons.append("Competing ERC candidates in the same school")
    if rival_score is not None and proposed-rival_score < 5:
        reasons.append("Less than five points separate the proposed pair from its strongest rival")
    if not section:
        reasons.append("Source section is missing")
    if (len(new_candidates) != 1 or
            new_candidates[0][0].row_number != erc_record.row_number):
        reasons.append("The existing mapped ERC row is not the sole qualifying candidate")
    if reasons:
        return PairAssessment(baseline,baseline,evidence,"Review only","; ".join(reasons),**details)
    decision = "Green + review" if section == erc_record.section else "Yellow + review"
    return PairAssessment(baseline,proposed,evidence,decision,
                          "Unique school-wide candidate; " +
                          ("section agrees" if section == erc_record.section else "section differs"),
                          **details)


@dataclass(frozen=True)
class UnmatchedAssessment:
    baseline: MatchOutcome
    # All positional-pattern candidates, even those whose boosts are withheld.
    candidates: tuple[tuple[BaseRecord, PairAssessment], ...] = ()
    promoted: tuple[BaseRecord, PairAssessment] | None = None


def assess_unmatched_student(
    matcher: BaseMatcher, district: str, school: str, section: str, student: str,
) -> UnmatchedAssessment:
    """Search the full resolved school without requiring an accepted pairing.

    Membership in a historical mismatch sheet is established by the caller.
    This function never credits an unchanged baseline green/yellow/orange as a
    new structured-rule success. A failed location is not bypassed.
    """
    baseline = matcher.classify(district, school, section, student)
    if baseline.status != 'red_name' or baseline.location is None:
        return UnmatchedAssessment(baseline)
    location = baseline.location
    candidates = []
    for record in matcher.records_by_location[(location.district, location.school)]:
        evidence = name_evidence(student, record.student)
        if not evidence.pattern:
            continue
        assessment = assess_existing_pair(
            matcher, district, school, section, student, record, record.student,
        )
        candidates.append((record, assessment))
    candidates.sort(key=lambda item: (
        -item[1].evidence.adjusted(item[1].baseline),
        0 if item[0].section == section else 1,
        item[0].row_number,
    ))
    promoted = [item for item in candidates
                if item[1].decision in {'Green + review', 'Yellow + review'}
                and item[1].adjusted > matcher.student_threshold
                and item[1].adjusted > item[1].baseline]
    if len(promoted) > 1:
        raise AssertionError('More than one candidate passed school-wide uniqueness checks')
    return UnmatchedAssessment(baseline, tuple(candidates), promoted[0] if promoted else None)
