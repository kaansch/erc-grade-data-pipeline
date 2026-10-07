"""Synthetic identity edge cases; no private student records in this file."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts/05_erc_roster_matching"))
from student_match_report import BaseMatcher, BaseRecord, normalize_for_match as norm, string_similarity
from structured_name_match import assess_existing_pair, name_similarity


class StructuredNameTests(unittest.TestCase):
    def score(self, a, b):
        return name_similarity(norm(a), norm(b))

    def test_initial_position_punctuation_case_and_turkish_letters(self):
        full = "Kaan Samed Çil"
        for short in ("K. Samed Çil", "kaan s çil", "K.Samed Çil", "KAAN S. CIL"):
            with self.subTest(short=short):
                old, new, ev = self.score(short, full)
                self.assertEqual(new, 97)
                self.assertGreater(new, old)
                self.assertEqual(ev.pattern, "Initial abbreviation")
                self.assertEqual(self.score(full, short)[1], new)

    def test_trailing_one_or_two_full_parts(self):
        long = "Adam Basel Cemal Deniz Emre Faruk"
        for short, expected in (("Adam Basel Cemal Deniz Emre", 97),
                                ("Adam Basel Cemal Deniz", 94)):
            self.assertEqual(self.score(short, long)[1], expected)
            self.assertEqual(self.score(long, short)[1], expected)

    def test_three_retained_parts_allow_two_omitted_parts(self):
        old, new, ev = self.score("Adam Basel Cemal", "Adam Basel Cemal Deniz Emre")
        self.assertEqual(ev.omitted, 2)
        self.assertEqual(ev.anchors, 3)
        self.assertEqual(new, max(old, 94))

    def test_full_name_parts_may_meet_ninety_percent_fuzzy_threshold(self):
        for short, full in (
            ("K. Bartholomew Çil", "Kaan Bartholomaw Çil"),
            ("K. Samed Bartholomew", "Kaan Samed Bartholomaw"),
        ):
            with self.subTest(full=full):
                self.assertGreaterEqual(string_similarity(norm("Bartholomew"),
                                                          norm("Bartholomaw")), 90)
                old, new, ev = self.score(short, full)
                self.assertEqual(ev.fuzzy_parts, 1)
                self.assertEqual(ev.anchors, 2)
                self.assertEqual(ev.rule_score, 94)
                self.assertEqual(new, max(old, 94))
                self.assertEqual(self.score(full, short)[1], new)

    def test_fuzzy_anchors_remain_distinct_in_both_directions(self):
        left = "K. Bartholomew Bartholomew"
        right = "Kaan Bartholomew Bartholomaw"
        for a, b in ((left, right), (right, left)):
            with self.subTest(a=a):
                old, new, ev = self.score(a, b)
                self.assertEqual(ev.anchors, 1)
                self.assertIsNone(ev.rule_score)
                self.assertEqual(new, old)

    def test_initial_uses_average_of_remaining_full_name_parts(self):
        self.assertLess(string_similarity(norm("Samed"), norm("Samet")), 90)
        old, new, ev = self.score("K. Samed Çil", "Kaan Samet Çil")
        self.assertEqual(ev.pattern, "Initial abbreviation")
        self.assertEqual(ev.full_parts_similarity, 90)
        self.assertEqual(ev.fuzzy_parts, 1)
        self.assertEqual(ev.rule_score, 94)
        self.assertEqual(new, max(old, 94))
        self.assertEqual(self.score("Kaan Samet Çil", "K. Samed Çil")[1], new)

    def test_below_ninety_average_is_not_a_structural_candidate(self):
        old, new, ev = self.score("K. Samed Çil", "Kaan Samet Çelik")
        self.assertEqual(ev.pattern, "")
        self.assertEqual(new, old)

    def test_omission_without_initial_keeps_per_part_threshold(self):
        old, new, ev = self.score("Adam Samed Çil", "Adam Samet Çil Deniz")
        self.assertEqual(ev.pattern, "")
        self.assertEqual(new, old)

    def test_initials_require_three_parts(self):
        old, new, ev = self.score("K. Çil", "Kaan Çil")
        self.assertEqual(ev.pattern, "Initial abbreviation")
        self.assertIsNone(ev.rule_score)
        self.assertIn("at least three parts", ev.reason)
        self.assertEqual(new, old)

    def test_combined_initial_and_tail(self):
        old, new, ev = self.score("A. Basel Cemal Deniz", "Adam Basel Cemal Deniz Emre")
        self.assertEqual(new, 94)
        self.assertEqual(ev.anchors, 3)

    def test_no_boost_for_insufficient_or_conflicting_evidence(self):
        pairs = [
            ("K. Çil", "Kaan Çil"),
            ("K. S. Çil", "Kaan Samed Çil"),
            ("Kaan Samed C", "Kaan Samed Cil"),
            ("Kaan Samed Çil", "Kaan Selim Çil"),
            ("K. Samed Çil", "Ali Samed Çil"),
            ("Kaan Samed Çil", "Samed Kaan Çil"),
            ("Adam Basel", "Adam Basel Cemal Deniz"),
            ("Adam Cemal Deniz", "Adam Basel Cemal Deniz"),
            ("Adam Basel Cemal", "Adam Basel Cemal Deniz Emre Faruk"),
            ("Kaan Samed Çil", "Kaan Samed Çil Çil"),
            ("Adam B. Cemal Deniz", "Adam B. Cemal Deniz Emre"),
            ("A B C D", "A B C D E F"),
            ("1 Adam Basel", "1 Adam Basel Cemal"),
            ("", "Kaan Çil"), ("Kaan Çil", "Kaan Çil"),
        ]
        for a, b in pairs:
            with self.subTest(a=a, b=b):
                old, new, _ = self.score(a,b)
                self.assertEqual(new, old)
                self.assertEqual(old, string_similarity(norm(a),norm(b)))

    def test_long_name_baseline_never_decreases(self):
        full = "Alexander Bartholomew Christopher Demetrius Edward Franklin George Henry Ilya John"
        short = full.rsplit(" ",1)[0]
        old,new,_ = self.score(short,full)
        self.assertEqual(new, max(old,97))

    def test_every_eligible_score_is_below_exact(self):
        for a,b in [("K. Samed Cil","Kaan Samed Cil"),
                    ("Adam Basel Cemal","Adam Basel Cemal Deniz Emre")]:
            old,new,ev = self.score(a,b)
            self.assertGreater(new,90)
            self.assertLess(new,100)
            self.assertGreaterEqual(new,old)


class ContextTests(unittest.TestCase):
    def record(self, row, name, section="a", school="test school"):
        return BaseRecord(row,"test district",school,section,norm(name),
                          "Test District",school,section,name)

    def assess(self, records, source="K. Samed Cil", section="a", school="test school", target=0):
        matcher = BaseMatcher(records,student_threshold=90,location_threshold=95)
        selected = records[target]
        return assess_existing_pair(matcher,"test district",school,section,
                                    norm(source),selected,selected.student)

    def test_unique_candidate_promotes_and_preserves_baseline(self):
        r = self.record(2,"Kaan Samed Cil")
        result = self.assess([r])
        self.assertEqual(result.decision,"Green + review")
        self.assertEqual(result.baseline_status,"Red")
        self.assertEqual(result.adjusted,97)
        self.assertEqual(result.baseline,string_similarity(norm("K. Samed Cil"),r.student))

    def test_initial_collision_across_sections_is_not_green(self):
        rows = [self.record(2,"Kaan Samed Cil"),self.record(3,"Kerem Samed Cil",section="b")]
        result = self.assess(rows)
        self.assertEqual(result.decision,"Review only")
        self.assertEqual(result.adjusted,result.baseline)
        self.assertEqual(result.augmented_candidates,2)
        self.assertEqual(len(result.rivals),1)

    def test_fuzzy_initial_collision_remains_review_only(self):
        rows = [self.record(2,"Kaan Bartholomaw Cil"),
                self.record(3,"Kerem Bartholomaw Cil",section="b")]
        result = self.assess(rows,source="K Bartholomew Cil")
        self.assertEqual(result.decision,"Review only")
        self.assertEqual(result.adjusted,result.baseline)
        self.assertEqual(result.augmented_candidates,2)

    def test_short_spelling_difference_with_initial_can_promote(self):
        rows = [self.record(2,"Kaan Samet Cil")]
        result = self.assess(rows,source="K Samed Cil")
        self.assertEqual(result.decision,"Green + review")
        self.assertEqual(result.evidence.full_parts_similarity,90)
        self.assertEqual(result.adjusted,94)

    def test_duplicate_roster_rows_are_not_silently_collapsed(self):
        rows = [self.record(2,"Kaan Samed Cil"),self.record(3,"Kaan Samed Cil")]
        self.assertEqual(self.assess(rows).decision,"Review only")

    def test_competing_existing_exact_match_blocks_boost(self):
        rows = [self.record(2,"Kaan Samed Cil"),self.record(3,"K Samed Cil")]
        result = self.assess(rows)
        self.assertEqual(result.adjusted,result.baseline)
        self.assertEqual(result.decision,"Review only")

    def test_same_prefix_long_names_compete(self):
        rows = [self.record(2,"Adam Basel Cemal Deniz"),
                self.record(3,"Adam Basel Cemal Faruk")]
        self.assertEqual(self.assess(rows,source="Adam Basel Cemal").decision,"Review only")

    def test_location_section_and_weak_evidence_guards(self):
        rows = [self.record(2,"Kaan Samed Cil")]
        self.assertEqual(self.assess(rows,school="unrelated school").decision,"Review only")
        self.assertEqual(self.assess(rows,section="").decision,"Review only")
        self.assertEqual(self.assess(rows,section="b").decision,"Yellow + review")
        self.assertEqual(self.assess([self.record(2,"Kaan Cil")],source="K Cil").decision,"Review only")

    def test_unrelated_score_remains_bitwise_equal(self):
        rows = [self.record(2,"Kaan Selim Cil")]
        result = self.assess(rows,source="Kaan Samed Cil")
        self.assertEqual(result.decision,"Unchanged")
        self.assertEqual(result.baseline,result.adjusted)


if __name__ == "__main__":
    unittest.main()
