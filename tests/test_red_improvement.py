"""Test the red-only denominator, weighted totals and original classifications."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts/05_erc_roster_matching'))
from audit_structured_names import baseline_red_status, summarize_red_improvement
from student_match_report import BaseMatcher, BaseRecord, string_similarity


def row(sheet, reason='red_name', decision='Review only', baseline=0.7, adjusted=0.7):
    return {'sheet':sheet, 'row':2, 'baseline_is_red':bool(reason),
            'baseline_red_reason':reason, 'decision':decision,
            'baseline':baseline, 'adjusted':adjusted}


class RedImprovementTests(unittest.TestCase):
    def test_denominator_includes_unflagged_and_location_reds(self):
        rows = [row('A', decision='Green + review', adjusted=0.97),
                row('A', decision='Unchanged'), row('A', reason='red_location'),
                row('A', reason='', decision='Green + review', baseline=0.94, adjusted=0.97)]
        counts = summarize_red_improvement(rows)['per_sheet']['A']
        self.assertEqual(counts['baseline_red'],3)
        self.assertEqual(counts['red_to_green'],1)
        self.assertEqual(counts['red_remaining'],2)
        self.assertAlmostEqual(counts['red_to_green_percent'],100/3)
        self.assertEqual(counts['name_red_to_green_percent'],50)

    def test_overall_uses_counts_not_mean_of_sheet_percentages(self):
        rows = [row('A', decision='Green + review', adjusted=0.97)] + [row('B') for _ in range(9)]
        report = summarize_red_improvement(rows)
        self.assertEqual(report['per_sheet']['A']['red_to_green_percent'],100)
        self.assertEqual(report['per_sheet']['B']['red_to_green_percent'],0)
        self.assertEqual(report['overall']['red_to_green_percent'],10)

    def test_zero_reds_and_empty_sheets_have_no_percentage(self):
        report = summarize_red_improvement([row('A',reason='')],['A','Empty'])
        self.assertIsNone(report['per_sheet']['A']['red_to_green_percent'])
        self.assertIsNone(report['per_sheet']['Empty']['red_to_green_percent'])
        self.assertIsNone(report['overall']['red_to_green_percent'])
        self.assertIsNone(report['overall']['name_red_to_green_percent'])

    def test_name_only_rate_excludes_location_promotions(self):
        rows = [row('A', reason='red_location', decision='Green + review', adjusted=0.97),
                row('A', reason='red_name')]
        counts = summarize_red_improvement(rows)['overall']
        self.assertEqual(counts['red_to_green_percent'],50)
        self.assertEqual(counts['name_red_to_green_percent'],0)
        self.assertEqual(counts['location_red_to_green'],1)

    def test_yellow_is_separate_and_ninety_does_not_qualify(self):
        rows = [row('A', decision='Yellow + review', adjusted=0.97),
                row('A', decision='Green + review', adjusted=0.9)]
        counts = summarize_red_improvement(rows)['overall']
        self.assertEqual(counts['red_to_green'],0)
        self.assertEqual(counts['red_to_yellow'],1)
        self.assertEqual(counts['red_remaining'],1)
        self.assertEqual(counts['red_to_green_percent'],0)
        self.assertEqual(counts['red_reduction_percent'],50)

    def test_repeated_periods_are_counted_per_sheet(self):
        rows = [row(s,decision='Green + review',adjusted=0.97) for s in ['Term 1','Term 2']]
        report = summarize_red_improvement(rows)
        self.assertEqual(report['overall']['baseline_red'],2)
        self.assertEqual(report['per_sheet']['Term 1']['baseline_red'],1)

    def test_missing_red_reason_is_rejected(self):
        item = row('A')
        item['baseline_red_reason']=''
        with self.assertRaises(ValueError):
            summarize_red_improvement([item])


class BaselineRedTests(unittest.TestCase):
    def record(self, n, name, school='school', section='a'):
        return BaseRecord(n,'district',school,section,name,'District',school,section,name)

    def test_shortcut_and_fallback_agree_with_original_matcher(self):
        records = [self.record(2,'kaan samed cil'),self.record(3,'kerem samed cil',section='b'),
                   self.record(4,'adam basel cemal'),self.record(5,'adam basel cemal',section='b'),
                   self.record(6,'kaan samed cil',school='other')]
        matcher=BaseMatcher(records,student_threshold=90,location_threshold=95)
        cases=[
            ('school','a','kaan samed cil',records[0]),
            ('school','z','kaan samed cil',records[0]),
            ('school','a','k samed cil',records[0]),
            # Low score to the selected pair, but a different exact ERC match.
            ('school','a','kerem samed cil',records[0]),
            ('school','a','adam basel cemal',records[2]),
            # A high pair score in the wrong school must not shortcut.
            ('other','a','adam basel cemal',records[2]),
            ('missing','a','kaan samed cil',records[0]),
            ('school','a','kaan samed cil',None),
            ('school','a','',records[0]),
        ]
        for school,section,name,target in cases:
            with self.subTest(school=school,section=section,name=name,target=target):
                score=string_similarity(name,target.student) if target else 0
                actual=baseline_red_status(matcher,'district',school,section,name,target,score)
                expected=matcher.classify('district',school,section,name).status
                self.assertEqual(actual,expected if expected.startswith('red_') else '')


if __name__=='__main__':
    unittest.main()
