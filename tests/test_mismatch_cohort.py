"""Historical cohort tests use synthetic records, never private student data."""
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1] / 'scripts/05_erc_roster_matching'))
from audit_mismatched_students import read_mismatch_cohort, summarize_cohort
from student_match_report import BaseMatcher, BaseRecord
from structured_name_match import assess_unmatched_student


class SheetStub:
    max_row=20  # Trailing formatting must not count as extra students.
    def iter_rows(self, min_row=1, **kwargs):
        if min_row==1:
            return iter([('Source Sheet','Source District','Source School','Source Section','Source Student Name')])
        return iter([('Term 1','district','school','a','k samed cil'),
                     (None,None,None,None,None),
                     ('Term 1','district','school','a','unresolved name')])


class BookStub:
    sheetnames=['Term 1','Mismatched Students']
    def __getitem__(self,key):
        return SheetStub()
    def close(self):
        pass


class HistoricalCohortTests(unittest.TestCase):
    def row(self,sheet='Term 1',status='red_name',decision='Unchanged red',review=False):
        return {'sheet':sheet,'baseline_status':status,'decision':decision,'review_required':review}

    def test_reader_excludes_header_and_empty_rows_but_keeps_unmatched(self):
        with patch('audit_mismatched_students.load_workbook',return_value=BookStub()):
            rows,order,physical=read_mismatch_cohort(Path('report.xlsx'))
        self.assertEqual(len(rows),2)
        self.assertEqual([r['mismatch_row'] for r in rows],[2,4])
        self.assertEqual(rows[1]['source_name'],'unresolved name')
        self.assertEqual(order,['Term 1'])
        self.assertEqual(physical,20)

    def test_wrong_input_cannot_silently_use_accepted_pair_subset(self):
        book=BookStub()
        book.sheetnames=['Term 1','Term 2']
        with patch('audit_mismatched_students.load_workbook',return_value=book):
            with self.assertRaisesRegex(ValueError,'original student matching report'):
                read_mismatch_cohort(Path('all_grades_final.xlsx'))

    def test_original_empty_report_message_is_not_a_student(self):
        sheet=SheetStub()
        original=sheet.iter_rows
        sheet.iter_rows=lambda min_row=1,**kwargs: (
            original(min_row=1) if min_row==1 else
            iter([('No mismatched students were found.',None,None,None,None)]))
        with patch('audit_mismatched_students.load_workbook',return_value=BookStub()), \
             patch.object(BookStub,'__getitem__',return_value=sheet):
            rows,_,_=read_mismatch_cohort(Path('report.xlsx'))
        self.assertEqual(rows,[])

    def test_denominator_keeps_original_members_even_if_baseline_changed(self):
        rows=[self.row(decision='Green + review',review=True), self.row(),
              self.row(status='red_location'),
              self.row(status='green',decision='Baseline already nonred')]
        total=summarize_cohort(rows)['overall']
        self.assertEqual(total['original_red_records'],4)
        self.assertEqual(total['red_to_green'],1)
        self.assertEqual(total['baseline_now_nonred'],1)
        self.assertEqual(total['still_red'],2)
        self.assertEqual(total['red_to_green_percent'],25)

    def test_unchanged_nonred_must_not_be_credited_as_improvement(self):
        with self.assertRaises(ValueError):
            summarize_cohort([self.row(status='green',decision='Green + review')])

    def test_yellow_does_not_inflate_green_rate_and_overall_is_weighted(self):
        rows=[self.row(decision='Green + review',review=True),
              self.row('Term 2',decision='Yellow + review',review=True)]
        rows += [self.row('Term 2') for _ in range(8)]
        report=summarize_cohort(rows)
        self.assertEqual(report['overall']['red_to_green_percent'],10)
        self.assertEqual(report['overall']['red_reduction_percent'],20)
        self.assertEqual(report['per_sheet']['Term 1']['red_to_green_percent'],100)
        self.assertEqual(report['per_sheet']['Term 2']['red_to_green_percent'],0)

    def test_zero_cohort_percentage_is_unavailable(self):
        self.assertIsNone(summarize_cohort([],['Empty'])['overall']['red_to_green_percent'])


class UnmatchedSearchTests(unittest.TestCase):
    def record(self,n,name,section='a'):
        return BaseRecord(n,'district','school',section,name,'District','School',section,name)

    def run_search(self,records,name='k samed cil',school='school',section='a'):
        matcher=BaseMatcher(records,student_threshold=90,location_threshold=95)
        return assess_unmatched_student(matcher,'district',school,section,name)

    def test_finds_initial_match_without_existing_pair(self):
        result=self.run_search([self.record(2,'kaan samed cil')])
        self.assertEqual(result.baseline.status,'red_name')
        self.assertEqual(result.promoted[0].row_number,2)
        self.assertEqual(result.promoted[1].decision,'Green + review')

    def test_searches_beyond_baseline_best_candidate(self):
        # The baseline favors a full-word typo; structured evidence should find
        # the positional initial match elsewhere in the same school.
        result=self.run_search([self.record(2,'k sami cil'),self.record(3,'kaan samed cil')])
        self.assertEqual(result.promoted[0].row_number,3)

    def test_ambiguous_initials_remain_review_only(self):
        result=self.run_search([self.record(2,'kaan samed cil'),self.record(3,'kerem samed cil','b')])
        self.assertIsNone(result.promoted)
        self.assertEqual(len(result.candidates),2)

    def test_nonred_baseline_and_failed_location_get_no_credit(self):
        record=self.record(2,'kaan samed cil')
        self.assertIsNone(self.run_search([record],name=record.student).promoted)
        result=self.run_search([record],school='unrelated')
        self.assertIsNone(result.promoted)
        self.assertEqual(result.baseline.status,'red_location')

    def test_section_mismatch_is_yellow_not_green(self):
        result=self.run_search([self.record(2,'kaan samed cil')],section='b')
        self.assertEqual(result.promoted[1].decision,'Yellow + review')


if __name__=='__main__':
    unittest.main()
