# School grade extraction and ERC student roster matching

I built a Python workflow to organize school grade records and compare them with the ERC full student roster. The project collection contains **more than 10,000 PDF and Excel files** across the original material and additional search folders, including working copies. The roster lists students by district, school, class section, and name. After extracting and consolidating the grade reports into `all_grades.xlsx`, the workflow compares its student records with that roster to identify likely matches, ambiguous candidates, and records needing review. The grade reports came in several layouts, from selectable tables to scanned pages, so building the workbook required extraction scripts and OCR.

This repository contains the code for extracting and combining grades, finding missing records, matching students to the ERC roster, and reviewing uncertain matches. The school files, student records, and ERC roster are kept separately.

## How the work progressed

1. **Extract the school reports.** I processed the PDFs school by school. There were seven main table layouts, each with its own script. Some PDFs contained selectable text; others were scans and needed OCR. A separate script handled a detailed fourth-grade Excel layout.
2. **Build `all_grades.xlsx`.** The extraction scripts produced school-level Excel files. I merged those files and also copied grades manually from some Excel sources that already contained the information I needed.
3. **Check the master workbook.** I first made three overview programs. `overview1` looks for missing school sections, `overview2` compares grade records with the ERC student roster, and `overview3` follows students across grade periods. Most of the later matching work grew out of `overview2`.
4. **Match students to the ERC roster.** [`student_match_report.py`](scripts/05_erc_roster_matching/student_match_report.py) first resolves the district and school, then compares each grade-record student with ERC students at that school. It distinguishes unique matches in the same class section, possible matches in another section, ambiguous candidates, and records without a reliable match. I developed `overview2` into a review workbook with **nine sheets**, adding manual `MyScore` assessments, duplicate checks, and reports showing which ERC students have grade records. Follow-up scripts transfer my decisions into refreshed reports and apply approved corrections.
5. **Look for missing records in other folders.** I checked additional folders after the first `all_grades` workbook was built. This search collection contains **more than 8,000 PDF and Excel files**. [`discover_sources.py`](scripts/04_extra_source_search/discover_sources.py) was added afterward so the candidate search could be run again in a consistent way. It searches PDF text and spreadsheet cells for school names in the missing-section report. I still inspect the suggested files for the right district, grade, term, class section, and actual grades. When a source really contains missing grades, it goes through extraction, merging, and ERC roster matching again.

## A custom algorithm for more precise fuzzy name matching

I developed and tested a fuzzy name-matching algorithm that achieved a **16.59% improvement** in identifying potential matches among previously unmatched student records. It addresses grade reports that use initials or omit the last parts of longer names. It compares name parts in their original order. For names with initials, it requires at least three parts and an average similarity of at least 90% across the remaining full name parts. It also checks the district, school, and class section while considering competing candidates. It preserves the existing matching scores for other cases.

I implemented the algorithm in [`structured_name_match.py`](scripts/05_erc_roster_matching/structured_name_match.py). The accompanying [`audit_mismatched_students.py`](scripts/05_erc_roster_matching/audit_mismatched_students.py) script runs the evaluation against the original matching report and ERC student roster, and saves the improvement counts, scores, and reasons for each suggestion.

## What is in each folder

| Folder | Purpose |
|---|---|
| [`scripts/01_extraction/selectable`](scripts/01_extraction/selectable/) | PDF extractors for the seven table layouts and the multi-page, two-term, and per-course variations. |
| [`scripts/01_extraction/ocr`](scripts/01_extraction/ocr/) | OCR versions for PDFs that do not have usable text. |
| [`scripts/01_extraction/spreadsheets`](scripts/01_extraction/spreadsheets/) | Extracts grades from a detailed fourth-grade Excel format. |
| [`scripts/01_extraction/utilities`](scripts/01_extraction/utilities/) | Splits long PDFs before extraction when needed. |
| [`scripts/02_consolidation`](scripts/02_consolidation/) | Merges school Excel outputs. `excel_merge.py` writes into the first workbook in the folder, so use copies of source files. |
| [`scripts/03_missing_sections`](scripts/03_missing_sections/) | The `overview1` check for school sections absent from `all_grades`. |
| [`scripts/04_extra_source_search`](scripts/04_extra_source_search/) | Searches additional folders and lists possible sources for missing sections. |
| [`scripts/05_erc_roster_matching`](scripts/05_erc_roster_matching/) | Matches grade-record students to the ERC roster and supports review of uncertain names, manual corrections, duplicates, and roster coverage. |
| [`scripts/06_tracking`](scripts/06_tracking/) | The `overview3` comparison across grade periods. |
| [`tests`](tests/) | Synthetic checks for the extra-folder search, name-matching rules, and improvement counts. |

The seven main PDF parsers are `extract_pdf_grades.py` and `extract_pdf_grades_type2_selectable.py` through `extract_pdf_grades_type7sultanbeyli.py` in [`selectable`](scripts/01_extraction/selectable/). The names identify the layouts I encountered; they do not mean every school used the same format. The OCR folder contains the scanned-PDF variants for the layouts that needed them.

## Searching additional folders

[`discover_sources.py`](scripts/04_extra_source_search/discover_sources.py) writes `candidates.xlsx` with each suggested source file, its page or Excel row range, and the reason it was suggested. It also saves an inventory, errors, and a scan summary as JSON logs. I check each suggested file for the right district, section, school year, term, and actual grades before using it. Scanned PDFs with no readable text need OCR or a visual check.

The scripts here record the steps and formats I worked with. They do not recreate the exact final workbook by themselves because some Excel additions and `MyScore` choices were manual.
