#!/usr/bin/env python3
"""Split a PDF file into smaller PDFs with a fixed number of pages."""

from argparse import ArgumentParser
from pathlib import Path
import sys

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("The required library 'pypdf' is not installed.")
    print("Install it with: python -m pip install pypdf")
    raise SystemExit(1)


DEFAULT_CHUNK_SIZE = 8


def build_parser():
    parser = ArgumentParser(description="Split a PDF into page chunks.")
    parser.add_argument(
        "input_pdf",
        nargs="?",
        help="Path to the PDF file. If omitted, you will be prompted for it.",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=DEFAULT_CHUNK_SIZE,
        help=f"Number of pages per split PDF. Default: {DEFAULT_CHUNK_SIZE}.",
    )
    parser.add_argument(
        "--output-folder",
        help=(
            "Folder where split files will be saved. If omitted, a folder named "
            "'split_output' is created next to the input PDF."
        ),
    )
    return parser


def clean_user_path(path_text):
    return path_text.strip().strip('"').strip("'")


def get_pdf_path(input_pdf):
    if not input_pdf:
        input_pdf = input("Enter the PDF file path: ")

    input_pdf = clean_user_path(input_pdf)
    if not input_pdf:
        raise ValueError("No PDF file path was entered.")

    pdf_path = Path(input_pdf).expanduser()
    if not pdf_path.exists():
        raise FileNotFoundError(f"File not found: {pdf_path}")
    if not pdf_path.is_file():
        raise ValueError(f"Path is not a file: {pdf_path}")
    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError("The input file must be a PDF file.")

    return pdf_path.resolve()


def get_output_folder(input_pdf_path, output_folder):
    if output_folder:
        output_path = Path(clean_user_path(output_folder)).expanduser()
        if not output_path.is_absolute():
            output_path = input_pdf_path.parent / output_path
    else:
        output_path = input_pdf_path.parent / "split_output"

    output_path.mkdir(parents=True, exist_ok=True)
    return output_path.resolve()


def split_pdf(input_pdf_path, output_folder, chunk_size):
    if chunk_size < 1:
        raise ValueError("Chunk size must be at least 1 page.")

    pdf_reader = PdfReader(str(input_pdf_path))
    total_pages = len(pdf_reader.pages)
    if total_pages == 0:
        raise ValueError("The PDF does not contain any pages.")

    created_files = []
    original_name = input_pdf_path.stem
    part_number = 1

    for start_page in range(0, total_pages, chunk_size):
        end_page = min(start_page + chunk_size, total_pages)
        pdf_writer = PdfWriter()

        for page_index in range(start_page, end_page):
            page = pdf_reader.pages[page_index]
            pdf_writer.add_page(page)

        output_file_name = (
            f"{original_name}_part_{part_number}_"
            f"pages_{start_page + 1}_to_{end_page}.pdf"
        )
        output_path = output_folder / output_file_name

        with output_path.open("wb") as output_file:
            pdf_writer.write(output_file)

        print(
            f"Created part {part_number} containing pages "
            f"{start_page + 1} to {end_page}: {output_path}"
        )

        created_files.append(output_path)
        part_number += 1

    print(f"Splitting completed. Created {len(created_files)} file(s).")
    print(f"Output folder: {output_folder}")
    return created_files


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        input_pdf_path = get_pdf_path(args.input_pdf)
        output_folder = get_output_folder(input_pdf_path, args.output_folder)
        split_pdf(input_pdf_path, output_folder, args.chunk_size)
    except (FileNotFoundError, ValueError) as error:
        print(error)
        return 1
    except Exception as error:
        print(f"Unable to split PDF: {error}")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
