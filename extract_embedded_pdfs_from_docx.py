#!/usr/bin/env python3
"""
extract_embedded_pdfs_from_docx.py

Extract PDF files embedded as OLE objects inside Microsoft Word .docx files.

Why this exists
---------------
A .docx file is a ZIP archive. Embedded objects are usually stored under:

    word/embeddings/

For embedded PDFs, Word often stores them as OLE binary containers named like:

    oleObject1.bin
    oleObject2.bin

Renaming those .bin files to .pdf does not always work, because the PDF may be
wrapped inside the OLE container. This script scans the binary content, locates
PDF byte streams starting with '%PDF' and ending with '%%EOF', and writes them
back as standalone .pdf files.

Usage
-----
Extract PDFs from one DOCX:

    python extract_embedded_pdfs_from_docx.py my_document.docx

Choose output directory:

    python extract_embedded_pdfs_from_docx.py my_document.docx -o extracted_pdfs

Process all DOCX files in a folder:

    python extract_embedded_pdfs_from_docx.py ./documents --recursive

Dry-run without writing files:

    python extract_embedded_pdfs_from_docx.py my_document.docx --dry-run
"""

from __future__ import annotations

import argparse
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

PDF_HEADER = b"%PDF"
PDF_FOOTER = b"%%EOF"


@dataclass(frozen=True)
class ExtractedPdf:
    source_docx: Path
    source_bin: str
    output_pdf: Path | None
    start_offset: int
    end_offset: int


def iter_docx_files(path: Path, recursive: bool = False) -> Iterable[Path]:
    """Yield .docx files from a file or directory."""
    if path.is_file():
        if path.suffix.lower() == ".docx":
            yield path
        return

    pattern = "**/*.docx" if recursive else "*.docx"
    yield from sorted(path.glob(pattern))


def find_pdf_streams(data: bytes) -> list[tuple[int, int]]:
    """
    Return list of (start, end_exclusive) boundaries for PDF streams in bytes.

    A single OLE .bin can theoretically contain more than one PDF stream, so this
    function keeps scanning after each detected %%EOF marker.
    """
    streams: list[tuple[int, int]] = []
    search_from = 0

    while True:
        start = data.find(PDF_HEADER, search_from)
        if start == -1:
            break

        eof = data.find(PDF_FOOTER, start)
        if eof == -1:
            # Header found but no explicit EOF. Avoid writing a partial file.
            break

        end = eof + len(PDF_FOOTER)

        # Some PDFs contain trailing CR/LF after %%EOF. Keep it if present.
        while end < len(data) and data[end:end + 1] in (b"\r", b"\n"):
            end += 1

        streams.append((start, end))
        search_from = end

    return streams


def safe_output_name(docx_path: Path, bin_name: str, index: int | None = None) -> str:
    """Build a readable output filename."""
    stem = docx_path.stem
    bin_stem = Path(bin_name).stem

    if index is None:
        return f"{stem}_{bin_stem}.pdf"

    return f"{stem}_{bin_stem}_{index}.pdf"


def extract_pdfs_from_docx(
    docx_path: Path,
    output_dir: Path,
    dry_run: bool = False,
) -> list[ExtractedPdf]:
    """Extract embedded PDF streams from one .docx file."""
    results: list[ExtractedPdf] = []

    if not zipfile.is_zipfile(docx_path):
        print(f"[SKIP] Not a valid DOCX/ZIP file: {docx_path}")
        return results

    with zipfile.ZipFile(docx_path, "r") as zf:
        embedding_files = [
            name for name in zf.namelist()
            if name.startswith("word/embeddings/") and name.lower().endswith(".bin")
        ]

        if not embedding_files:
            print(f"[INFO] No OLE .bin embeddings found in: {docx_path}")
            return results

        for embedded_name in embedding_files:
            data = zf.read(embedded_name)
            streams = find_pdf_streams(data)

            if not streams:
                print(f"[SKIP] No PDF stream found in {docx_path.name} -> {embedded_name}")
                continue

            for i, (start, end) in enumerate(streams, start=1):
                output_name = safe_output_name(
                    docx_path,
                    Path(embedded_name).name,
                    None if len(streams) == 1 else i,
                )
                output_pdf = output_dir / output_name

                if dry_run:
                    output_path: Path | None = None
                    print(
                        f"[DRY] Would extract {docx_path.name} -> {embedded_name} "
                        f"[{start}:{end}] as {output_pdf}"
                    )
                else:
                    output_dir.mkdir(parents=True, exist_ok=True)
                    output_pdf.write_bytes(data[start:end])
                    output_path = output_pdf
                    print(f"[OK] {output_pdf}")

                results.append(
                    ExtractedPdf(
                        source_docx=docx_path,
                        source_bin=embedded_name,
                        output_pdf=output_path,
                        start_offset=start,
                        end_offset=end,
                    )
                )

    return results


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract PDFs embedded as OLE .bin objects inside DOCX files."
    )
    parser.add_argument(
        "input",
        type=Path,
        help="Input .docx file or directory containing .docx files.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=None,
        help="Output directory. Default: '<input_stem>_extracted_pdfs' for one file, or './extracted_pdfs' for a folder.",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="When input is a directory, process .docx files recursively.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be extracted without writing files.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = args.input

    if not input_path.exists():
        print(f"[ERROR] Input does not exist: {input_path}")
        return 1

    docx_files = list(iter_docx_files(input_path, recursive=args.recursive))

    if not docx_files:
        print(f"[ERROR] No .docx file found in: {input_path}")
        return 1

    if args.output_dir is not None:
        output_dir = args.output_dir
    elif input_path.is_file():
        output_dir = input_path.with_suffix("").parent / f"{input_path.stem}_extracted_pdfs"
    else:
        output_dir = input_path / "extracted_pdfs"

    total = 0
    for docx_path in docx_files:
        extracted = extract_pdfs_from_docx(
            docx_path=docx_path,
            output_dir=output_dir,
            dry_run=args.dry_run,
        )
        total += len(extracted)

    print(f"\nDone. PDF streams found: {total}")
    if not args.dry_run:
        print(f"Output directory: {output_dir}")

    return 0 if total > 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
