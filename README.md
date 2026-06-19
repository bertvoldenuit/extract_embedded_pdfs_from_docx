# Extract Embedded PDFs from DOCX

A small Python utility to extract PDF files embedded as OLE objects inside Microsoft Word `.docx` documents.

## Motivation

Sometimes a PDF is embedded in a Word document as an object:

```text
Insert → Object → Create from File
```

Word stores these embedded objects under:

```text
word/embeddings/
```

inside the `.docx` archive.

In some situations:

- Double-clicking the embedded PDF does nothing.
- Word returns an OLE error.
- Adobe Reader cannot open the embedded object.
- Renaming `oleObject1.bin` to `.pdf` does not work.

However, the original PDF data is often still present inside the OLE container.

This script automatically:

1. Opens the `.docx` file as a ZIP archive.
2. Finds embedded OLE objects (`*.bin`).
3. Locates PDF streams using the `%PDF` and `%%EOF` markers.
4. Extracts the PDF into standalone files.

## Features

- Extract PDFs directly from `.docx` files
- Process entire folders
- Recursive scan option
- Dry-run mode
- No external dependencies
- Works with standard Python 3

## Installation

Clone the repository:

```bash
git clone <your_repository_url>
cd extract-embedded-pdfs-from-docx
```

No additional packages are required.

## Usage

### Extract PDFs from a DOCX

```bash
python extract_embedded_pdfs_from_docx.py document.docx
```

Output:

```text
document_extracted_pdfs/
```

### Specify an output directory

```bash
python extract_embedded_pdfs_from_docx.py document.docx -o output
```

### Process all DOCX files in a folder

```bash
python extract_embedded_pdfs_from_docx.py ./documents
```

### Process recursively

```bash
python extract_embedded_pdfs_from_docx.py ./documents --recursive
```

### Dry run

```bash
python extract_embedded_pdfs_from_docx.py document.docx --dry-run
```

## Example

Input:

```text
Report.docx
└── embedded PDF object
```

Output:

```text
Report_extracted_pdfs/
└── Report_oleObject1.pdf
```

## How It Works

A `.docx` file is actually a ZIP archive:

```text
Report.docx
└── word
    └── embeddings
        └── oleObject1.bin
```

Some embedded PDF objects are stored inside OLE containers rather than as plain PDF files.

Instead of trying to open or rename the `.bin` file, this tool scans the binary content and extracts the actual PDF byte stream:

```text
...binary data...
%PDF
...
%%EOF
...binary data...
```

The extracted stream is then written back as a valid `.pdf` file.

## Limitations

- This tool only extracts PDF streams that are present inside OLE embedded objects.
- If the embedded PDF itself is corrupted, extraction cannot repair it.
- Non-PDF OLE objects (Excel, PowerPoint, ZIP, etc.) are ignored.

## Why This Exists

After receiving a Word document containing embedded PDFs, Word detected the objects but failed to open them.

Renaming:

```text
oleObject1.bin → oleObject1.pdf
```

did not work.

Extracting the PDF stream directly from the OLE binary container solved the issue.

Hopefully this utility saves someone else from the same frustration.

## License

MIT License
