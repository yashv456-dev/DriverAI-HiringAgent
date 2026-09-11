# Version stack — P1 and P2

Verified against the live environment on 2026-09-11 by reading the installed
packages, the running model host and the built flow package. Not copied from
the requirements files, which declare floors rather than what is installed.

## Runtime

| Layer | P1 | P2 |
|---|---|---|
| Execution | Power Automate cloud flow | Local Python process |
| Python | 3.11.9 (tooling only) | 3.14.0 in its virtualenv |
| SQL | none — Excel table over Graph | SQLite 3.50.4 |
| LLM | none | Ollama 0.34.0 |
| Model | none | `qwen3-1.7b-p2:latest` (1.4 GB) |
| Model host | none | `http://localhost:11434` |
| Regex | **not available on the platform** | Python stdlib `re` |
| Pattern matching | chained `contains` / `replace` | full regex |

P1 runs no Python at execution time. The flow executes inside Power Automate and
keeps its data in an Excel table reached over Graph, with no database. Python
appears only in the build and test tooling (`flow/build_zip.py`, `test_p1.py`).

## OCR — P2 only

| Role | Component | Version |
|---|---|---|
| OCR primary | RapidOCR ONNX Runtime | 1.2.3 |
| OCR fallback | Tesseract | 5.4.0 |
| Fallback wrapper | pytesseract | 0.3.13 |
| Page rendering | PyMuPDF | 1.28.2 |
| Image handling | Pillow | 12.3.0 |
| PDF text, 1st | pymupdf4llm | installed |
| PDF text, 2nd | pypdf | 6.16.2 |
| PDF text, 3rd | pdfplumber | 0.11.10 |

Extraction is a ladder and OCR is its last rung. A PDF is read first with
pymupdf4llm, which preserves multi-column reading order; then pypdf against the
raw text layer; then pdfplumber. Only when all three return nothing — a scan or
image-only document — does OCR run.

OCR executes in a spawned, killable subprocess against a deadline of 120 s
(`HIRING_OCR_TIMEOUT`). The budget exists because the ONNX work is native code
and would otherwise be uninterruptible. RapidOCR renders each page at 150 dpi;
Tesseract is tried only if RapidOCR raises or returns nothing.

OCR is declared in `requirements-desktop.txt`, not `requirements.txt`. The
serverless path deliberately omits it, so OCR exists on local runs only.

## Supporting libraries — P2

| Library | Version |
|---|---|
| requests | 2.34.2 |
| openpyxl | 3.1.5 |
| pandas | 3.0.5 |
| numpy | 2.5.2 |
| PyYAML | 6.0.3 |
| python-dotenv | 1.2.3 |
| cryptography | 50.0.1 |
| azure-functions | 2.2.0 |

## Known inconsistencies

**Three Python versions are in play.** The P2 virtualenv runs 3.14.0, the
GitHub Actions workflow pins 3.11, and the Azure Functions path assumes 3.11.
Anything relying on 3.12+ syntax or stdlib behaviour will pass locally and fail
on both remote paths.

**P1's lack of regex is a platform limit, not an omission.** Power Automate
offers only case-insensitive substring matching, with no word boundaries and no
patterns. This shapes real design decisions in the flow: executable extensions
are stored space-terminated so `.exe ` matches a real attachment but not
`exeter.ac.uk`, and the resume filename sanitiser strips illegal characters
through eleven nested `replace()` calls because there is no pattern to do it in
one pass. Any new spam term must be safe as a bare substring.

**The `ollama` Python package is not installed.** The model is reached over
HTTP through `requests`, so there is no client library to pin.
