# AP English Rubric Guide

A tool for educators that combines essay feedback from PDF and DOCX source files into a cleanly formatted report document.

Available as both a **Python desktop application** and a **web-based HTML application**.

---

## Overview

This tool processes graded essays and generates consolidated reports containing:
- Student essay (with original paragraph formatting preserved)
- AP Rubric grading table
- Quotes and feedback organized by category

### Report Structure

1. **Header** — Student name, essay title, and date
2. **Essay** — Original student essay with paragraphs preserved
3. **AP Rubric** — Grading table (Overall, Thesis, Evidence and Commentary, Sophistication)
4. **Feedback Sections** — Each category shows quotes from the essay with specific feedback

---

## Features

- **Batch Processing** — Load multiple PDF and DOCX files at once
- **Automatic Matching** — Files are matched by student name
- **De-hyphenation** — Line-break hyphens are automatically removed
- **Preserved Formatting** — Essay paragraphs maintain original structure
- **Table Integrity** — AP Rubric table and header stay together on one page
- **Two Download Options** (Web version) — Individual files or single ZIP archive

---

## Versions

### Python Desktop Application

Located in the root folder (`essay_processor.py`).

**Requirements:**
```
pip install python-docx PyMuPDF
```

**Run:**
```
python essay_processor.py
```

Reports are saved to a dated folder on your Desktop (e.g., `report_data_18_Feb_2026/`).

### Web Application

Located in the `docs/` folder (`index.html`).

**Try it online:** [https://akingry.github.io/ap-english-rubric-guide/](https://akingry.github.io/ap-english-rubric-guide/)

**Usage:**
1. Open `index.html` in any modern browser
2. Load PDF files (left panel)
3. Load matching DOCX files (left panel)
4. Click "Process All" (right panel)
5. Download individual reports or use "Download as Zip"

No installation required — runs entirely in the browser using:
- PDF.js for PDF parsing
- JSZip for DOCX parsing and ZIP creation
- docx library for report generation

---

## Input File Requirements

### PDF File (Feedback)
Contains graded feedback with sections:
- **Grading** — Overall grade and comments
- **Evidence and Commentary** — Grade, overview, quotes with feedback
- **Sophistication** — Grade, overview, quotes with feedback
- **Thesis** — Grade, overview, quotes with feedback

### DOCX File (Essay + Rubric)
Contains:
- **Grading Table** — 4 rows × 3 columns (header row + 4 categories)
- **Content Review** — The original student essay

### File Naming Convention

```
LastName_ Essay Title_review.ext
```

**Example:**
```
Taylor_ Light pollution_review.pdf
Taylor_ Light pollution_review.docx
```

The student name is extracted from before the first underscore.

---

## Output Format

Each report includes:

| Section | Content |
|---------|---------|
| Header | Name, Essay title, Date |
| ESSAY | Original student essay (paragraphs preserved) |
| AP RUBRIC | Reordered grading table |
| OVERALL | Grade only (e.g., "OVERALL: 3/6") |
| Thesis | Heading + quotes with indented feedback |
| Evidence and Commentary | Heading + quotes with indented feedback |
| Sophistication | Heading + quotes with indented feedback |

### Section Order

Both the table and feedback sections follow this order:
1. Overall
2. Thesis
3. Evidence and Commentary
4. Sophistication

---

## Layout

### Python Version
- Single-column window with file lists and process button
- Status bar at bottom

### Web Version
- Two-column layout
- **Left:** Source files (PDF and DOCX lists with load buttons)
- **Right:** Generated reports, Process button, Download buttons
- Status bar spanning full width

---

## License

MIT License
