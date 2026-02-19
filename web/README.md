# Essay Processor - Web Version

A pure HTML/JavaScript application that runs entirely in the browser. No installation required.

## Usage

1. Open `index.html` in any modern web browser (Chrome, Firefox, Edge)
2. Click **Load PDFs** to select PDF feedback files
3. Click **Load DOCX** to select matching DOCX files
4. Click **Process All** to generate reports
5. Click each report name to download

## Features

- Runs entirely client-side (no server needed)
- Same functionality as the Python version
- Reports generated as downloadable DOCX files

## Libraries Used (via CDN)

- **PDF.js** - PDF text extraction
- **Mammoth.js** - DOCX reading
- **docx** - DOCX file generation
- **JSZip** - ZIP/DOCX parsing
- **FileSaver.js** - File downloads

## Notes

- Requires internet connection on first load (CDN libraries)
- All processing happens in your browser - files never leave your computer
