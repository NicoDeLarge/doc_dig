# Document Digitalizer

... providing a simple and lean workflow to create digital, searchable documents based on scans.

## Prequisites

This project uses https://github.com/tesseract-ocr/tesseract for OCR and making scanned documents searchable. Make sure to have the binaries in your PATH environment variable.

## Usage

Set the folders to the correct paths in generatePdf.py
- inbox: where new scans are saved e.g. from your scanner
- processed_dir: where the generated PDFs shall be stored
- archive_dir: where the orignal image files shall be saved as backup (if desired)

Then just run the script e.g. with pipenv having installed the necessary packages

    pipenv run python generatePdf.py
