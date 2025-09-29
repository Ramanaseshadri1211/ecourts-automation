
# eCourts Automation

Automates extraction of case details from the **eCourts India portal** using Python, Selenium, and OCR.  
This script handles CAPTCHA decoding, retrieves case information, downloads judgment PDFs, and saves results in Excel.

---

## Features

- Reads case numbers, years, and types from an Excel file.
- Automates portal navigation (state, district, court selection).
- Solves CAPTCHA using OCR (Tesseract + OpenCV preprocessing).
- Extracts case details:
  - First Hearing Date
  - Decision Date
  - Case Status
  - Nature of Disposal
  - Court and Judge
  - FIR Details
- Downloads judgment PDFs (if available).
- Saves results in Excel and screenshots for reference.
- Generic and configurable – safe for sharing on GitHub.

---

## Requirements

- Python 3.9+  
- Google Chrome  
- ChromeDriver (matching your Chrome version)  
- Python packages:
  ```bash
  pip install selenium pandas numpy opencv-python pillow pytesseract requests
