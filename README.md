# Inspection Report Automation

Python application for automating the conversion of dimensional inspection reports.

The application reads an inspection report in PDF format, extracts dimensional data and measurement results, combines them with user-provided information, and generates standardized Excel and PDF inspection reports.

## Features

- Extracts Part Number automatically from the source PDF
- Extracts product dimensions and inspection quantity
- Extracts dimensional inspection measurements
- Supports single and multi-line tolerances
- Preserves the formatting of an existing Excel report template
- Allows the user to provide:
  - Customer name
  - Final Part Number
  - Report description
- Automatically inserts the current date
- Generates both Excel and PDF reports
- Supports multiple report conversions in the same execution
- Can be packaged as a standalone Windows executable

## Workflow

```text
Source Inspection PDF
        ↓
PDF Data Extraction
        ↓
Structured Python Data
        ↓
User Input
        ↓
Excel Template
        ↓
Standardized Excel Report
        ↓
PDF Report
```

## Technologies

- Python
- pypdf
- Regular Expressions
- pywin32
- Microsoft Excel COM Automation
- PyInstaller

## Requirements

- Windows
- Microsoft Excel
- Python 3.x when running from source

Install the required Python packages with:

```bash
pip install -r requirements.txt
```

## Running from source

```bash
python converter.py
```

The application will ask for:

1. The Excel template path
2. The source inspection report PDF
3. The final Part Number
4. The customer/company name
5. The report description

The generated Excel and PDF reports are saved in the same directory as the source inspection report.

## Packaging as an executable

The application can be packaged with PyInstaller:

```bash
python -m PyInstaller --clean --onefile --console --name ReportControl --hidden-import=win32timezone converter.py
```

The executable will be created inside the `dist` directory.

## Notes

The generated reports rely on Microsoft Excel COM automation. Therefore, Microsoft Excel must be installed on the computer running the application.

Real inspection reports, customer information and production templates are not included in this repository.