# MPN Banner Campaign Report Generator

A FastAPI backend service that consolidates multiple Excel campaign report files into a single Master Performance Report.

## Features

- **Multi-file upload API**: Upload multiple campaign report Excel files in one request
- **Template-based processing**: Use a template file to control output formatting
- **Automatic filename parsing**: Extracts audience and burst information from filenames
- **Consolidated reporting**: Produces a single `.xlsx` output file
- **Performance breakdown**: Includes device, creative, age, and gender summaries
- **API-first design**: No built-in HTML frontend required

## Requirements

- Python 3.8+
- pip (Python package manager)

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

## Running the Application

1. Start the server:
```bash
python run.py
```

2. Use the JSON API endpoints or an external HTTP client to upload and process report files.

## Usage

1. **Upload Campaign Files**: 
   - Call `POST /api/upload` with multipart form data
   - Attach one or more Final Report Excel files
   - Each campaign file should contain these sheets: REACH, DATE, APP URL, TIME OF DAY, EXCHANGE, DEVICE, CREATIVE, CITY, AGE, GENDER

2. **Upload Template File**:
   - Include one template `.xlsx` file in the same `POST /api/upload` request
   - The filename should contain `Report_Template`
   - The template should contain a `MPN & CPN Breakdown` sheet

3. **Generate Report**:
   - Call `POST /api/generate-report` with the uploaded file paths returned from `/api/upload`
   - Wait for the report generation to complete

4. **Download**:
   - Call `GET /api/download/MPN_Campaign_Report.xlsx` to retrieve the generated report

## Project Structure

```
MAster_File_Format/
├── backend/
│   ├── __init__.py
│   ├── main.py              # FastAPI application
│   ├── models.py            # Data models
│   ├── excel_parser.py      # Excel file parsing logic
│   └── report_generator.py  # Report generation logic
├── requirements.txt         # Python dependencies
├── run.py                   # Application entry point
└── README.md                # This file
```

## Input File Format

### Campaign Report Files
- **Filename**: Should contain audience name (Vietnamese/Punjabi) and burst number (e.g., `..._Vietnamese_Burst-1_...`)
- **Sheets Required**: 
  - REACH, DATE, APP URL, TIME OF DAY, EXCHANGE, DEVICE, CREATIVE, CITY, AGE, GENDER
- **Data Format**:
  - REACH sheet, Row 2: Impressions, Clicks, CTR, Reach, Frequency
  - Other sheets: Column 1 = Category, Columns 2-4 = Impressions, Clicks, CTR

### Template File
- **Filename**: Should contain "Report_Template"
- **Required Sheet**: "MPN & CPN Breakdown"

## Output Format

The generated report contains:
- **Overview section**: Summary metrics for each campaign
- **Performance breakdowns** (for each campaign):
  - By Device (Mobile, Tablet, Desktop)
  - By Creative
  - By Age (18-24, 25-34, 35-44, 45-54, 55-64, 65+)
  - By Gender (Male, Female)

## API Endpoints

- `GET /` - API root for the campaign report service
- `POST /api/upload` - Upload Excel files
- `POST /api/generate-report` - Generate consolidated report
- `GET /api/download/{filename}` - Download generated report

## Technology Stack

- **Backend**: FastAPI (Python)
- **Excel Processing**: openpyxl
- **Frontend**: Vanilla HTML/CSS/JavaScript
- **Server**: Uvicorn

## Error Handling

The application validates:
- File format (must be .xlsx)
- Required sheets in input files
- Filename format for audience/burst extraction
- Template file structure

Errors are displayed in the UI with clear messages.

## Performance

- Optimized for processing multiple campaigns
- Efficient Excel file reading and writing
- Progress indication during report generation

## License

© 2024 MPN Campaign Report Generator

## Support

For issues or questions, please check the error messages displayed in the application.
