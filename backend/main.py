from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import tempfile
from pathlib import Path
from pydantic import BaseModel
from typing import List
from backend.excel_parser import ExcelParser
from backend.report_generator import ReportGenerator
from backend.models import TemplateMetadata

app = FastAPI()

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Store uploaded files temporarily
TEMP_DIR = tempfile.gettempdir()

# Request models
class GenerateReportRequest(BaseModel):
    input_files: List[str]
    template_file: str
    selectedPlatform: str = "MPN"
    selectedFormat: str = "Banner"

@app.post("/api/upload")
async def upload_files(files: list[UploadFile] = File(...)):
    """Upload and validate input files"""
    uploaded_files = {}
    errors = []
    
    for file in files:
        if not file.filename.endswith('.xlsx'):
            errors.append(f"{file.filename}: Not an Excel file")
            continue
        
        try:
            # Save file temporarily
            temp_path = os.path.join(TEMP_DIR, file.filename)
            contents = await file.read()
            with open(temp_path, 'wb') as f:
                f.write(contents)
            
            # Validate it's a proper campaign report
            is_template = 'template' in file.filename.lower()
            
            if is_template:
                uploaded_files['template'] = {
                    'filename': file.filename,
                    'path': temp_path
                }
            else:
                # Try to parse filename to extract audience/burst
                try:
                    audience, burst = ExcelParser.parse_filename(file.filename)
                    label = f"{audience} - Burst {burst}"
                    
                    if 'inputs' not in uploaded_files:
                        uploaded_files['inputs'] = []
                    
                    uploaded_files['inputs'].append({
                        'filename': file.filename,
                        'path': temp_path,
                        'label': label,
                        'audience': audience,
                        'burst': burst
                    })
                except Exception as e:
                    errors.append(f"{file.filename}: {str(e)}")
        
        except Exception as e:
            errors.append(f"{file.filename}: {str(e)}")
    
    return {
        'success': len(errors) == 0 or len(uploaded_files) > 0,
        'uploadedFiles': uploaded_files,
        'errors': errors
    }

@app.post("/api/generate-report")
async def generate_report(request: GenerateReportRequest):
    """Generate the consolidated report"""
    try:
        if not request.input_files or not request.template_file:
            raise HTTPException(
                status_code=400,
                detail="Both input files and template file are required"
            )
        
        # Validate files exist
        for file_path in request.input_files:
            if not os.path.exists(file_path):
                raise HTTPException(
                    status_code=400,
                    detail=f"Input file not found: {file_path}"
                )
        
        if not os.path.exists(request.template_file):
            raise HTTPException(
                status_code=400,
                detail=f"Template file not found: {request.template_file}"
            )
        
        # Parse all input files
        campaigns = []
        for file_path in request.input_files:
            try:
                campaign_data = ExcelParser.parse_input_file(file_path, request.selectedFormat)
                campaigns.append(campaign_data)
            except Exception as e:
                raise HTTPException(
                    status_code=400,
                    detail=f"Error parsing file {os.path.basename(file_path)}: {str(e)}"
                )
        
        # Parse template
        try:
            template_metadata = ExcelParser.parse_template_file(request.template_file)
        except Exception as e:
            raise HTTPException(
                status_code=400,
                detail=f"Error parsing template: {str(e)}"
            )
        
        # Generate report
        try:
            generator = ReportGenerator(
                campaigns,
                template_metadata,
                selected_platform=request.selectedPlatform,
                selected_format=request.selectedFormat,
            )
            workbook = generator.generate()
            
            # Save to temporary file
            output_filename = "MPN_Campaign_Report.xlsx"
            output_path = os.path.join(TEMP_DIR, output_filename)
            workbook.save(output_path)
            
            return {
                'success': True,
                'output_file': output_filename,
                'filename': output_filename
            }
        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=f"Error generating report: {str(e)}"
            )
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """Download the generated report"""
    try:
        full_path = os.path.join(TEMP_DIR, filename)
        if not os.path.exists(full_path):
            raise HTTPException(status_code=404, detail="File not found")
        
        return FileResponse(
            full_path,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename='MPN_Campaign_Report.xlsx'
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
async def root():
    """API root for the campaign report service"""
    return {
        "message": "MPN Campaign Report API is running.",
        "endpoints": [
            "/api/upload",
            "/api/generate-report",
            "/api/download/{filename}"
        ]
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
