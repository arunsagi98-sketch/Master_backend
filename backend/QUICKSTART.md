# MPN Campaign Report Generator - Quick Start

## 🚀 Get Started in 3 Minutes

### Step 1: Install Dependencies (2 min)
```bash
pip install -r requirements.txt
```

### Step 2: Generate Sample Files (1 min)
```bash
python generate_samples.py
```
This creates sample test files in a `sample_files/` folder.

### Step 3: Run the App
```bash
python run.py
```
Or on Windows, double-click: `run.bat`

### Step 4: Open in Browser
Navigate to: **http://localhost:8000**

You should see the application interface.

## 📁 Upload Files

1. **Campaign Files**: Upload one or more campaign report files
   - Must be .xlsx format
   - Must have "Vietnamese" or "Punjabi" in filename
   - Must have "Burst-1", "Burst-2", etc. in filename
   
2. **Template File**: Upload the template file
   - Must be .xlsx format
   - Must have "Report_Template" in filename

The app will auto-detect and validate the files.

## 🎯 Generate Report

1. Click **"Generate Report"** button (enables when files are uploaded)
2. Wait for progress bar to complete
3. Click **"Download Report"** to save your Excel file

## 📖 Documentation

- **[README.md](README.md)** - Full user guide
- **[PROJECT_SUMMARY.md](PROJECT_SUMMARY.md)** - Technical details
- **[TESTING_GUIDE.md](TESTING_GUIDE.md)** - Detailed testing and troubleshooting

## 🆘 Troubleshooting

**Port 8000 already in use?**
- Close other apps using port 8000
- Or modify `run.py` to use port 8001

**Missing dependencies?**
```bash
pip install -r requirements.txt
```

**Still stuck?**
- See TESTING_GUIDE.md for common issues
- Check that input files have all 10 required sheets
- Use sample files to test the app first

## 📋 What to Expect

**Input**: Multiple Excel campaign reports  
**Process**: Consolidates data from all files  
**Output**: Single Excel file with master report  

**Report includes:**
- Overview section (one row per campaign)
- Performance breakdowns by:
  - Device (Mobile, Tablet, Desktop)
  - Creative (all creatives from your files)
  - Age (18-24, 25-34, 35-44, 45-54, 55-64, 65+)
  - Gender (Male, Female)

## ✅ Verify Everything Works

1. Run the app: `python run.py`
2. Open: http://localhost:8000
3. See the UI load successfully ✓
4. Upload sample files from `sample_files/` folder
5. Click "Generate Report" ✓
6. Download and open the generated Excel file ✓

**Done! You're ready to process real campaign reports.**

---

For detailed information, see the other documentation files.
