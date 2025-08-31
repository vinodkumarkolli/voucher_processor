# Voucher Processor - Standalone GUI Application

A standalone Python GUI application that processes voucher templates (DOCX) and CSV data to generate PDF output files with processing history tracking.

## Features

- **User-friendly GUI**: Easy-to-use interface built with tkinter for maximum compatibility
- **File Upload**: Browse and select template DOCX and CSV files
- **Automated Processing**: Converts templates with CSV data into formatted PDF output
- **History Tracking**: Maintains processing history in an Excel spreadsheet
- **Progress Indicators**: Real-time status updates and progress tracking
- **File Management**: Automatic input/output file organization
- **Standalone Deployment**: Can be packaged as executable for any system

## Requirements

### Runtime Dependencies
- Python 3.7 or higher
- LibreOffice (for PDF conversion)

### Python Packages
```
pandas>=1.3.0
openpyxl>=3.0.0
python-docx>=0.8.11
lxml>=4.6.0
```

## Installation & Setup

### Option 1: Run from Python Source

1. **Clone or download the application files**
   ```bash
   # Ensure you have these files:
   # - voucher_gui.py (main GUI application)
   # - converter.py (processing engine)
   # - requirements.txt (dependencies)
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Install LibreOffice** (required for PDF conversion)
   - **Ubuntu/Debian**: `sudo apt-get install libreoffice`
   - **CentOS/RHEL**: `sudo yum install libreoffice`
   - **Windows**: Download from https://www.libreoffice.org/
   - **macOS**: `brew install libreoffice` or download from website

4. **Run the application**
   ```bash
   python voucher_gui.py
   ```

### Option 2: Create Standalone Executable

1. **Install cx_Freeze** (for creating executables)
   ```bash
   pip install cx_Freeze
   ```

2. **Build the executable**
   ```bash
   python setup.py build
   ```

3. **Distribute the built application**
   - The executable will be created in the `build/` directory
   - Copy the entire build directory to target systems
   - Ensure LibreOffice is installed on target systems

## Usage

### 1. Starting the Application
- Run `python voucher_gui.py` or execute the standalone application
- The GUI will open with file upload options

### 2. File Upload
- **Template File**: Click "Browse" next to "Template File (.docx)" to select your DOCX template
- **Voucher Data**: Click "Browse" next to "Voucher Data (.csv)" to select your CSV file

### 3. CSV File Format
Your CSV file must contain these columns:
- `ID`: Unique identifier for each voucher
- `Secret Code`: Secret code for each voucher

Example CSV structure:
```csv
ID,Secret Code
V001,ABC123
V002,XYZ789
V003,DEF456
```

### 4. Template File Format
Your DOCX template should contain placeholders in this format:
- `[sno_1]` to `[sno_14]`: Will be replaced with ID values
- `[code_1]` to `[code_14]`: Will be replaced with Secret Code values

The template supports up to 14 vouchers per page. Multiple pages will be generated automatically if you have more than 14 records.

### 5. Processing
1. Click "Process Vouchers" after uploading both files
2. Monitor the progress in the Status section
3. Processing will create intermediate files in `filesIn/` and output files in `fileOut/`

### 6. Download Output
- Once processing is complete, the "Download PDF" button will be enabled
- Click to save the generated PDF to your desired location
- You can also click "Open Output Folder" to view all output files

### 7. Processing History
- All processing activities are logged in `processing_history.xlsx`
- The History section shows recent processing activities
- Click "Refresh History" to update the display

## File Organization

The application creates and manages these directories:

```
Application Directory/
├── voucher_gui.py          # Main GUI application
├── converter.py            # Processing engine
├── requirements.txt        # Python dependencies
├── setup.py               # Executable build script
├── README.md              # This file
├── processing_history.xlsx # Processing history (auto-created)
├── filesIn/               # Input files (auto-created)
│   ├── template.docx      # Copied template file
│   └── Gift Voucher.csv   # Copied CSV file
└── fileOut/               # Output files (auto-created)
    └── template_output.pdf # Generated PDF
```

## Deployment on Other Systems

### Method 1: Python Environment
1. Copy all application files to target system
2. Install Python 3.7+ and LibreOffice
3. Install dependencies: `pip install -r requirements.txt`
4. Run: `python voucher_gui.py`

### Method 2: Standalone Executable
1. Build executable on development system: `python setup.py build`
2. Copy the entire `build/` directory to target system
3. Install LibreOffice on target system
4. Run the executable from the build directory

### Cross-Platform Notes
- **Windows**: Executable will have `.exe` extension
- **Linux/macOS**: Executable will be named `VoucherProcessor`
- **LibreOffice Path**: The application assumes LibreOffice is in system PATH

## Troubleshooting

### Common Issues

1. **"LibreOffice conversion error"**
   - Ensure LibreOffice is installed and accessible via command line
   - Test: `libreoffice --version` should work

2. **"CSV file format error"**
   - Verify CSV has 'ID' and 'Secret Code' columns
   - Check for proper CSV formatting (commas, quotes, encoding)

3. **"Template processing error"**
   - Ensure DOCX template contains the expected placeholders
   - Verify template is not corrupted or password-protected

4. **"Permission denied" errors**
   - Ensure write permissions in application directory
   - Run with appropriate user permissions

5. **GUI not responsive during processing**
   - This is normal; processing runs in background thread
   - Monitor status messages for progress updates

### Performance Notes
- Processing time depends on number of records and system performance
- Large CSV files (1000+ records) may take several minutes
- PDF conversion is usually the longest step

## Technical Details

### Architecture
- **GUI Framework**: tkinter (included with Python)
- **Document Processing**: Direct XML manipulation of DOCX files
- **PDF Conversion**: LibreOffice headless mode
- **Data Handling**: pandas for CSV processing
- **History Storage**: Excel format using openpyxl

### Security Considerations
- Application processes files locally (no network communication)
- Temporary files are created and cleaned up automatically
- Input files are copied to `filesIn/` directory for processing
- Original files remain unchanged

## Support & Maintenance

### Adding Features
The application is modular and can be extended:
- Modify `voucher_gui.py` for GUI changes
- Modify `converter.py` for processing logic changes
- Update `requirements.txt` for new dependencies

### Logs and Debugging
- Status messages are displayed in the GUI
- Processing history is stored in Excel format
- Add print statements to converter.py for detailed debugging

## Version History

- **v1.0.0**: Initial release with full GUI functionality
  - File upload interface
  - Automated processing
  - PDF generation
  - History tracking
  - Standalone deployment support