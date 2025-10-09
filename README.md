# Excel File Processor - Attendance & Payroll

A modern Python desktop application for processing Excel files containing attendance and payroll data with intelligent structure recognition and data extraction.

## Features

- **Modern GUI Interface**: Clean, professional design with intuitive file upload functionality
- **Intelligent Attendance Processing**:
  - Automatically extracts Date, Company Name, and Narration from header rows
  - Recognizes employee data structure with proper validation
  - Extracts complete employee information (ID, Name, Attendance Type, Days)
  - Validates attendance sheet format and shows helpful error messages
- **Payroll Sheet Processing**: Basic Excel file analysis and data extraction
- **Data Export**: Export processed data to JSON format for further use
- **Excel File Support**: Supports both .xlsx and .xls file formats
- **Real-time Results Display**: View processing results with detailed information
- **Robust Validation**: Proper error handling and format validation

## Project Structure

```
excel_processor/
├── main.py                 # Main application file
├── src/
│   ├── __init__.py
│   └── excel_processor.py  # Excel processing logic
├── assets/                 # (Reserved for future icons/images)
├── requirements.txt        # Python dependencies
├── build.spec             # PyInstaller configuration
├── build.bat              # Build script for creating exe
├── run.bat                # Script to run the application
└── README.md              # This file
```

## Installation & Setup

### Option 1: Running from Source Code

1. **Install Python** (3.8 or higher) from [python.org](https://python.org)

2. **Navigate to the project directory**:

   ```powershell
   cd E:\TellyIMport\excel_processor
   ```

3. **Install dependencies**:

   ```powershell
   pip install -r requirements.txt
   ```

4. **Run the application**:

   ```powershell
   python main.py
   ```

   Or simply double-click `run.bat`

### Option 2: Creating Windows Executable

1. **Install dependencies** (if not already done):

   ```powershell
   pip install -r requirements.txt
   ```

2. **Build the executable**:

   ```powershell
   pyinstaller build.spec
   ```

   Or simply double-click `build.bat`

3. **Run the executable**:
   - The .exe file will be created in the `dist/` folder
   - Double-click `ExcelProcessor.exe` to run

## How to Use

1. **Launch the Application**:

   - Run from source: `python main.py` or double-click `run.bat`
   - Or run the compiled executable: `ExcelProcessor.exe`

2. **Upload Files**:

   - Click "Browse File" button under either "Attendance Sheet" or "Payroll Sheet"
   - Select your Excel file (.xlsx or .xls)
   - The selected file path will be displayed

3. **Process Files**:

   - Click "Process Attendance Sheet" or "Process Payroll Sheet" button
   - View the results in the results area below

4. **View Results**:

   - File information (name, rows, columns)
   - Column names and data types
   - Sample data preview
   - Any error messages if processing fails

5. **Clear Results**:
   - Click "Clear Results" button to clear the results area

## What the Application Currently Does

### Attendance Sheet Processing

- **Header Information Extraction**:
  - Date (from row 1, column B)
  - Company Name (from row 2, column B) 
  - Narration (from row 3, column B)
- **Employee Data Processing**:
  - Automatically finds where employee data starts
  - Extracts Employee Number, Employee Name, Attendance Type, and Attendance Days
  - Supports various attendance types (Present, Absent, Overtime @ rate, Half Day, etc.)
- **Format Validation**:
  - Validates expected attendance sheet structure
  - Shows clear error messages for invalid formats
  - Provides guidance on expected format structure
- **Data Export**: Export all extracted data to JSON format

### Payroll Sheet Processing

- **Basic File Analysis**:
  - File information (name, rows, columns)
  - Column names and data types
  - Sample data preview
  - Missing value analysis

### Expected Attendance Format

```
Row 1: [blank] | Date (e.g., 09-10-2025) | [blank] | [blank]
Row 2: [blank] | Company Name (e.g., LIGHT) | [blank] | [blank]  
Row 3: [blank] | Narration (e.g., Test attendance) | [blank] | [blank]
Row 4: [blank row]
Row 5: EMPL NO | EMPLOYEE NAME | Attendance/Production Types | Attendance Days
Row 6+: Employee data...
```

## Dependencies

- **pandas**: Excel file reading and data manipulation
- **openpyxl**: Excel file format support (.xlsx)
- **xlrd**: Legacy Excel file format support (.xls)
- **tkinter**: GUI framework (included with Python)
- **pyinstaller**: For creating standalone executable

## Troubleshooting

### Common Issues:

1. **"Module not found" errors**:

   - Make sure all dependencies are installed: `pip install -r requirements.txt`

2. **Excel file not opening**:

   - Ensure the file is a valid Excel format (.xlsx or .xls)
   - Check that the file is not corrupted or password-protected

3. **Executable not working**:
   - Try running from source first to identify any issues
   - Ensure all dependencies are properly included in the build

### Getting Help:

If you encounter any issues:

1. Check the results area for error messages
2. Ensure your Excel files are in the correct format
3. Verify all dependencies are installed correctly

## Future Enhancements

This is a foundation that can be extended with:

- Advanced data processing and analysis
- Data export functionality
- Database integration
- Custom report generation
- Data visualization features
- User configuration options

## Technical Details

- **Language**: Python 3.8+
- **GUI Framework**: tkinter with modern styling
- **Data Processing**: pandas library
- **File Support**: Excel (.xlsx, .xls)
- **Threading**: Non-blocking file processing
- **Packaging**: PyInstaller for Windows executable

## License

This project is created for internal use and data processing tasks.
