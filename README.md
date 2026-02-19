# SCADA Signal Synchronizer

A Python-based GUI application for synchronizing SCADA signal descriptions across Excel workbooks. This tool reads descriptions from area sheets and intelligently updates the SCADA_SIGNAL sheet with properly formatted descriptions for different signal variants.

## Features

- **Excel File Processing**: Load and process SCADA signal data from Excel workbooks
- **Smart Description Synchronization**: Automatically synchronize descriptions from area sheets to SCADA_SIGNAL sheet
- **Intelligent Formatting**: Format variant names (CamelCase to readable text) like `HiAlarm` → `HIGH ALARM`
- **Custom Output**: Save synchronized files to a user-selected location with custom naming
- **Processing Log**: Real-time logging of all operations with detailed status messages
- **User-Friendly GUI**: Modern tkinter-based interface with professional styling

## How It Works

### Step 1: Read Descriptions from Area Sheets
The application scans all non-SCADA_SIGNAL sheets (AC, BILGES, etc.) to extract tag names and their descriptions.

### Step 2: Synchronize SCADA_SIGNAL Sheet
Updates the SCADA_SIGNAL sheet by:
- Matching base tag names from area sheets to SCADA_SIGNAL entries
- Appending properly formatted variant names to descriptions (e.g., "Status" stays as-is, "HiAlarm" becomes "HIGH ALARM")
- Preserving data-type-only variants without adding suffixes
- Never adding ".Status" suffix to Status variants

### Example
- **Area Sheet Description**: "Tank Level"
- **Variant**: HiAlarm
- **Result**: "Tank Level HIGH ALARM"

## Requirements

See `requirements.txt` for Python package dependencies.

**System Requirements:**
- Python 3.7+
- Windows/Mac/Linux
- Excel files in .xlsx or .xls format

## Installation

1. Clone or download this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python ScadaSignalSynchronizer.py
   ```

2. **Select Excel File**: Click "Select Excel File" and choose your SCADA configuration workbook

3. **Configure Settings**:
   - **Update tag names** (toggleable): Enable/disable tag name updating
   - Smart descriptions are always enabled

4. **Process File**:
   - Click "▶️ Synchronize SCADA_SIGNAL"
   - Select output folder and filename in the dialog
   - Click "OK" to begin synchronization

5. **View Results**: The output folder opens automatically when complete

## File Structure

```
ScadaSignalSynchronizer.py  - Main application
requirements.txt             - Python dependencies
README.md                    - This file
```

## Excel Sheet Structure

### Expected Columns (Area Sheets)
- **Tag Name**: The base tag identifier
- **Description**: The description text
- **UDT Type** (optional): Data type information

### Expected Columns (SCADA_SIGNAL Sheet)
- **DB**: Database/Area name
- **Scada Tag Path**: Full path (format: DB.TAG.VARIANT)
- **Description**: Description field to be synchronized

## Error Handling

The application includes comprehensive error handling:
- File validation before processing
- Column existence checking
- Detailed error messages in the UI
- Processing continues gracefully if optional columns are missing

## Logging

All operations are logged to the in-app "Processing Log" panel, including:
- Sheet scanning results
- Row update counts
- Processing time for each step
- Completion status and output file location

## Settings

- **Update tag names**: Toggle whether to update tag names during synchronization
- **Smart descriptions**: Always enabled - formats variant names intelligently without data-type suffixes

## Output

The synchronized Excel file is saved to your selected location with the custom filename you provide.

## License

[Add your license information here]

## Support

For issues or questions, please contact the development team.
