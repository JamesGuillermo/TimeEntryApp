# Time Entry Application

A Python GUI application for entering digit-based time values, converting them to HH:MM format, and saving to Excel files.

## Features

### GUI Components
- **Date Field**: Persistent date entry (doesn't reset between submissions)
- **ID Number Field**: Entry for identification numbers
- **7 Step Time Inputs**: Each with digit entry, AM/PM selection, and real-time conversion display
- **File Operations**: Create, Load, and Submit buttons for Excel file management

### Time Conversion Examples
- `62` → `06:02`
- `634` → `06:34`
- `101` → `10:01`
- `1021` → `10:21`
- `121` → `01:21`
- `1216` → `12:16`

### Excel Structure
- Column A: Date
- Column B: ID Number
- Columns C-I: Step 1 through Step 7

## Installation

1. **Install Python Dependencies**:
   ```powershell
   pip install -r requirements.txt
   ```

2. **Run the Application**:
   ```powershell
   python time_entry_app.py
   ```

## Usage Instructions

### First Time Setup
1. **Create Excel File**: Click "Create Excel File" to create a new `.xlsx` file
2. **Or Load Existing**: Click "Load Excel File" to open an existing file

### Entering Data
1. **Set Date**: Enter the date (persists until manually changed)
2. **Enter ID**: Input the ID number for this batch
3. **Fill Steps 1-7**:
   - Enter digits in the time field (e.g., 634, 1021)
   - Select AM or PM (default is AM)
   - View the converted time for verification
4. **Submit**: Click "Submit Entry" to save to Excel

### Button Functions
- **Create Excel File**: Creates new Excel file with proper headers
- **Load Excel File**: Opens existing Excel file for data entry
- **Submit Entry**: Saves current batch to Excel (appends to last row)
- **Clear All**: Resets all fields to defaults

## Time Conversion Logic

The application converts digit inputs using these rules:

1. **1 digit** (e.g., `5`): `05:00`
2. **2 digits** (e.g., `62`):
   - If ≤ 59: `00:62` → Invalid, converts to `62:00`
   - If > 59: `62:00`
3. **3 digits** (e.g., `634`): `6:34` → `06:34`
4. **4 digits** (e.g., `1021`): `10:21`

AM/PM selection affects the final 12-hour format display.

## Error Handling

- **Invalid Digits**: Shows "Invalid" for non-numeric input
- **Invalid Times**: Shows "Invalid" for impossible times (e.g., 25:70)
- **File Errors**: User-friendly error messages for file operations
- **Missing Data**: Warnings for required fields

## File Structure

```
TimeEntryApp/
├── time_entry_app.py      # Main application file
├── requirements.txt       # Python dependencies
├── README.md             # This file
└── Plan.txt              # Original requirements
```

## Technical Details

- **GUI Framework**: ttkbootstrap (modern tkinter styling)
- **Excel Handling**: pandas + openpyxl
- **Theme**: Cosmo (clean, modern appearance)
- **Window**: 800x700 pixels, resizable, auto-centered

## Troubleshooting

### Common Issues

1. **"Module not found" errors**: Install dependencies with `pip install -r requirements.txt`
2. **Excel file won't open**: Ensure file isn't open in Excel when loading
3. **Time conversion shows "Invalid"**: Check that input contains only digits

### Dependencies
- Python 3.7+
- ttkbootstrap 1.10.1+
- pandas 1.5.0+
- openpyxl 3.0.10+

## Version History

- **v1.0**: Initial release with all planned features
  - 7-step time entry with AM/PM selection
  - Real-time time conversion and display
  - Excel file creation, loading, and data appending
  - Modern GUI with ttkbootstrap styling
  - Comprehensive error handling and validation
