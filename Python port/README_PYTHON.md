# TilikausiPullautin - Python Version

This is a Python port of the C# TilikausiPullautin Excel generation tool.

## Installation

1. Make sure you have Python 3.7 or later installed
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Run the Python script:

```bash
python pullautin.py
```

The program will:
1. Ask for the year to generate the accounting period for
2. Ask for the name of the person
3. Generate an Excel file with monthly sheets and a yearly summary
4. Save the file to `Documents/Vip-Hius/` folder

## Differences from C# Version

- Uses `openpyxl` library instead of EPPlus
- Uses Python's built-in `locale` module for Finnish date formatting
- File paths use Python's `pathlib` for cross-platform compatibility
- The behavior and output are identical to the C# version

## Requirements

- Python 3.7+
- openpyxl 3.1.2

## Notes

- The program tries to set the Finnish locale for proper month and day names
- If Finnish locale is not available on your system, it will fall back to the default locale with a warning
- All Excel formulas and formatting are preserved from the original C# version
