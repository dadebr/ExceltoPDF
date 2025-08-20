# ExceltoPDF

[🇧🇷 Português](README.md) | [🇺🇸 English](README_EN.md)

A tool with graphical and command-line interface for converting Excel files to PDF with optimized formatting, ensuring all columns fit on one page per worksheet.

> **⚠️ Important Notice:** This package is not yet available on PyPI. To install, clone the repository and install from source code.

## Features

• **Smart Column Adjustment**: Automatically adjusts all columns to fit the width of one page per worksheet
• **Multiple Conversion Methods**:
  • Windows with Excel installed: Uses win32com for native Excel to PDF export
  • Cross-platform alternative: Uses pandas + reportlab for universal compatibility
• **Batch Processing**: Processes multiple worksheets in a single Excel file
• **Command Line Interface**: Easy to use in terminal or scripts
• **Graphical Interface**: User-friendly GUI with file selectors and options
• **Flexible Output**: Maintains data integrity while optimizing layout

## Installation

### From Source Code (recommended)

```bash
git clone https://github.com/dadebr/ExceltoPDF.git
cd ExceltoPDF
pip install -e .
```

### Dependencies

The tool will automatically use the best available method:

For Windows with Microsoft Excel:
```bash
pip install pywin32
```

For cross-platform compatibility:
```bash
pip install pandas openpyxl reportlab
```

All dependencies are listed in requirements.txt and will be installed automatically.

## Usage

### Graphical Interface

To run the graphical interface:
```bash
exceltopdf-gui
```

Alternatively, if the command is not available:
```bash
python -m exceltopdf.gui
```

#### Graphical Interface Features

• **File Selection**: Navigation buttons to choose input Excel files and output PDF
• **Conversion Methods**: Dropdown menu with options:
  • auto - Automatically detects the best method
  • excel - Uses native Excel (Windows)
  • reportlab - Uses pandas + reportlab (cross-platform)
• **Output Options**: Checkbox to enable verbose output
• **Convert All Sheets**: "Convert all sheets" checkbox to process all worksheets into a single PDF
• **Log Area**: Shows conversion progress and details in real-time
• **Progress Bar**: Visual indicator during the conversion process

#### How to Use the Graphical Interface

1. Run `exceltopdf-gui` in the terminal
2. Click "Browse..." next to "Input Excel File" to select your Excel file
3. Click "Browse..." next to "Output PDF File" to choose where to save the PDF
4. Select the desired conversion method from the dropdown menu
5. Check "Verbose output" if you want detailed information
6. Check "Convert all sheets" if you want to process all worksheets
7. Click "Convert" to start the conversion
8. Monitor progress in the log area

The interface runs in a separate thread to prevent freezing during conversion and displays success or error messages at the end of the process.

### Command Line Interface

#### Basic Usage

```bash
# Convert Excel file to PDF
exceltopdf input.xlsx output.pdf

# With verbose output
exceltopdf input.xlsx output.pdf --verbose
```

#### Advanced Options

```bash
# Force specific conversion method
exceltopdf input.xlsx output.pdf --method win32com
exceltopdf input.xlsx output.pdf --method pandas

# Automatically detect the best method (default)
exceltopdf input.xlsx output.pdf --method auto

# Convert all sheets from Excel file to a single PDF
exceltopdf input.xlsx output.pdf --all-sheets

# Combine options
exceltopdf input.xlsx output.pdf --all-sheets --verbose --method auto
```

### Python API

```python
from exceltopdf.cli import convert_with_pandas_reportlab, convert_with_win32com

# Using pandas/reportlab (cross-platform)
convert_with_pandas_reportlab('input.xlsx', 'output.pdf')

# Using win32com (Windows + Excel only)
convert_with_win32com('input.xlsx', 'output.pdf')
```

## Supported Formats

• Input: .xlsx, .xls
• Output: .pdf

## How It Works

### Method 1: win32com (Windows + Excel)

• Uses Microsoft Excel's built-in PDF export functionality
• Configures page setup to fit all columns on one page
• Provides highest quality output with native formatting
• Automatically applies scaling to ensure columns fit

### Method 2: pandas + reportlab (Cross-platform)

• Reads Excel data using pandas
• Converts to PDF using reportlab
• Automatically calculates column widths to fit the page
• Works on any platform without Excel installed

## Examples

```bash
# Simple conversion
exceltopdf sales_report.xlsx sales_report.pdf

# Conversion with detailed logging
exceltopdf financial_data.xlsx financial_data.pdf -v

# Force cross-platform method
exceltopdf data.xlsx output.pdf --method pandas

# Convert all sheets to a single PDF
exceltopdf workbook.xlsx complete_report.pdf --all-sheets
```

## Development

### Setting Up Development Environment

```bash
git clone https://github.com/dadebr/ExceltoPDF.git
cd ExceltoPDF
pip install -e .[dev]
```

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=exceltopdf

# Run specific test file
pytest tests/test_basic.py
```

### Building Package

```bash
# Build distribution packages
python -m build

# Upload to PyPI (maintainers only)
twine upload dist/*
```

## Contributing

1. Fork the repository
2. Create a branch for your feature (git checkout -b feature/amazing-feature)
3. Commit your changes (git commit -m 'Add amazing feature')
4. Push to the branch (git push origin feature/amazing-feature)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Changelog

### v0.1.0

• Initial release
• Basic Excel to PDF conversion
• Cross-platform compatibility
• Command line interface
• Automatic column adjustment

## Troubleshooting

### Common Issues

**"Failed to import win32com"**
• Install pywin32: `pip install pywin32`
• Or use the pandas method: `--method pandas`

**"Required packages not available"**
• Install dependencies: `pip install pandas openpyxl reportlab`

**"Input file does not exist"**
• Check the file path and make sure the file exists
• Use absolute paths if necessary

**PDF output is cut off**
• The tool automatically adjusts columns, but very wide spreadsheets may need manual adjustment
• Consider using landscape orientation in the source Excel file

## Support

If you encounter issues or have questions:

1. Check the [troubleshooting section](#troubleshooting)
2. Search [existing issues](https://github.com/dadebr/ExceltoPDF/issues)
3. Create a [new issue](https://github.com/dadebr/ExceltoPDF/issues/new) with details about your problem

Made with ❤️ to make Excel to PDF conversion easier
