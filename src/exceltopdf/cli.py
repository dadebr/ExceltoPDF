#!/usr/bin/env python3
"""CLI tool to convert Excel files to PDF with all columns fitting on one page per sheet."""

import argparse
import os
import sys
import platform
from pathlib import Path


def convert_with_win32com(excel_path, pdf_path):
    """Convert Excel to PDF using win32com (Windows with Excel installed)."""
    try:
        import win32com.client as win32
    except ImportError:
        raise ImportError("pywin32 not available")
    
    excel_path = Path(excel_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    
    # Start Excel application
    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    
    try:
        # Open workbook
        wb = xl.Workbooks.Open(str(excel_path))
        
        # Configure each worksheet for fitting columns
        for ws in wb.Worksheets:
            ws.Activate()
            # Set page setup to fit all columns on one page
            ws.PageSetup.FitToPagesWide = 1
            ws.PageSetup.FitToPagesTall = False
            ws.PageSetup.Zoom = False
        
        # Export to PDF
        wb.ExportAsFixedFormat(0, str(pdf_path))  # 0 = xlTypePDF
        
    finally:
        wb.Close()
        xl.Quit()


def convert_with_pandas_reportlab(excel_path, pdf_path):
    """Convert Excel to PDF using pandas and reportlab (fallback method)."""
    try:
        import pandas as pd
        from reportlab.lib.pagesizes import letter, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib import colors
        from reportlab.lib.units import inch
    except ImportError as e:
        raise ImportError(f"Required packages not available: {e}")
    
    # Read Excel file
    excel_file = pd.ExcelFile(excel_path)
    
    # Create PDF document
    doc = SimpleDocTemplate(str(pdf_path), pagesize=landscape(letter))
    story = []
    styles = getSampleStyleSheet()
    
    for sheet_name in excel_file.sheet_names:
        # Read sheet
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # Add sheet title
        if len(excel_file.sheet_names) > 1:
            title = Paragraph(f"<b>{sheet_name}</b>", styles['Heading2'])
            story.append(title)
            story.append(Spacer(1, 12))
        
        # Convert DataFrame to list of lists for reportlab Table
        data = [df.columns.tolist()] + df.fillna('').astype(str).values.tolist()
        
        # Create table
        table = Table(data)
        
        # Style the table
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))
        
        # Auto-size columns to fit page width
        available_width = landscape(letter)[0] - 2 * inch
        col_count = len(df.columns)
        col_width = available_width / col_count if col_count > 0 else 1 * inch
        
        # Set column widths
        table._argW = [col_width] * col_count
        
        story.append(table)
        story.append(Spacer(1, 24))
    
    # Build PDF
    doc.build(story)


def main():
    """Main CLI function."""
    parser = argparse.ArgumentParser(
        description="Convert Excel files to PDF with all columns fitting on one page per sheet."
    )
    parser.add_argument("input", help="Input Excel file path (.xlsx, .xls)")
    parser.add_argument("output", help="Output PDF file path")
    parser.add_argument(
        "--method", 
        choices=["auto", "win32com", "pandas"],
        default="auto",
        help="Conversion method to use (default: auto)"
    )
    parser.add_argument(
        "--verbose", "-v", 
        action="store_true",
        help="Enable verbose output"
    )
    
    args = parser.parse_args()
    
    # Validate input file
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: Input file '{input_path}' does not exist.", file=sys.stderr)
        sys.exit(1)
    
    if input_path.suffix.lower() not in ['.xlsx', '.xls']:
        print(f"Error: Input file must be an Excel file (.xlsx or .xls).", file=sys.stderr)
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Determine conversion method
    method = args.method
    if method == "auto":
        if platform.system() == "Windows":
            method = "win32com"
        else:
            method = "pandas"
    
    if args.verbose:
        print(f"Converting '{input_path}' to '{output_path}' using method: {method}")
    
    try:
        if method == "win32com":
            convert_with_win32com(input_path, output_path)
        else:
            convert_with_pandas_reportlab(input_path, output_path)
        
        if args.verbose:
            print(f"Successfully converted to '{output_path}'")
        else:
            print(f"Converted: {output_path}")
            
    except ImportError as e:
        print(f"Error: {e}", file=sys.stderr)
        print("Please install required dependencies:", file=sys.stderr)
        if "win32com" in str(e):
            print("  pip install pywin32", file=sys.stderr)
        else:
            print("  pip install pandas openpyxl reportlab", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error during conversion: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
