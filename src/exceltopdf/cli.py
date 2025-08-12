#!/usr/bin/env python3
"""CLI tool to convert Excel files to PDF with all columns fitting on one page per sheet."""
import argparse
import os
import sys
import platform
from pathlib import Path

def merge_pdfs_with_pypdf2(pdf_paths, output_path):
    """Merge multiple PDF files into one using PyPDF2."""
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        raise ImportError("PyPDF2 not available for PDF merging")
    
    writer = PdfWriter()
    
    for pdf_path in pdf_paths:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            writer.add_page(page)
    
    with open(output_path, 'wb') as output_file:
        writer.write(output_file)
    
    # Clean up temporary files
    for pdf_path in pdf_paths:
        try:
            os.remove(pdf_path)
        except OSError:
            pass

def convert_with_win32com(excel_path, pdf_path, all_sheets=False, verbose=False, log=None):
    """Convert Excel to PDF using win32com (Windows with Excel installed)."""
    try:
        import win32com.client as win32
        import tempfile
    except ImportError:
        raise ImportError("pywin32 not available")
    
    excel_path = Path(excel_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    
    if verbose and log:
        log(f"Using win32com to convert {excel_path} to {pdf_path}")
    elif verbose:
        print(f"Using win32com to convert {excel_path} to {pdf_path}")
    
    # Start Excel application
    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    
    try:
        # Open workbook
        wb = xl.Workbooks.Open(str(excel_path))
        
        total_sheets = len(wb.Worksheets)
        if verbose and log:
            log(f"Opened workbook with {total_sheets} worksheets")
        elif verbose:
            print(f"Opened workbook with {total_sheets} worksheets")
        
        if all_sheets and total_sheets > 1:
            # Process all sheets - try to export to single PDF first
            try:
                # Configure each worksheet for fitting columns
                for ws in wb.Worksheets:
                    ws.Activate()
                    # Set page setup to fit all columns on one page
                    ws.PageSetup.FitToPagesWide = 1
                    ws.PageSetup.FitToPagesTall = False
                    ws.PageSetup.Zoom = False
                    if verbose and log:
                        log(f"Configured worksheet: {ws.Name}")
                    elif verbose:
                        print(f"Configured worksheet: {ws.Name}")
                
                # Try to export entire workbook to single PDF
                wb.ExportAsFixedFormat(0, str(pdf_path))  # 0 = xlTypePDF
                
                if verbose and log:
                    log("PDF export completed (all sheets in single file)")
                elif verbose:
                    print("PDF export completed (all sheets in single file)")
                
            except Exception as e:
                # If single PDF export fails, export each sheet separately and merge
                if verbose and log:
                    log(f"Single PDF export failed: {e}. Trying sheet-by-sheet approach.")
                elif verbose:
                    print(f"Single PDF export failed: {e}. Trying sheet-by-sheet approach.")
                
                temp_dir = tempfile.mkdtemp()
                temp_pdfs = []
                
                for i, ws in enumerate(wb.Worksheets, 1):
                    ws.Activate()
                    # Set page setup to fit all columns on one page
                    ws.PageSetup.FitToPagesWide = 1
                    ws.PageSetup.FitToPagesTall = False
                    ws.PageSetup.Zoom = False
                    
                    temp_pdf = os.path.join(temp_dir, f"sheet_{i:03d}_{ws.Name.replace('/', '_')}.pdf")
                    # Export only current sheet
                    ws.ExportAsFixedFormat(0, temp_pdf, From=1, To=1)
                    temp_pdfs.append(temp_pdf)
                    
                    if verbose and log:
                        log(f"Exported sheet {i}: {ws.Name}")
                    elif verbose:
                        print(f"Exported sheet {i}: {ws.Name}")
                
                # Merge all temp PDFs
                merge_pdfs_with_pypdf2(temp_pdfs, pdf_path)
                
                # Clean up temp directory
                try:
                    os.rmdir(temp_dir)
                except OSError:
                    pass
                
                if verbose and log:
                    log("PDF export completed (merged from individual sheets)")
                elif verbose:
                    print("PDF export completed (merged from individual sheets)")
        else:
            # Process single sheet or default behavior
            # Configure each worksheet for fitting columns (in case there are multiple)
            for ws in wb.Worksheets:
                ws.Activate()
                # Set page setup to fit all columns on one page
                ws.PageSetup.FitToPagesWide = 1
                ws.PageSetup.FitToPagesTall = False
                ws.PageSetup.Zoom = False
                if verbose and log:
                    log(f"Configured worksheet: {ws.Name}")
                elif verbose:
                    print(f"Configured worksheet: {ws.Name}")
            
            # Export to PDF (default behavior - all sheets in workbook)
            wb.ExportAsFixedFormat(0, str(pdf_path))  # 0 = xlTypePDF
            
            if verbose and log:
                log("PDF export completed")
            elif verbose:
                print("PDF export completed")
        
    finally:
        wb.Close()
        xl.Quit()

def convert_with_pandas_reportlab(excel_path, pdf_path, all_sheets=False, verbose=False, log=None):
    """Convert Excel to PDF using pandas and reportlab (fallback method)."""
    try:
        import pandas as pd
        from reportlab.lib.pagesizes import letter, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib import colors
        from reportlab.lib.units import inch
    except ImportError as e:
        raise ImportError(f"Required packages not available: {e}")
    
    if verbose and log:
        log(f"Using pandas+reportlab to convert {excel_path} to {pdf_path}")
    elif verbose:
        print(f"Using pandas+reportlab to convert {excel_path} to {pdf_path}")
    
    # Read Excel file
    excel_file = pd.ExcelFile(excel_path)
    
    # Create PDF document
    doc = SimpleDocTemplate(str(pdf_path), pagesize=landscape(letter))
    story = []
    styles = getSampleStyleSheet()
    
    sheet_names = excel_file.sheet_names
    if verbose and log:
        log(f"Found {len(sheet_names)} sheets: {', '.join(sheet_names)}")
    elif verbose:
        print(f"Found {len(sheet_names)} sheets: {', '.join(sheet_names)}")
    
    # Process sheets based on the all_sheets parameter
    sheets_to_process = sheet_names if all_sheets else sheet_names[:1] if sheet_names else []
    
    for i, sheet_name in enumerate(sheets_to_process):
        # Read sheet
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        if verbose and log:
            log(f"Processing sheet '{sheet_name}' with {len(df)} rows and {len(df.columns)} columns")
        elif verbose:
            print(f"Processing sheet '{sheet_name}' with {len(df)} rows and {len(df.columns)} columns")
        
        # Add sheet title (only if processing multiple sheets)
        if len(sheets_to_process) > 1:
            title = Paragraph(f"{sheet_name}", styles['Heading2'])
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
        
        # Auto-size columns to fit page width (fit all columns to one page width)
        available_width = landscape(letter)[0] - 2 * inch
        col_count = len(df.columns)
        col_width = available_width / col_count if col_count > 0 else 1 * inch
        
        # Set column widths
        table._argW = [col_width] * col_count
        
        story.append(table)
        
        # Add page break between sheets (except for the last sheet)
        if i < len(sheets_to_process) - 1:
            story.append(PageBreak())
        else:
            story.append(Spacer(1, 24))
    
    # Build PDF
    if verbose and log:
        log("Building PDF document")
    elif verbose:
        print("Building PDF document")
    
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
        "--all-sheets",
        action="store_true",
        help="Convert all sheets in Excel file to single PDF (default: False)"
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
        print(f"Processing all sheets: {args.all_sheets}")
    
    try:
        if method == "win32com":
            convert_with_win32com(input_path, output_path, all_sheets=args.all_sheets, verbose=args.verbose)
        else:
            convert_with_pandas_reportlab(input_path, output_path, all_sheets=args.all_sheets, verbose=args.verbose)
        
        if args.verbose:
            print(f"Successfully converted to '{output_path}'")
        else:
            print(f"Converted: {output_path}")
            
    except ImportError as e:
        print(f"Error: {e}", file=sys.stderr)
        print("Please install required dependencies:", file=sys.stderr)
        if "win32com" in str(e):
            print("  pip install pywin32", file=sys.stderr)
        elif "PyPDF2" in str(e):
            print("  pip install PyPDF2", file=sys.stderr)
        else:
            print("  pip install pandas openpyxl reportlab PyPDF2", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error during conversion: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
