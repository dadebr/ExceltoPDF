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

def convert_with_win32com(excel_path, pdf_path, all_sheets=False, verbose=False, log=None, auto_adjust=True, aggressive_adjust=False):
    """Convert Excel to PDF using win32com (Windows with Excel installed)."""
    try:
        import win32com.client as win32
        import tempfile
        import time
    except ImportError:
        raise ImportError("pywin32 not available")
    
    excel_path = Path(excel_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    
    if verbose and log:
        log(f"Using win32com to convert {excel_path} to {pdf_path}")
        log(f"Auto-adjust cell dimensions: {auto_adjust}")
        log(f"Aggressive adjustment: {aggressive_adjust}")
    elif verbose:
        print(f"Using win32com to convert {excel_path} to {pdf_path}")
        print(f"Auto-adjust cell dimensions: {auto_adjust}")
        print(f"Aggressive adjustment: {aggressive_adjust}")
    
    # Start Excel application with timeout protection
    start_time = time.time()
    timeout = 300  # 5 minutes timeout
    
    try:
        xl = win32.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        
        # Check timeout
        if time.time() - start_time > timeout:
            raise TimeoutError("Excel application startup timeout")
        
        # Open workbook with timeout check
        if verbose and log:
            log("Opening Excel workbook...")
        elif verbose:
            print("Opening Excel workbook...")
            
        wb = xl.Workbooks.Open(str(excel_path))
        
        # Check timeout
        if time.time() - start_time > timeout:
            raise TimeoutError("Workbook opening timeout")
        
        total_sheets = len(wb.Worksheets)
        if verbose and log:
            log(f"Opened workbook with {total_sheets} worksheets")
        elif verbose:
            print(f"Opened workbook with {total_sheets} worksheets")
        
        # Function to optimize worksheet layout
        def optimize_worksheet_layout(worksheet):
            """Optimize worksheet layout to prevent text cutting."""
            try:
                # Check timeout
                if time.time() - start_time > timeout:
                    raise TimeoutError("Worksheet optimization timeout")
                
                worksheet.Activate()
                
                # Get used range
                used_range = worksheet.UsedRange
                if used_range is None:
                    return
                
                if auto_adjust:
                    # AutoFit columns based on content
                    if verbose and log:
                        log(f"  Auto-fitting columns for worksheet: {worksheet.Name}")
                    elif verbose:
                        print(f"  Auto-fitting columns for worksheet: {worksheet.Name}")
                    
                    # First, enable text wrapping for all cells to prevent cutting
                    used_range.WrapText = True
                    
                    # Simple AutoFit approach to avoid loops
                    try:
                        # Check timeout before processing
                        if time.time() - start_time > timeout:
                            raise TimeoutError("AutoFit processing timeout")
                        
                        # Get the actual used range dimensions
                        last_row = used_range.Rows.Count
                        last_col = used_range.Columns.Count
                        
                        if verbose and log:
                            log(f"    Worksheet dimensions: {last_row} rows x {last_col} columns")
                        elif verbose:
                            print(f"    Worksheet dimensions: {last_row} rows x {last_col} columns")
                        
                        # Apply AutoFit to all columns at once (more efficient)
                        used_range.Columns.AutoFit()
                        
                        # Check timeout
                        if time.time() - start_time > timeout:
                            raise TimeoutError("Column AutoFit timeout")
                        
                        # Apply AutoFit to all rows at once (more efficient)
                        used_range.Rows.AutoFit()
                        
                        # Check timeout
                        if time.time() - start_time > timeout:
                            raise TimeoutError("Row AutoFit timeout")
                        
                        # Apply aggressive adjustment if enabled
                        if aggressive_adjust:
                            if verbose and log:
                                log(f"    Applying aggressive adjustment...")
                            elif verbose:
                                print(f"    Applying aggressive adjustment...")
                            
                            # Force minimum dimensions for better readability (limited scope)
                            for col_idx in range(1, min(last_col + 1, 10)):  # Limit to first 10 columns
                                try:
                                    # Check timeout every few operations
                                    if col_idx % 5 == 0 and time.time() - start_time > timeout:
                                        raise TimeoutError("Aggressive column adjustment timeout")
                                        
                                    col = used_range.Columns(col_idx)
                                    if col.ColumnWidth < 12:
                                        col.ColumnWidth = 12
                                    elif col.ColumnWidth > 50:
                                        col.ColumnWidth = 50
                                except:
                                    pass  # Skip problematic columns
                            
                            for row_idx in range(1, min(last_row + 1, 50)):  # Limit to first 50 rows
                                try:
                                    # Check timeout every few operations
                                    if row_idx % 10 == 0 and time.time() - start_time > timeout:
                                        raise TimeoutError("Aggressive row adjustment timeout")
                                        
                                    row = used_range.Rows(row_idx)
                                    if row.RowHeight < 20:
                                        row.RowHeight = 20
                                    elif row.RowHeight > 100:
                                        row.RowHeight = 100
                                except:
                                    pass  # Skip problematic rows
                    
                    except Exception as e:
                        if verbose and log:
                            log(f"    Warning: AutoFit failed: {e}")
                        elif verbose:
                            print(f"    Warning: AutoFit failed: {e}")
                    
                    # Basic formatting to prevent text cutting
                    try:
                        # Check timeout
                        if time.time() - start_time > timeout:
                            raise TimeoutError("Formatting timeout")
                        
                        # Set alignment for better text wrapping
                        used_range.HorizontalAlignment = -4131  # xlLeft
                        used_range.VerticalAlignment = -4160    # xlTop
                        
                        # Ensure text wrapping is enabled
                        used_range.WrapText = True
                        
                        # Set reasonable font size
                        if aggressive_adjust:
                            used_range.Font.Size = 11
                        else:
                            used_range.Font.Size = 10
                    
                    except Exception as e:
                        if verbose and log:
                            log(f"    Warning: Basic formatting failed: {e}")
                        elif verbose:
                            print(f"    Warning: Basic formatting failed: {e}")
                    
                    if verbose and log:
                        log(f"    Applied {'aggressive' if aggressive_adjust else 'basic'} auto-adjustment for worksheet: {worksheet.Name}")
                    elif verbose:
                        print(f"    Applied {'aggressive' if aggressive_adjust else 'basic'} auto-adjustment for worksheet: {worksheet.Name}")
                        
                else:
                    if verbose and log:
                        log(f"  Skipping auto-adjustment for worksheet: {worksheet.Name}")
                    elif verbose:
                        print(f"  Skipping auto-adjustment for worksheet: {worksheet.Name}")
                
                # Set page setup for optimal PDF export
                try:
                    # Check timeout
                    if time.time() - start_time > timeout:
                        raise TimeoutError("Page setup timeout")
                    
                    page_setup = worksheet.PageSetup
                    
                    # Set orientation to landscape for better column fitting
                    page_setup.Orientation = 2  # xlLandscape
                    
                    # Fit all columns to one page width
                    page_setup.FitToPagesWide = 1
                    page_setup.FitToPagesTall = False
                    page_setup.Zoom = False
                    
                    # Use reasonable margins
                    page_setup.LeftMargin = 0.5 * 72   # 0.5 inch in points
                    page_setup.RightMargin = 0.5 * 72
                    page_setup.TopMargin = 0.5 * 72
                    page_setup.BottomMargin = 0.5 * 72
                    
                    # Set header and footer margins
                    page_setup.HeaderMargin = 0.3 * 72
                    page_setup.FooterMargin = 0.3 * 72
                    
                    # Center horizontally
                    page_setup.CenterHorizontally = True
                    page_setup.CenterVertically = False
                    
                    # Set paper size to A4
                    page_setup.PaperSize = 7  # xlPaperA4
                    
                    if verbose and log:
                        log(f"  Optimized layout for worksheet: {worksheet.Name}")
                    elif verbose:
                        print(f"  Optimized layout for worksheet: {worksheet.Name}")
                        
                except Exception as e:
                    if verbose and log:
                        log(f"  Warning: Page setup failed: {e}")
                    elif verbose:
                        print(f"  Warning: Page setup failed: {e}")
                    
            except Exception as e:
                if verbose and log:
                    log(f"  Warning: Could not fully optimize worksheet {worksheet.Name}: {e}")
                elif verbose:
                    print(f"  Warning: Could not fully optimize worksheet {worksheet.Name}: {e}")
        
        if all_sheets and total_sheets > 1:
            # Process all sheets - try to export to single PDF first
            try:
                # Check timeout before starting
                if time.time() - start_time > timeout:
                    raise TimeoutError("Processing timeout before sheet optimization")
                
                # Optimize each worksheet layout
                for i, ws in enumerate(wb.Worksheets):
                    # Check timeout for each worksheet
                    if time.time() - start_time > timeout:
                        raise TimeoutError(f"Worksheet optimization timeout at sheet {i+1}")
                    
                    optimize_worksheet_layout(ws)
                
                # Check timeout before export
                if time.time() - start_time > timeout:
                    raise TimeoutError("Export timeout after optimization")
                
                # Try to export entire workbook to single PDF
                if verbose and log:
                    log("Exporting all sheets to single PDF...")
                elif verbose:
                    print("Exporting all sheets to single PDF...")
                
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
                    # Check timeout for each sheet
                    if time.time() - start_time > timeout:
                        raise TimeoutError(f"Sheet-by-sheet export timeout at sheet {i}")
                    
                    # Optimize worksheet layout
                    optimize_worksheet_layout(ws)
                    
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
            # Check timeout before starting
            if time.time() - start_time > timeout:
                raise TimeoutError("Processing timeout before single sheet optimization")
            
            # Optimize each worksheet layout
            for i, ws in enumerate(wb.Worksheets):
                # Check timeout for each worksheet
                if time.time() - start_time > timeout:
                    raise TimeoutError(f"Single sheet optimization timeout at sheet {i+1}")
                
                optimize_worksheet_layout(ws)
            
            # Check timeout before export
            if time.time() - start_time > timeout:
                raise TimeoutError("Single sheet export timeout after optimization")
            
            # Export to PDF (default behavior - all sheets in workbook)
            if verbose and log:
                log("Exporting to PDF...")
            elif verbose:
                print("Exporting to PDF...")
                
            wb.ExportAsFixedFormat(0, str(pdf_path))  # 0 = xlTypePDF
            
            if verbose and log:
                log("PDF export completed")
            elif verbose:
                print("PDF export completed")
        
    except TimeoutError as e:
        error_msg = f"Timeout error during conversion: {e}"
        if verbose and log:
            log(error_msg)
        elif verbose:
            print(error_msg)
        raise TimeoutError(error_msg)
    except Exception as e:
        error_msg = f"Error during conversion: {e}"
        if verbose and log:
            log(error_msg)
        elif verbose:
            print(error_msg)
        raise Exception(error_msg)
    finally:
        try:
            if 'wb' in locals():
                wb.Close()
            if 'xl' in locals():
                xl.Quit()
        except Exception as cleanup_error:
            if verbose and log:
                log(f"Warning: Error during cleanup: {cleanup_error}")
            elif verbose:
                print(f"Warning: Error during cleanup: {cleanup_error}")

def convert_with_pandas_reportlab(excel_path, pdf_path, all_sheets=False, verbose=False, log=None, auto_adjust=True, aggressive_adjust=False):
    """Convert Excel to PDF using pandas and reportlab (fallback method)."""
    try:
        import pandas as pd
        from reportlab.lib.pagesizes import letter, landscape, A4
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib import colors
        from reportlab.lib.units import inch, cm
        from reportlab.platypus.flowables import KeepTogether
    except ImportError as e:
        raise ImportError(f"Required packages not available: {e}")
    
    if verbose and log:
        log(f"Using pandas+reportlab to convert {excel_path} to {pdf_path}")
        log(f"Auto-adjust cell dimensions: {auto_adjust}")
        log(f"Aggressive adjustment: {aggressive_adjust}")
    elif verbose:
        print(f"Using pandas+reportlab to convert {excel_path} to {pdf_path}")
        print(f"Auto-adjust cell dimensions: {auto_adjust}")
        print(f"Aggressive adjustment: {aggressive_adjust}")
    
    # Read Excel file
    excel_file = pd.ExcelFile(excel_path)
    
    # Create PDF document with A4 landscape for better column fitting
    doc = SimpleDocTemplate(str(pdf_path), pagesize=landscape(A4))
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
        
        if auto_adjust:
            # Calculate optimal column widths based on content
            available_width = landscape(A4)[0] - 2 * cm  # A4 landscape width minus margins
            col_count = len(df.columns)
            
            # Calculate column widths based on content length with better algorithm
            col_widths = []
            for col_idx in range(col_count):
                # Get maximum content length in this column
                max_length = len(str(df.columns[col_idx])) if col_idx < len(df.columns) else 10
                
                # Check data in this column with more detailed analysis
                if col_idx < len(df.columns):
                    col_data = df.iloc[:, col_idx].astype(str)
                    if len(col_data) > 0:
                        # Get the longest content in this column
                        max_data_length = col_data.str.len().max()
                        max_length = max(max_length, max_data_length)
                        
                        # Check for cells with multiple lines (containing \n)
                        multiline_cells = col_data[col_data.str.contains('\n', na=False)]
                        if len(multiline_cells) > 0:
                            # For multiline cells, calculate width based on longest line
                            max_line_length = 0
                            for cell in multiline_cells:
                                lines = cell.split('\n')
                                for line in lines:
                                    max_line_length = max(max_line_length, len(line))
                            max_length = max(max_length, max_line_length)
                
                # Set minimum and maximum widths with better constraints
                min_width = 2 * cm  # Increased minimum width
                max_width = 6 * cm  # Increased maximum width for better readability
                
                # Calculate width based on content with better padding
                # Use a more generous formula to prevent text cutting
                content_width = max(min_width, min(max_width, (max_length + 4) * 0.4 * cm))
                col_widths.append(content_width)
            
            # Adjust total width to fit available space with better distribution
            total_width = sum(col_widths)
            if total_width > available_width:
                # Scale down proportionally but maintain minimum widths
                scale_factor = available_width / total_width
                # Ensure no column goes below minimum width
                scaled_widths = [max(2 * cm, w * scale_factor) for w in col_widths]
                # Redistribute remaining space
                remaining_space = available_width - sum(scaled_widths)
                if remaining_space > 0:
                    extra_per_col = remaining_space / col_count
                    col_widths = [w + extra_per_col for w in scaled_widths]
                else:
                    col_widths = scaled_widths
            elif total_width < available_width:
                # Distribute extra space evenly but maintain proportions
                extra_space = available_width - total_width
                extra_per_col = extra_space / col_count
                col_widths = [w + extra_per_col for w in col_widths]
            
            # Set column widths
            table._argW = col_widths
            
            if verbose and log:
                log(f"  Applied comprehensive auto-adjustment for sheet: {sheet_name}")
            elif verbose:
                print(f"  Applied comprehensive auto-adjustment for sheet: {sheet_name}")
        else:
            # Use equal column widths when auto-adjust is disabled
            available_width = landscape(A4)[0] - 2 * cm
            col_count = len(df.columns)
            col_width = available_width / col_count if col_count > 0 else 2 * cm
            table._argW = [col_width] * col_count
            
            if verbose and log:
                log(f"  Using equal column widths for sheet: {sheet_name}")
            elif verbose:
                print(f"  Using equal column widths for sheet: {sheet_name}")
        
        # Style the table with better formatting
        table.setStyle(TableStyle([
            # Header styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            
            # Data styling
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Changed to TOP for better text wrapping
            ('TOPPADDING', (0, 1), (-1, -1), 8),   # Increased padding
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8), # Increased padding
            ('LEFTPADDING', (0, 1), (-1, -1), 6),   # Increased padding
            ('RIGHTPADDING', (0, 1), (-1, -1), 6),  # Increased padding
            
            # Grid styling
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            
            # Text wrapping and overflow prevention
            ('WORDWRAP', (0, 0), (-1, -1), True),
            ('LEADING', (0, 0), (-1, -1), 12),  # Line spacing for better readability
        ]))
        
        # Wrap table in KeepTogether to prevent splitting across pages
        story.append(KeepTogether(table))
        
        # Add page break between sheets (except for the last sheet)
        if i < len(sheets_to_process) - 1:
            story.append(PageBreak())
        else:
            story.append(Spacer(1, 24))
    
    # Build PDF
    if verbose and log:
        log("Building PDF document with optimized column widths")
    elif verbose:
        print("Building PDF document with optimized column widths")
    
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
