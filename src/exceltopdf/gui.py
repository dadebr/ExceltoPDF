#!/usr/bin/env python3
"""
Tkinter GUI for ExceltoPDF

A graphical user interface for converting Excel files to PDF with various options.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
from .cli import convert_with_pandas_reportlab, convert_with_win32com


class ExcelToPDFGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Converter")
        self.root.geometry("600x500")
        
        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.method = tk.StringVar(value="auto")
        self.verbose = tk.BooleanVar()
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Input file selection
        ttk.Label(main_frame, text="Input Excel File:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=0, column=1, columnspan=2, sticky="ew", pady=(0, 5))
        input_frame.columnconfigure(0, weight=1)
        
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_file)
        self.input_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        
        ttk.Button(input_frame, text="Browse...", command=self.browse_input_file).grid(row=0, column=1)
        
        # Output file selection
        ttk.Label(main_frame, text="Output PDF File:").grid(row=1, column=0, sticky="w", pady=(10, 5))
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=1, column=1, columnspan=2, sticky="ew", pady=(10, 5))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_file)
        self.output_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        
        ttk.Button(output_frame, text="Browse...", command=self.browse_output_file).grid(row=0, column=1)
        
        # Method selection
        ttk.Label(main_frame, text="Conversion Method:").grid(row=2, column=0, sticky="w", pady=(10, 5))
        
        method_frame = ttk.Frame(main_frame)
        method_frame.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(10, 5))
        
        method_combo = ttk.Combobox(method_frame, textvariable=self.method, state="readonly", width=20)
        method_combo['values'] = ('auto', 'excel', 'reportlab')
        method_combo.grid(row=0, column=0, sticky="w")
        
        # Verbose checkbox
        ttk.Checkbutton(main_frame, text="Verbose output", variable=self.verbose).grid(
            row=3, column=1, sticky="w", pady=(10, 5))
        
        # Convert button
        self.convert_btn = ttk.Button(main_frame, text="Convert", command=self.start_conversion)
        self.convert_btn.grid(row=4, column=1, pady=(20, 10))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        
        # Log area
        ttk.Label(main_frame, text="Log:").grid(row=6, column=0, sticky="nw", pady=(0, 5))
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=6, column=1, columnspan=2, sticky="nsew", pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
    def browse_input_file(self):
        """Open file dialog to select input Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            # Auto-generate output filename
            if not self.output_file.get():
                base = os.path.splitext(filename)[0]
                self.output_file.set(f"{base}.pdf")
                
    def browse_output_file(self):
        """Open file dialog to select output PDF file."""
        filename = filedialog.asksaveasfilename(
            title="Save PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
            
    def log_message(self, message):
        """Add message to log area."""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def start_conversion(self):
        """Start conversion in a separate thread."""
        if not self.input_file.get():
            messagebox.showerror("Error", "Please select an input Excel file.")
            return
            
        if not self.output_file.get():
            messagebox.showerror("Error", "Please select an output PDF file.")
            return
            
        # Disable convert button and start progress
        self.convert_btn.config(state="disabled")
        self.progress.start()
        self.log_text.delete(1.0, tk.END)
        
        # Start conversion in background thread
        thread = threading.Thread(target=self.convert_file)
        thread.daemon = True
        thread.start()
        
    def convert_file(self):
        """Convert Excel file to PDF."""
        try:
            input_path = self.input_file.get()
            output_path = self.output_file.get()
            method = self.method.get()
            verbose = self.verbose.get()
            
            self.log_message(f"Starting conversion...")
            self.log_message(f"Input: {input_path}")
            self.log_message(f"Output: {output_path}")
            self.log_message(f"Method: {method}")
            self.log_message("")
            
            if not os.path.exists(input_path):
                raise FileNotFoundError(f"Input file not found: {input_path}")
            
            # Choose conversion method
            if method == "auto":
                # Try to determine best method
                try:
                    import win32com.client
                    self.log_message("Auto-detected: Using Excel (win32com) method")
                    convert_with_win32com(input_path, output_path, verbose=verbose)
                except ImportError:
                    self.log_message("Auto-detected: Using ReportLab (pandas) method")
                    convert_with_pandas_reportlab(input_path, output_path, verbose=verbose)
            elif method == "excel":
                self.log_message("Using Excel (win32com) method")
                convert_with_win32com(input_path, output_path, verbose=verbose)
            elif method == "reportlab":
                self.log_message("Using ReportLab (pandas) method")
                convert_with_pandas_reportlab(input_path, output_path, verbose=verbose)
                
            self.log_message("")
            self.log_message("Conversion completed successfully!")
            
            # Show success message
            self.root.after(0, lambda: messagebox.showinfo(
                "Success", 
                f"Excel file converted successfully!\n\nOutput saved to:\n{output_path}"
            ))
            
        except Exception as e:
            error_msg = f"Error during conversion: {str(e)}"
            self.log_message("")
            self.log_message(error_msg)
            
            # Show error message
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            
        finally:
            # Re-enable convert button and stop progress
            self.root.after(0, self.conversion_finished)
            
    def conversion_finished(self):
        """Called when conversion is finished."""
        self.convert_btn.config(state="normal")
        self.progress.stop()


def main():
    """Main entry point for the GUI application."""
    root = tk.Tk()
    app = ExcelToPDFGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
