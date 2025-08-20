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
        self.root.geometry("800x700")  # Larger initial size for better layout
        
        # Configure styles for better appearance
        self.setup_styles()
        
        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.method = tk.StringVar(value="auto")
        self.verbose = tk.BooleanVar()
        self.all_sheets = tk.BooleanVar()
        self.auto_adjust = tk.BooleanVar(value=True)  # New option for auto-adjusting cell dimensions
        self.aggressive_adjust = tk.BooleanVar(value=True)  # New option for aggressive adjustment - now default
        
        self.setup_ui()
        
    def setup_styles(self):
        """Configure custom styles for better appearance."""
        style = ttk.Style()
        
        # Configure Accent.TButton style
        style.configure("Accent.TButton", 
                       background="#0078d4", 
                       foreground="white",
                       borderwidth=0,
                       focuscolor="none")
        
        # Configure LabelFrame style
        style.configure("TLabelframe", 
                       borderwidth=2, 
                       relief="groove",
                       background="#f0f0f0")
        
        # Configure LabelFrame.Label style
        style.configure("TLabelframe.Label", 
                       font=("TkDefaultFont", 10, "bold"),
                       foreground="#2c2c2c")
        
    def setup_ui(self):
        # Main frame with better padding and responsive design
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Configure grid weights for responsive design
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Input file selection with better layout
        ttk.Label(main_frame, text="Input Excel File:", font=("TkDefaultFont", 10, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 8))
        
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=0, column=1, columnspan=2, sticky="ew", pady=(0, 8))
        input_frame.columnconfigure(0, weight=1)
        
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_file, font=("TkDefaultFont", 9))
        self.input_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        
        # Bind input file changes to auto-generate output filename
        self.input_file.trace_add("write", self.on_input_file_changed)
        
        ttk.Button(input_frame, text="Browse...", command=self.browse_input_file, 
                  style="Accent.TButton").grid(row=0, column=1)
        
        # Output file selection with better layout
        ttk.Label(main_frame, text="Output PDF File:", font=("TkDefaultFont", 10, "bold")).grid(
            row=1, column=0, sticky="w", pady=(15, 8))
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=1, column=1, columnspan=2, sticky="ew", pady=(15, 8))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_file, font=("TkDefaultFont", 9))
        self.output_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        
        ttk.Button(output_frame, text="Browse...", command=self.browse_output_file, 
                  style="Accent.TButton").grid(row=0, column=1)
        
        # Info label about auto-generation with better positioning
        ttk.Label(main_frame, text="(Nome gerado automaticamente)", 
                 font=("TkDefaultFont", 8), foreground="gray").grid(
                     row=1, column=2, sticky="w", padx=(8, 0), pady=(15, 8))
        
        # Method selection with better spacing
        ttk.Label(main_frame, text="Conversion Method:", font=("TkDefaultFont", 10, "bold")).grid(
            row=2, column=0, sticky="w", pady=(15, 8))
        
        method_frame = ttk.Frame(main_frame)
        method_frame.grid(row=2, column=1, columnspan=2, sticky="w", pady=(15, 8))
        
        method_combo = ttk.Combobox(method_frame, textvariable=self.method, state="readonly", width=25)
        method_combo['values'] = ('auto', 'excel', 'reportlab')
        method_combo.grid(row=0, column=0, sticky="w")
        
        # Options frame with better responsive layout
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(15, 10))
        options_frame.columnconfigure(0, weight=1)
        options_frame.columnconfigure(1, weight=1)
        options_frame.columnconfigure(2, weight=1)
        options_frame.columnconfigure(3, weight=1)
        
        # First row of options
        row1_frame = ttk.Frame(options_frame)
        row1_frame.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 10))
        row1_frame.columnconfigure(0, weight=1)
        row1_frame.columnconfigure(1, weight=1)
        row1_frame.columnconfigure(2, weight=1)
        row1_frame.columnconfigure(3, weight=1)
        
        # Verbose checkbox
        verbose_frame = ttk.Frame(row1_frame)
        verbose_frame.grid(row=0, column=0, sticky="w", padx=(0, 15))
        ttk.Checkbutton(verbose_frame, text="Verbose output", variable=self.verbose).pack(side="left")
        
        # All sheets checkbox
        all_sheets_frame = ttk.Frame(row1_frame)
        all_sheets_frame.grid(row=0, column=1, sticky="w", padx=(0, 15))
        ttk.Checkbutton(all_sheets_frame, text="Converter todas as abas", variable=self.all_sheets).pack(side="left")
        
        # Auto-adjust cell dimensions checkbox
        auto_adjust_frame = ttk.Frame(row1_frame)
        auto_adjust_frame.grid(row=0, column=2, sticky="w", padx=(0, 15))
        ttk.Checkbutton(auto_adjust_frame, text="Ajustar dimensões das células", variable=self.auto_adjust).pack(side="left")
        
        # Aggressive adjust checkbox
        aggressive_frame = ttk.Frame(row1_frame)
        aggressive_frame.grid(row=0, column=3, sticky="w")
        ttk.Checkbutton(aggressive_frame, text="Ajuste agressivo", variable=self.aggressive_adjust).pack(side="left")
        
        # Second row with info labels
        row2_frame = ttk.Frame(options_frame)
        row2_frame.grid(row=1, column=0, columnspan=4, sticky="ew")
        row2_frame.columnconfigure(0, weight=1)
        row2_frame.columnconfigure(1, weight=1)
        row2_frame.columnconfigure(2, weight=1)
        row2_frame.columnconfigure(3, weight=1)
        
        # Info labels about options with better spacing
        ttk.Label(row2_frame, text="(Evita corte de texto nas células)", 
                 font=("TkDefaultFont", 8), foreground="blue").grid(
                     row=0, column=2, sticky="w", padx=(0, 15))
        
        ttk.Label(row2_frame, text="(Força ajuste mais agressivo se ainda houver corte)", 
                 font=("TkDefaultFont", 8), foreground="red").grid(
                     row=0, column=3, sticky="w")
        
        # Convert button with better styling
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=(20, 15))
        
        self.convert_btn = ttk.Button(button_frame, text="Convert", command=self.start_conversion, 
                                     style="Accent.TButton", padding=(20, 10))
        self.convert_btn.pack()
        
        # Progress bar with better positioning
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        
        # Log area with better responsive design
        log_label_frame = ttk.Frame(main_frame)
        log_label_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        log_label_frame.columnconfigure(0, weight=1)
        
        ttk.Label(log_label_frame, text="Log:", font=("TkDefaultFont", 10, "bold")).grid(
            row=0, column=0, sticky="w")
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(7, weight=1)
        
        # Log text area with better font and responsive design
        self.log_text = tk.Text(log_frame, height=12, wrap=tk.WORD, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Configure minimum window size
        self.root.minsize(700, 600)
        
        # Bind window resize event
        self.root.bind('<Configure>', self.on_window_resize)
        
    def on_window_resize(self, event):
        """Handle window resize events for better responsive design."""
        if event.widget == self.root:
            # Adjust font sizes based on window width
            width = event.width
            if width < 800:
                # Small window - use smaller fonts
                font_size = 8
                title_font_size = 9
            elif width < 1000:
                # Medium window - use medium fonts
                font_size = 9
                title_font_size = 10
            else:
                # Large window - use larger fonts
                font_size = 10
                title_font_size = 11
            
            # Update log text font size
            self.log_text.configure(font=("Consolas", font_size))
            
            # Update entry field font sizes
            self.input_entry.configure(font=("TkDefaultFont", font_size))
            self.output_entry.configure(font=("TkDefaultFont", font_size))
    
    def browse_input_file(self):
        """Open file dialog to select input Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            # Auto-generate output filename
            self.auto_generate_output_filename(filename)
                
    def auto_generate_output_filename(self, input_filename):
        """Automatically generate output filename based on input filename."""
        base = os.path.splitext(input_filename)[0]
        output_filename = f"{base}.pdf"
        
        # Check if file exists and add number if necessary
        counter = 1
        while os.path.exists(output_filename):
            output_filename = f"{base}_{counter}.pdf"
            counter += 1
            
        self.output_file.set(output_filename)
        
    def on_input_file_changed(self, *args):
        """Called when input file path changes."""
        input_path = self.input_file.get()
        if input_path and os.path.exists(input_path):
            # Only auto-generate if output is empty or if user hasn't manually set it
            if not self.output_file.get() or self.output_file.get().endswith('.pdf'):
                self.auto_generate_output_filename(input_path)
        
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
            all_sheets = self.all_sheets.get()
            auto_adjust = self.auto_adjust.get() # Get the new option value
            aggressive_adjust = self.aggressive_adjust.get() # Get the new option value
            
            self.log_message(f"Starting conversion...")
            self.log_message(f"Input: {input_path}")
            self.log_message(f"Output: {output_path}")
            self.log_message(f"Method: {method}")
            self.log_message(f"Convert all sheets: {all_sheets}")
            self.log_message(f"Auto-adjust cell dimensions: {auto_adjust}")
            self.log_message(f"Aggressive adjust: {aggressive_adjust}")
            self.log_message("")
            
            if not os.path.exists(input_path):
                raise FileNotFoundError(f"Input file not found: {input_path}")
            
            # Choose conversion method
            if method == "auto":
                # Try to determine best method
                try:
                    import win32com.client
                    self.log_message("Auto-detected: Using Excel (win32com) method")
                    convert_with_win32com(input_path, output_path, all_sheets=all_sheets, verbose=verbose, log=self.log_message, auto_adjust=auto_adjust, aggressive_adjust=aggressive_adjust)
                except ImportError:
                    self.log_message("Auto-detected: Using ReportLab (pandas) method")
                    convert_with_pandas_reportlab(input_path, output_path, all_sheets=all_sheets, verbose=verbose, log=self.log_message, auto_adjust=auto_adjust, aggressive_adjust=aggressive_adjust)
            elif method == "excel":
                self.log_message("Using Excel (win32com) method")
                convert_with_win32com(input_path, output_path, all_sheets=all_sheets, verbose=verbose, log=self.log_message, auto_adjust=auto_adjust, aggressive_adjust=aggressive_adjust)
            elif method == "reportlab":
                self.log_message("Using ReportLab (pandas) method")
                convert_with_pandas_reportlab(input_path, output_path, all_sheets=all_sheets, verbose=verbose, log=self.log_message, auto_adjust=auto_adjust, aggressive_adjust=aggressive_adjust)
                
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
