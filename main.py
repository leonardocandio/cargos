import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import logging
from pathlib import Path
import traceback
from datetime import datetime

class FileGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Generator - Excel to PDF")
        self.root.geometry("800x600")
        
        # Initialize variables
        self.excel_file_path = tk.StringVar()
        self.templates_path = tk.StringVar(value="templates/")
        self.destination_path = tk.StringVar(value="output/")
        
        # Setup logging
        self.setup_logging()
        
        # Create UI
        self.create_ui()
        
        # Create default directories
        self.create_default_directories()
    
    def setup_logging(self):
        """Setup logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('app.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def create_ui(self):
        """Create the main UI with tabs"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create tabs
        self.cargos_frame = ttk.Frame(self.notebook)
        self.stock_frame = ttk.Frame(self.notebook)
        
        self.notebook.add(self.cargos_frame, text="Cargos")
        self.notebook.add(self.stock_frame, text="Stock")
        
        # Setup Cargos tab
        self.setup_cargos_tab()
        
        # Setup Stock tab (placeholder)
        self.setup_stock_tab()
    
    def setup_cargos_tab(self):
        """Setup the Cargos tab interface"""
        # Main frame
        main_frame = ttk.Frame(self.cargos_frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # File selection section
        file_section = ttk.LabelFrame(main_frame, text="File Selection", padding=10)
        file_section.pack(fill="x", pady=(0, 10))
        
        # Excel file selection
        ttk.Label(file_section, text="Excel File:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        ttk.Entry(file_section, textvariable=self.excel_file_path, width=50).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(file_section, text="Browse", command=self.browse_excel_file).grid(row=0, column=2)
        
        # Templates path
        ttk.Label(file_section, text="Templates Path:").grid(row=1, column=0, sticky="w", padx=(0, 5), pady=(5, 0))
        ttk.Entry(file_section, textvariable=self.templates_path, width=50).grid(row=1, column=1, padx=(0, 5), pady=(5, 0))
        ttk.Button(file_section, text="Browse", command=self.browse_templates_folder).grid(row=1, column=2, pady=(5, 0))
        
        # Destination path
        ttk.Label(file_section, text="Destination Path:").grid(row=2, column=0, sticky="w", padx=(0, 5), pady=(5, 0))
        ttk.Entry(file_section, textvariable=self.destination_path, width=50).grid(row=2, column=1, padx=(0, 5), pady=(5, 0))
        ttk.Button(file_section, text="Browse", command=self.browse_destination_folder).grid(row=2, column=2, pady=(5, 0))
        
        # Configure grid weights
        file_section.columnconfigure(1, weight=1)
        
        # Control section
        control_section = ttk.LabelFrame(main_frame, text="Controls", padding=10)
        control_section.pack(fill="x", pady=(0, 10))
        
        # Load and Process buttons
        ttk.Button(control_section, text="Load Excel", command=self.load_excel_file).pack(side="left", padx=(0, 10))
        ttk.Button(control_section, text="Generate Files", command=self.generate_files).pack(side="left", padx=(0, 10))
        ttk.Button(control_section, text="Clear Logs", command=self.clear_logs).pack(side="left")
        
        # Data preview section
        preview_section = ttk.LabelFrame(main_frame, text="Data Preview", padding=10)
        preview_section.pack(fill="both", expand=True, pady=(0, 10))
        
        # Treeview for data preview
        self.tree = ttk.Treeview(preview_section)
        self.tree.pack(fill="both", expand=True)
        
        # Scrollbars for treeview
        tree_scroll_y = ttk.Scrollbar(preview_section, orient="vertical", command=self.tree.yview)
        tree_scroll_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=tree_scroll_y.set)
        
        tree_scroll_x = ttk.Scrollbar(preview_section, orient="horizontal", command=self.tree.xview)
        tree_scroll_x.pack(side="bottom", fill="x")
        self.tree.configure(xscrollcommand=tree_scroll_x.set)
        
        # Logs section
        logs_section = ttk.LabelFrame(main_frame, text="Logs", padding=10)
        logs_section.pack(fill="both", expand=True)
        
        # Text widget for logs
        self.log_text = scrolledtext.ScrolledText(logs_section, height=10, wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True)
        
        # Data storage
        self.excel_data = None
    
    def setup_stock_tab(self):
        """Setup the Stock tab interface (placeholder)"""
        placeholder_label = ttk.Label(self.stock_frame, text="Stock functionality will be implemented later")
        placeholder_label.pack(expand=True)
    
    def create_default_directories(self):
        """Create default directories if they don't exist"""
        try:
            Path("templates").mkdir(exist_ok=True)
            Path("output").mkdir(exist_ok=True)
            self.log_message("Default directories created/verified")
        except Exception as e:
            self.log_message(f"Error creating directories: {str(e)}", "ERROR")
    
    def browse_excel_file(self):
        """Browse and select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_file_path.set(file_path)
            self.log_message(f"Excel file selected: {file_path}")
    
    def browse_templates_folder(self):
        """Browse and select templates folder"""
        folder_path = filedialog.askdirectory(title="Select Templates Folder")
        if folder_path:
            self.templates_path.set(folder_path + "/")
            self.log_message(f"Templates folder selected: {folder_path}")
    
    def browse_destination_folder(self):
        """Browse and select destination folder"""
        folder_path = filedialog.askdirectory(title="Select Destination Folder")
        if folder_path:
            self.destination_path.set(folder_path + "/")
            self.log_message(f"Destination folder selected: {folder_path}")
    
    def load_excel_file(self):
        """Load and parse Excel file"""
        try:
            if not self.excel_file_path.get():
                messagebox.showerror("Error", "Please select an Excel file first")
                return
            
            if not os.path.exists(self.excel_file_path.get()):
                messagebox.showerror("Error", "Selected Excel file does not exist")
                return
            
            self.log_message("Loading Excel file...")
            
            # Read Excel file
            self.excel_data = pd.read_excel(self.excel_file_path.get())
            
            # Clear existing treeview data
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Setup treeview columns
            if not self.excel_data.empty:
                self.tree["columns"] = list(self.excel_data.columns)
                self.tree["show"] = "headings"
                
                # Configure columns
                for col in self.excel_data.columns:
                    self.tree.heading(col, text=col)
                    self.tree.column(col, width=100, minwidth=50)
                
                # Insert data (limit to first 100 rows for preview)
                for index, row in self.excel_data.head(100).iterrows():
                    self.tree.insert("", "end", values=list(row))
                
                self.log_message(f"Excel file loaded successfully. {len(self.excel_data)} rows found.")
                
                if len(self.excel_data) > 100:
                    self.log_message("Note: Only first 100 rows are shown in preview")
            else:
                self.log_message("Excel file is empty", "WARNING")
                
        except Exception as e:
            error_msg = f"Error loading Excel file: {str(e)}"
            self.log_message(error_msg, "ERROR")
            self.logger.error(f"{error_msg}\n{traceback.format_exc()}")
            messagebox.showerror("Error", error_msg)
    
    def generate_files(self):
        """Generate PDF files based on Excel data and Word templates"""
        try:
            if self.excel_data is None or self.excel_data.empty:
                messagebox.showerror("Error", "Please load Excel data first")
                return
            
            if not os.path.exists(self.templates_path.get()):
                messagebox.showerror("Error", f"Templates folder does not exist: {self.templates_path.get()}")
                return
            
            # Create destination folder if it doesn't exist
            Path(self.destination_path.get()).mkdir(parents=True, exist_ok=True)
            
            self.log_message("Starting file generation...")
            
            # TODO: Implement actual file generation logic
            # This is where the parsing logic and PDF generation will be implemented
            self.log_message("File generation logic will be implemented based on your requirements")
            
            # Placeholder for now
            template_files = list(Path(self.templates_path.get()).glob("*.docx"))
            if not template_files:
                self.log_message("No Word template files found in templates folder", "WARNING")
                return
            
            self.log_message(f"Found {len(template_files)} template files:")
            for template in template_files:
                self.log_message(f"  - {template.name}")
            
            self.log_message(f"Ready to process {len(self.excel_data)} rows of data")
            
        except Exception as e:
            error_msg = f"Error generating files: {str(e)}"
            self.log_message(error_msg, "ERROR")
            self.logger.error(f"{error_msg}\n{traceback.format_exc()}")
            messagebox.showerror("Error", error_msg)
    
    def log_message(self, message, level="INFO"):
        """Add message to log display"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {level}: {message}\n"
        
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        
        # Also log to file
        if level == "ERROR":
            self.logger.error(message)
        elif level == "WARNING":
            self.logger.warning(message)
        else:
            self.logger.info(message)
    
    def clear_logs(self):
        """Clear the log display"""
        self.log_text.delete(1.0, tk.END)
        self.log_message("Logs cleared")

def main():
    root = tk.Tk()
    app = FileGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
