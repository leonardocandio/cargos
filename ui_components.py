"""
UI components for the File Generator application.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import Callable, Optional
from datetime import datetime
import pandas as pd

from models import ExcelData, AppConfig


class FileSelectionFrame:
    """Frame for file and path selection."""
    
    def __init__(self, parent, config: AppConfig):
        self.config = config
        self.frame = ttk.LabelFrame(parent, text="File Selection", padding=10)
        
        # Variables for file paths
        self.excel_file_var = tk.StringVar()
        self.templates_path_var = tk.StringVar(value=config.templates_path)
        self.destination_path_var = tk.StringVar(value=config.destination_path)
        
        # Callbacks
        self.on_excel_file_changed: Optional[Callable] = None
        self.on_templates_path_changed: Optional[Callable] = None
        self.on_destination_path_changed: Optional[Callable] = None
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create and layout widgets."""
        # Excel file selection
        ttk.Label(self.frame, text="Excel File:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        ttk.Entry(self.frame, textvariable=self.excel_file_var, width=50).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(self.frame, text="Browse", command=self._browse_excel_file).grid(row=0, column=2)
        
        # Templates path
        ttk.Label(self.frame, text="Templates Path:").grid(row=1, column=0, sticky="w", padx=(0, 5), pady=(5, 0))
        ttk.Entry(self.frame, textvariable=self.templates_path_var, width=50).grid(row=1, column=1, padx=(0, 5), pady=(5, 0))
        ttk.Button(self.frame, text="Browse", command=self._browse_templates_folder).grid(row=1, column=2, pady=(5, 0))
        
        # Destination path
        ttk.Label(self.frame, text="Destination Path:").grid(row=2, column=0, sticky="w", padx=(0, 5), pady=(5, 0))
        ttk.Entry(self.frame, textvariable=self.destination_path_var, width=50).grid(row=2, column=1, padx=(0, 5), pady=(5, 0))
        ttk.Button(self.frame, text="Browse", command=self._browse_destination_folder).grid(row=2, column=2, pady=(5, 0))
        
        # Configure grid weights
        self.frame.columnconfigure(1, weight=1)
    
    def _browse_excel_file(self):
        """Browse and select Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_file_var.set(file_path)
            self.config.excel_file_path = file_path
            if self.on_excel_file_changed:
                self.on_excel_file_changed(file_path)
    
    def _browse_templates_folder(self):
        """Browse and select templates folder."""
        folder_path = filedialog.askdirectory(title="Select Templates Folder")
        if folder_path:
            folder_path = folder_path.replace("/", "\\") + "\\"
            self.templates_path_var.set(folder_path)
            self.config.templates_path = folder_path
            if self.on_templates_path_changed:
                self.on_templates_path_changed(folder_path)
    
    def _browse_destination_folder(self):
        """Browse and select destination folder."""
        folder_path = filedialog.askdirectory(title="Select Destination Folder")
        if folder_path:
            folder_path = folder_path.replace("/", "\\") + "\\"
            self.destination_path_var.set(folder_path)
            self.config.destination_path = folder_path
            if self.on_destination_path_changed:
                self.on_destination_path_changed(folder_path)
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class ControlFrame:
    """Frame for control buttons."""
    
    def __init__(self, parent):
        self.frame = ttk.LabelFrame(parent, text="Controls", padding=10)
        
        # Callbacks
        self.on_load_excel: Optional[Callable] = None
        self.on_generate_files: Optional[Callable] = None
        self.on_clear_logs: Optional[Callable] = None
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create and layout widgets."""
        ttk.Button(self.frame, text="Load Excel", command=self._load_excel).pack(side="left", padx=(0, 10))
        ttk.Button(self.frame, text="Generate Files", command=self._generate_files).pack(side="left", padx=(0, 10))
        ttk.Button(self.frame, text="Clear Logs", command=self._clear_logs).pack(side="left")
    
    def _load_excel(self):
        """Handle load Excel button click."""
        if self.on_load_excel:
            self.on_load_excel()
    
    def _generate_files(self):
        """Handle generate files button click."""
        if self.on_generate_files:
            self.on_generate_files()
    
    def _clear_logs(self):
        """Handle clear logs button click."""
        if self.on_clear_logs:
            self.on_clear_logs()
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class DataPreviewFrame:
    """Frame for data preview with treeview."""
    
    def __init__(self, parent):
        self.frame = ttk.LabelFrame(parent, text="Data Preview", padding=10)
        self._create_widgets()
    
    def _create_widgets(self):
        """Create and layout widgets."""
        # Treeview for data preview
        self.tree = ttk.Treeview(self.frame)
        self.tree.pack(fill="both", expand=True)
        
        # Scrollbars for treeview
        tree_scroll_y = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
        tree_scroll_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=tree_scroll_y.set)
        
        tree_scroll_x = ttk.Scrollbar(self.frame, orient="horizontal", command=self.tree.xview)
        tree_scroll_x.pack(side="bottom", fill="x")
        self.tree.configure(xscrollcommand=tree_scroll_x.set)
    
    def update_data(self, excel_data: ExcelData):
        """Update treeview with Excel data."""
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if not excel_data.is_loaded:
            return
        
        # Setup columns
        self.tree["columns"] = excel_data.columns
        self.tree["show"] = "headings"
        
        # Configure columns
        for col in excel_data.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, minwidth=50)
        
        # Insert preview data
        preview_data = excel_data.preview_data
        for index, row in preview_data.iterrows():
            self.tree.insert("", "end", values=list(row))
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class LogFrame:
    """Frame for logging display."""
    
    def __init__(self, parent):
        self.frame = ttk.LabelFrame(parent, text="Logs", padding=10)
        self._create_widgets()
    
    def _create_widgets(self):
        """Create and layout widgets."""
        self.log_text = scrolledtext.ScrolledText(self.frame, height=10, wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True)
    
    def add_message(self, message: str, level: str = "INFO"):
        """Add message to log display."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {level}: {message}\n"
        
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
    
    def clear(self):
        """Clear log display."""
        self.log_text.delete(1.0, tk.END)
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class CargosTab:
    """Main tab for Cargos functionality."""
    
    def __init__(self, parent, config: AppConfig):
        self.frame = ttk.Frame(parent)
        self.config = config
        
        # Create main container
        self.main_frame = ttk.Frame(self.frame)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create UI components
        self.file_selection = FileSelectionFrame(self.main_frame, config)
        self.control_frame = ControlFrame(self.main_frame)
        self.data_preview = DataPreviewFrame(self.main_frame)
        self.log_frame = LogFrame(self.main_frame)
        
        # Layout components
        self._layout_components()
        
        # Callbacks
        self.on_load_excel: Optional[Callable] = None
        self.on_generate_files: Optional[Callable] = None
        
        # Setup callbacks
        self._setup_callbacks()
    
    def _layout_components(self):
        """Layout all components."""
        self.file_selection.pack(fill="x", pady=(0, 10))
        self.control_frame.pack(fill="x", pady=(0, 10))
        self.data_preview.pack(fill="both", expand=True, pady=(0, 10))
        self.log_frame.pack(fill="both", expand=True)
    
    def _setup_callbacks(self):
        """Setup component callbacks."""
        self.file_selection.on_excel_file_changed = self._on_excel_file_changed
        self.file_selection.on_templates_path_changed = self._on_templates_path_changed
        self.file_selection.on_destination_path_changed = self._on_destination_path_changed
        
        self.control_frame.on_load_excel = self._on_load_excel
        self.control_frame.on_generate_files = self._on_generate_files
        self.control_frame.on_clear_logs = self._on_clear_logs
    
    def _on_excel_file_changed(self, file_path: str):
        """Handle Excel file path change."""
        self.log_message(f"Excel file selected: {file_path}")
    
    def _on_templates_path_changed(self, path: str):
        """Handle templates path change."""
        self.log_message(f"Templates folder selected: {path}")
    
    def _on_destination_path_changed(self, path: str):
        """Handle destination path change."""
        self.log_message(f"Destination folder selected: {path}")
    
    def _on_load_excel(self):
        """Handle load Excel button click."""
        if self.on_load_excel:
            self.on_load_excel()
    
    def _on_generate_files(self):
        """Handle generate files button click."""
        if self.on_generate_files:
            self.on_generate_files()
    
    def _on_clear_logs(self):
        """Handle clear logs button click."""
        self.log_frame.clear()
        self.log_message("Logs cleared")
    
    def update_data_preview(self, excel_data: ExcelData):
        """Update data preview."""
        self.data_preview.update_data(excel_data)
    
    def log_message(self, message: str, level: str = "INFO"):
        """Add message to log."""
        self.log_frame.add_message(message, level)
    
    def show_error(self, title: str, message: str):
        """Show error message box."""
        messagebox.showerror(title, message)
    
    def show_info(self, title: str, message: str):
        """Show info message box."""
        messagebox.showinfo(title, message)
