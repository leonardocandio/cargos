"""
UI components for the File Generator application.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import Callable, Optional
from datetime import datetime
import pandas as pd

from models import ExcelData, AppConfig


class CompactFileSelectionFrame:
    """Compact frame for file and path selection."""
    
    def __init__(self, parent, config: AppConfig):
        self.config = config
        self.frame = ttk.LabelFrame(parent, text="Files", padding=8)
        
        # Variables for file paths
        self.excel_file_var = tk.StringVar()
        self.cargo_template_var = tk.StringVar(value=config.cargo_template_path)
        self.autorizacion_template_var = tk.StringVar(value=config.autorizacion_template_path)
        self.destination_path_var = tk.StringVar(value=config.destination_path)
        
        # Callbacks
        self.on_excel_file_changed: Optional[Callable] = None
        self.on_cargo_template_changed: Optional[Callable] = None
        self.on_autorizacion_template_changed: Optional[Callable] = None
        self.on_destination_path_changed: Optional[Callable] = None
        self.on_load_excel: Optional[Callable] = None
        self.on_reload_excel: Optional[Callable] = None
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create and layout widgets."""
        from pathlib import Path
        
        # Formato Uniforme file selection (compact)
        ttk.Label(self.frame, text="Formato Uniforme:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.excel_path_label = ttk.Label(self.frame, text="(no file)", width=30, anchor="w", foreground="gray")
        self.excel_path_label.grid(row=0, column=1, sticky="w", padx=(0, 5))
        ttk.Button(self.frame, text="Browse", command=self._browse_excel_file).grid(row=0, column=2)
        self.load_button = ttk.Button(self.frame, text="Load", command=self._load_excel, state="disabled")
        self.load_button.grid(row=0, column=3, padx=(5, 0))
        self.reload_button = ttk.Button(self.frame, text="Reload", command=self._reload_excel, state="disabled")
        self.reload_button.grid(row=0, column=4, padx=(2, 0))
        
        # CARGO template file (compact)
        ttk.Label(self.frame, text="CARGO Template:", font=("TkDefaultFont", 8)).grid(row=1, column=0, sticky="w", padx=(0, 5), pady=(3, 0))
        self.cargo_path_label = ttk.Label(self.frame, text=Path(self.cargo_template_var.get()).name if self.cargo_template_var.get() else "(no file)", 
                                         width=30, anchor="w", foreground="gray", font=("TkDefaultFont", 8))
        self.cargo_path_label.grid(row=1, column=1, sticky="w", padx=(0, 5), pady=(3, 0))
        ttk.Button(self.frame, text="Browse", command=self._browse_cargo_template).grid(row=1, column=2, pady=(3, 0))
        
        # AUTORIZACION template file (compact)
        ttk.Label(self.frame, text="AUTORIZACION Template:", font=("TkDefaultFont", 8)).grid(row=2, column=0, sticky="w", padx=(0, 5), pady=(3, 0))
        self.autorizacion_path_label = ttk.Label(self.frame, text=Path(self.autorizacion_template_var.get()).name if self.autorizacion_template_var.get() else "(no file)", 
                                               width=30, anchor="w", foreground="gray", font=("TkDefaultFont", 8))
        self.autorizacion_path_label.grid(row=2, column=1, sticky="w", padx=(0, 5), pady=(3, 0))
        ttk.Button(self.frame, text="Browse", command=self._browse_autorizacion_template).grid(row=2, column=2, pady=(3, 0))
        
        # Destination folder (compact)
        ttk.Label(self.frame, text="Destination:", font=("TkDefaultFont", 8)).grid(row=3, column=0, sticky="w", padx=(0, 5), pady=(3, 0))
        self.destination_path_label = ttk.Label(self.frame, text=Path(self.destination_path_var.get()).name if self.destination_path_var.get() else "(no folder)", 
                                              width=30, anchor="w", foreground="gray", font=("TkDefaultFont", 8))
        self.destination_path_label.grid(row=3, column=1, sticky="w", padx=(0, 5), pady=(3, 0))
        ttk.Button(self.frame, text="Browse", command=self._browse_destination_folder).grid(row=3, column=2, pady=(3, 0))
        
        # Configure grid weights
        self.frame.columnconfigure(1, weight=1)
    
    def _browse_excel_file(self):
        """Browse and select Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Formato Uniforme File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            from pathlib import Path
            self.excel_file_var.set(file_path)
            self.excel_path_label.configure(text=Path(file_path).name, foreground="black")
            self.config.excel_file_path = file_path
            
            # Enable load and reload buttons
            self.load_button.configure(state="normal")
            self.reload_button.configure(state="normal")
            
            if self.on_excel_file_changed:
                self.on_excel_file_changed(file_path)
    
    def _load_excel(self):
        """Handle load Excel button click."""
        if self.on_load_excel:
            self.on_load_excel()
    
    def _reload_excel(self):
        """Handle reload Excel button click."""
        if self.on_reload_excel:
            self.on_reload_excel()
    
    def _browse_cargo_template(self):
        """Browse and select CARGO template file."""
        file_path = filedialog.askopenfilename(
            title="Select CARGO Template",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        if file_path:
            from pathlib import Path
            self.cargo_template_var.set(file_path)
            self.cargo_path_label.configure(text=Path(file_path).name, foreground="black")
            self.config.cargo_template_path = file_path
            if self.on_cargo_template_changed:
                self.on_cargo_template_changed(file_path)
    
    def _browse_autorizacion_template(self):
        """Browse and select AUTORIZACION template file."""
        file_path = filedialog.askopenfilename(
            title="Select AUTORIZACION Template",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        if file_path:
            from pathlib import Path
            self.autorizacion_template_var.set(file_path)
            self.autorizacion_path_label.configure(text=Path(file_path).name, foreground="black")
            self.config.autorizacion_template_path = file_path
            if self.on_autorizacion_template_changed:
                self.on_autorizacion_template_changed(file_path)
    
    def _browse_destination_folder(self):
        """Browse and select destination folder."""
        folder_path = filedialog.askdirectory(title="Select Destination Folder")
        if folder_path:
            from pathlib import Path
            folder_path = folder_path.replace("/", "\\") + "\\"
            self.destination_path_var.set(folder_path)
            self.destination_path_label.configure(text=Path(folder_path).name, foreground="black")
            self.config.destination_path = folder_path
            if self.on_destination_path_changed:
                self.on_destination_path_changed(folder_path)
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class GenerateButtonFrame:
    """Frame for generate files button at bottom."""
    
    def __init__(self, parent):
        self.frame = ttk.Frame(parent)
        
        # Callbacks
        self.on_generate_files: Optional[Callable] = None
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create and layout widgets."""
        # Create a prominent generate button
        self.generate_btn = ttk.Button(
            self.frame, 
            text="🚀 GENERATE FILES", 
            command=self._generate_files,
            state="disabled"
        )
        # Make button bigger and more prominent
        self.generate_btn.configure(width=25)
        self.generate_btn.pack(side="right", padx=10, pady=8, ipady=5)
    
    def _generate_files(self):
        """Handle generate files button click."""
        if self.on_generate_files:
            self.on_generate_files()
    
    def set_enabled(self, enabled: bool):
        """Enable or disable the generate button."""
        self.generate_btn.configure(state="normal" if enabled else "disabled")
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class WorksheetSummaryFrame:
    """Frame for showing worksheet parsing summary."""
    
    def __init__(self, parent):
        self.frame = ttk.LabelFrame(parent, text="Worksheet Summary", padding=10)
        self._create_widgets()
    
    def _create_widgets(self):
        """Create and layout widgets."""
        from constants import DEFAULT_TREE_HEIGHT
        
        # Create treeview for worksheet summary
        columns = ("Sheet", "People", "Errors", "Status")
        self.tree = ttk.Treeview(self.frame, columns=columns, show="headings", height=DEFAULT_TREE_HEIGHT)
        
        # Configure columns
        self.tree.heading("Sheet", text="Sheet Name")
        self.tree.heading("People", text="People Parsed")
        self.tree.heading("Errors", text="Errors")
        self.tree.heading("Status", text="Status")
        
        self.tree.column("Sheet", width=180, minwidth=120)
        self.tree.column("People", width=120, minwidth=100)
        self.tree.column("Errors", width=80, minwidth=60)
        self.tree.column("Status", width=100, minwidth=80)
        
        self.tree.pack(fill="both", expand=True)
        
        # Add double-click binding to view worksheet details
        self.tree.bind("<Double-1>", self._on_worksheet_double_click)
        
        # Callback for worksheet selection
        self.on_worksheet_selected: Optional[Callable] = None
    
    def _on_worksheet_double_click(self, event):
        """Handle double-click on worksheet."""
        selection = self.tree.selection()
        if selection and self.on_worksheet_selected:
            item = self.tree.item(selection[0])
            sheet_name = item['values'][0]
            self.on_worksheet_selected(sheet_name)
    
    def update_data(self, excel_data: ExcelData):
        """Update treeview with worksheet summary."""
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if not excel_data.is_loaded:
            return
        
        # Add worksheet summary rows
        for worksheet in excel_data.worksheets:
            status = "✓ Success" if worksheet.data is not None else "✗ Failed"
            error_count = len(worksheet.errors)
            
            self.tree.insert("", "end", values=(
                worksheet.metadata.sheet_name,
                worksheet.people_parsed,
                error_count,
                status
            ))
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class DataPreviewFrame:
    """Frame for data preview with treeview."""
    
    def __init__(self, parent):
        self.frame = ttk.LabelFrame(parent, text="Data Preview", padding=10)
        self.current_excel_data: Optional[ExcelData] = None
        self._create_widgets()
    
    def _create_widgets(self):
        """Create and layout widgets."""
        # Create notebook for different views
        self.notebook = ttk.Notebook(self.frame)
        self.notebook.pack(fill="both", expand=True)
        
        # Summary tab
        self.summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="Summary")
        
        # Worksheet summary
        self.worksheet_summary = WorksheetSummaryFrame(self.summary_frame)
        self.worksheet_summary.pack(fill="both", expand=True)
        self.worksheet_summary.on_worksheet_selected = self._show_worksheet_details
        
        # Details tab (will be populated dynamically)
        self.details_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.details_frame, text="Details")
        
        # Uniform Data tab
        self.uniform_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.uniform_frame, text="Uniforms")
        
        # Logs tab
        self.logs_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.logs_frame, text="Logs")
        
        # Create details content
        self._create_details_widgets()
        
        # Create uniform data content
        self._create_uniform_widgets()
        
        # Create logs content
        self._create_logs_widgets()
    
    def _create_details_widgets(self):
        """Create widgets for details tab."""
        # Worksheet selector
        selector_frame = ttk.Frame(self.details_frame)
        selector_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(selector_frame, text="Select Worksheet:").pack(side="left", padx=(0, 5))
        self.worksheet_var = tk.StringVar()
        self.worksheet_combo = ttk.Combobox(selector_frame, textvariable=self.worksheet_var, state="readonly")
        self.worksheet_combo.pack(side="left", padx=(0, 10))
        self.worksheet_combo.bind("<<ComboboxSelected>>", self._on_worksheet_selected)
        
        ttk.Button(selector_frame, text="Refresh", command=self._refresh_worksheet_details).pack(side="left")
        
        # Data treeview (expanded to take full space)
        data_frame = ttk.LabelFrame(self.details_frame, text="Data", padding=5)
        data_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Metadata footnote (one line at bottom of data frame)
        self.metadata_label = ttk.Label(data_frame, text="", font=("TkDefaultFont", 8), foreground="gray")
        self.metadata_label.pack(side="bottom", fill="x", pady=(2, 0))
        
        self.data_tree = ttk.Treeview(data_frame)
        self.data_tree.pack(fill="both", expand=True, pady=(0, 2))
        
        # Scrollbars for data treeview
        data_scroll_y = ttk.Scrollbar(data_frame, orient="vertical", command=self.data_tree.yview)
        data_scroll_y.pack(side="right", fill="y")
        self.data_tree.configure(yscrollcommand=data_scroll_y.set)
        
        data_scroll_x = ttk.Scrollbar(data_frame, orient="horizontal", command=self.data_tree.xview)
        data_scroll_x.pack(side="bottom", fill="x")
        self.data_tree.configure(xscrollcommand=data_scroll_x.set)
    
    def _create_logs_widgets(self):
        """Create widgets for logs tab."""
        # Control buttons
        control_frame = ttk.Frame(self.logs_frame)
        control_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(control_frame, text="Clear Logs", command=self.clear_logs).pack(side="right", padx=2)
        ttk.Label(control_frame, text="Application Logs:", font=("TkDefaultFont", 9)).pack(side="left")
        
        # Log text area
        self.log_text = scrolledtext.ScrolledText(
            self.logs_frame, 
            wrap=tk.WORD, 
            font=("Consolas", 9)
        )
        self.log_text.pack(fill="both", expand=True, padx=5, pady=(0, 5))
    
    def _create_uniform_widgets(self):
        """Create widgets for uniform data tab."""
        # Worksheet selector for uniform data
        selector_frame = ttk.Frame(self.uniform_frame)
        selector_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(selector_frame, text="Select Worksheet:").pack(side="left", padx=(0, 5))
        self.uniform_worksheet_var = tk.StringVar()
        self.uniform_worksheet_combo = ttk.Combobox(selector_frame, textvariable=self.uniform_worksheet_var, state="readonly")
        self.uniform_worksheet_combo.pack(side="left", padx=(0, 10))
        self.uniform_worksheet_combo.bind("<<ComboboxSelected>>", self._on_uniform_worksheet_selected)
        
        ttk.Button(selector_frame, text="Refresh", command=self._refresh_uniform_details).pack(side="left")
        
        # Uniform data treeview
        uniform_data_frame = ttk.LabelFrame(self.uniform_frame, text="Uniform Data (Columns J-R)", padding=5)
        uniform_data_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Metadata footnote for uniform data
        self.uniform_metadata_label = ttk.Label(uniform_data_frame, text="", font=("TkDefaultFont", 8), foreground="gray")
        self.uniform_metadata_label.pack(side="bottom", fill="x", pady=(2, 0))
        
        self.uniform_data_tree = ttk.Treeview(uniform_data_frame)
        self.uniform_data_tree.pack(fill="both", expand=True, pady=(0, 2))
        
        # Scrollbars for uniform data treeview
        uniform_scroll_y = ttk.Scrollbar(uniform_data_frame, orient="vertical", command=self.uniform_data_tree.yview)
        uniform_scroll_y.pack(side="right", fill="y")
        self.uniform_data_tree.configure(yscrollcommand=uniform_scroll_y.set)
        
        uniform_scroll_x = ttk.Scrollbar(uniform_data_frame, orient="horizontal", command=self.uniform_data_tree.xview)
        uniform_scroll_x.pack(side="bottom", fill="x")
        self.uniform_data_tree.configure(xscrollcommand=uniform_scroll_x.set)
    
    def _show_worksheet_details(self, sheet_name: str):
        """Show details for a specific worksheet."""
        self.worksheet_var.set(sheet_name)
        self._refresh_worksheet_details()
        self.notebook.select(1)  # Switch to details tab
    
    def _on_worksheet_selected(self, event):
        """Handle worksheet selection from combobox."""
        self._refresh_worksheet_details()
    
    def _refresh_worksheet_details(self):
        """Refresh the worksheet details view."""
        if not self.current_excel_data or not self.worksheet_var.get():
            return
        
        sheet_name = self.worksheet_var.get()
        worksheet = self.current_excel_data.get_worksheet_by_name(sheet_name)
        
        if not worksheet:
            return
        
        # Update metadata footnote (one line)
        metadata_info = f"Sheet: {worksheet.metadata.sheet_name} | "
        metadata_info += f"Fecha: {worksheet.metadata.fecha_solicitud} | "
        metadata_info += f"Tienda: {worksheet.metadata.tienda} | "
        metadata_info += f"Admin: {worksheet.metadata.administrador}"
        
        if worksheet.errors:
            metadata_info += f" | Errors: {len(worksheet.errors)}"
        if worksheet.warnings:
            metadata_info += f" | Warnings: {len(worksheet.warnings)}"
        
        self.metadata_label.configure(text=metadata_info)
        
        # Update data treeview
        self._update_data_tree(worksheet)
    
    def _on_uniform_worksheet_selected(self, event):
        """Handle uniform worksheet selection from combobox."""
        self._refresh_uniform_details()
    
    def _refresh_uniform_details(self):
        """Refresh the uniform worksheet details view."""
        if not self.current_excel_data or not self.uniform_worksheet_var.get():
            return
        
        sheet_name = self.uniform_worksheet_var.get()
        worksheet = self.current_excel_data.get_worksheet_by_name(sheet_name)
        
        if not worksheet:
            return
        
        # Update uniform metadata footnote (one line)
        metadata_info = f"Sheet: {worksheet.metadata.sheet_name} | "
        metadata_info += f"Fecha: {worksheet.metadata.fecha_solicitud} | "
        metadata_info += f"Tienda: {worksheet.metadata.tienda} | "
        metadata_info += f"Admin: {worksheet.metadata.administrador}"
        
        if worksheet.uniform_data is not None:
            metadata_info += f" | Uniform Rows: {len(worksheet.uniform_data)}"
        else:
            metadata_info += " | No uniform data"
        
        self.uniform_metadata_label.configure(text=metadata_info)
        
        # Update uniform data treeview
        self._update_uniform_data_tree(worksheet)
    
    def _update_uniform_data_tree(self, worksheet):
        """Update uniform data treeview with worksheet uniform data."""
        # Clear existing data
        for item in self.uniform_data_tree.get_children():
            self.uniform_data_tree.delete(item)
        
        if worksheet.uniform_data is None or worksheet.uniform_data.empty:
            return
        
        # Setup columns
        columns = list(worksheet.uniform_data.columns)
        self.uniform_data_tree["columns"] = columns
        self.uniform_data_tree["show"] = "headings"
        
        # Configure columns
        for col in columns:
            self.uniform_data_tree.heading(col, text=str(col))
            self.uniform_data_tree.column(col, width=100, minwidth=50)
        
        # Insert data (limit to first 100 rows for performance)
        for index, row in worksheet.uniform_data.head(100).iterrows():
            values = [str(val) if pd.notna(val) else "" for val in row]
            self.uniform_data_tree.insert("", "end", values=values)
    
    def _update_data_tree(self, worksheet):
        """Update data treeview with worksheet data."""
        # Clear existing data
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)
        
        if worksheet.data is None or worksheet.data.empty:
            return
        
        # Setup columns
        columns = list(worksheet.data.columns)
        self.data_tree["columns"] = columns
        self.data_tree["show"] = "headings"
        
        # Configure columns
        for col in columns:
            self.data_tree.heading(col, text=str(col))
            self.data_tree.column(col, width=100, minwidth=50)
        
        # Insert data (limit to first 100 rows for performance)
        for index, row in worksheet.data.head(100).iterrows():
            values = [str(val) if pd.notna(val) else "" for val in row]
            self.data_tree.insert("", "end", values=values)
    
    def update_data(self, excel_data: ExcelData):
        """Update preview with Excel data."""
        self.current_excel_data = excel_data
        
        # Update summary
        self.worksheet_summary.update_data(excel_data)
        
        # Update worksheet selector
        if excel_data.worksheets:
            sheet_names = [w.metadata.sheet_name for w in excel_data.worksheets]
            self.worksheet_combo['values'] = sheet_names
            self.uniform_worksheet_combo['values'] = sheet_names
            if sheet_names:
                self.worksheet_var.set(sheet_names[0])
                self.uniform_worksheet_var.set(sheet_names[0])
                self._refresh_worksheet_details()
                self._refresh_uniform_details()
        else:
            self.worksheet_combo['values'] = []
            self.uniform_worksheet_combo['values'] = []
            self.worksheet_var.set("")
            self.uniform_worksheet_var.set("")
    
    def add_log_message(self, message: str, level: str = "INFO"):
        """Add message to log display."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {level}: {message}\n"
        
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        
        # Auto-switch to logs tab if there's an error
        if level == "ERROR":
            self.notebook.select(3)  # Logs tab (now 4th tab: Summary, Details, Uniforms, Logs)
    
    def clear_logs(self):
        """Clear log display."""
        self.log_text.delete(1.0, tk.END)
        self.add_log_message("Logs cleared")
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)




class CargosTab:
    """Main tab for Cargos functionality with optimized layout."""
    
    def __init__(self, parent, config: AppConfig):
        self.frame = ttk.Frame(parent)
        self.config = config
        
        # Create main container
        self.main_frame = ttk.Frame(self.frame)
        self.main_frame.pack(fill="both", expand=True, padx=8, pady=8)
        
        # Create UI components
        self.file_selection = CompactFileSelectionFrame(self.main_frame, config)
        self.data_preview = DataPreviewFrame(self.main_frame)
        self.generate_button = GenerateButtonFrame(self.main_frame)
        
        # Layout components
        self._layout_components()
        
        # Callbacks
        self.on_load_excel: Optional[Callable] = None
        self.on_generate_files: Optional[Callable] = None
        self.on_config_changed: Optional[Callable] = None
        
        # Setup callbacks
        self._setup_callbacks()
    
    def _layout_components(self):
        """Layout all components with optimized space usage."""
        # File selection at top (compact)
        self.file_selection.pack(fill="x", pady=(0, 8))
        
        # Data preview takes most space
        self.data_preview.pack(fill="both", expand=True, pady=(0, 8))
        
        # Generate button at bottom
        self.generate_button.pack(fill="x", pady=(0, 8))
    
    def _setup_callbacks(self):
        """Setup component callbacks."""
        self.file_selection.on_excel_file_changed = self._on_excel_file_changed
        self.file_selection.on_cargo_template_changed = self._on_cargo_template_changed
        self.file_selection.on_autorizacion_template_changed = self._on_autorizacion_template_changed
        self.file_selection.on_destination_path_changed = self._on_destination_path_changed
        self.file_selection.on_load_excel = self._on_load_excel
        self.file_selection.on_reload_excel = self._on_reload_excel
        
        self.generate_button.on_generate_files = self._on_generate_files
    
    def _on_excel_file_changed(self, file_path: str):
        """Handle Excel file path change."""
        self.log_message(f"Excel file selected: {file_path}")
    
    def _on_cargo_template_changed(self, path: str):
        """Handle CARGO template path change."""
        self.log_message(f"CARGO template selected: {path}")
        # Trigger configuration save
        if hasattr(self, 'on_config_changed') and self.on_config_changed:
            self.on_config_changed()
    
    def _on_autorizacion_template_changed(self, path: str):
        """Handle AUTORIZACION template path change."""
        self.log_message(f"AUTORIZACION template selected: {path}")
        # Trigger configuration save
        if hasattr(self, 'on_config_changed') and self.on_config_changed:
            self.on_config_changed()
    
    def _on_destination_path_changed(self, path: str):
        """Handle destination path change."""
        self.log_message(f"Destination folder selected: {path}")
    
    def _on_load_excel(self):
        """Handle load Excel button click."""
        if self.on_load_excel:
            self.on_load_excel()
    
    def _on_reload_excel(self):
        """Handle reload Excel button click."""
        if self.on_load_excel:  # Use same callback as load
            self.on_load_excel()
    
    def _on_generate_files(self):
        """Handle generate files button click."""
        if self.on_generate_files:
            self.on_generate_files()
    
    def update_data_preview(self, excel_data: ExcelData):
        """Update data preview."""
        self.data_preview.update_data(excel_data)
        
        # Enable generate button if at least one sheet was successfully parsed
        self.generate_button.set_enabled(excel_data.successful_worksheets > 0)
    
    def log_message(self, message: str, level: str = "INFO"):
        """Add message to log."""
        self.data_preview.add_log_message(message, level)
    
    def show_error(self, title: str, message: str):
        """Show error message box."""
        messagebox.showerror(title, message)
    
    def show_info(self, title: str, message: str):
        """Show info message box."""
        messagebox.showinfo(title, message)
