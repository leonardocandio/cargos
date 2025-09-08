"""
File Generator Application - Main Entry Point

A GUI application for generating PDF files from Excel data using Word templates.
Follows clean architecture principles with separation of concerns.
"""
import tkinter as tk
from tkinter import ttk
import logging
from pathlib import Path
from typing import Optional

from models import ExcelData, AppConfig
from services import ExcelService, FileGenerationService, ConfigService
from ui_components import CargosTab


class FileGeneratorApp:
    """
    Main application class that orchestrates the UI and services.
    
    This class acts as the controller, coordinating between the UI components
    and the business logic services.
    """
    
    def __init__(self, root: tk.Tk):
        """
        Initialize the application.
        
        Args:
            root: The main tkinter window
        """
        self.root = root
        self.root.title("File Generator - Excel to PDF")
        self.root.geometry("800x600")
        
        # Initialize services
        self.config_service = ConfigService()
        self.config = self.config_service.get_config()
        
        # Setup logging
        self.logger = self._setup_logging()
        
        # Initialize services with logger
        self.excel_service = ExcelService(self.logger)
        self.file_generation_service = FileGenerationService(self.logger)
        
        # Data storage
        self.excel_data: Optional[ExcelData] = None
        
        # Create UI
        self._create_ui()
        
        # Create default directories
        self._create_default_directories()
        
        self.logger.info("Application initialized successfully")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging configuration."""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.config.log_file),
                logging.StreamHandler()
            ]
        )
        return logging.getLogger(__name__)
    
    def _create_ui(self):
        """Create the main UI with tabs."""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create Cargos tab
        self.cargos_tab = CargosTab(self.notebook, self.config)
        self.notebook.add(self.cargos_tab.frame, text="Cargos")
        
        # Create Stock tab (placeholder)
        self.stock_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.stock_frame, text="Stock")
        self._setup_stock_tab()
        
        # Setup callbacks
        self._setup_callbacks()
    
    def _setup_stock_tab(self):
        """Setup the Stock tab interface (placeholder)."""
        placeholder_label = ttk.Label(
            self.stock_frame, 
            text="Stock functionality will be implemented later"
        )
        placeholder_label.pack(expand=True)
    
    def _setup_callbacks(self):
        """Setup callbacks between UI and business logic."""
        self.cargos_tab.on_load_excel = self._handle_load_excel
        self.cargos_tab.on_generate_files = self._handle_generate_files
    
    def _create_default_directories(self):
        """Create default directories if they don't exist."""
        try:
            Path(self.config.templates_path).mkdir(exist_ok=True)
            Path(self.config.destination_path).mkdir(exist_ok=True)
            self.cargos_tab.log_message("Default directories created/verified")
        except Exception as e:
            error_msg = f"Error creating directories: {str(e)}"
            self.cargos_tab.log_message(error_msg, "ERROR")
            self.logger.error(error_msg)
    
    def _handle_load_excel(self):
        """Handle Excel file loading."""
        try:
            if not self.config.excel_file_path:
                self.cargos_tab.show_error("Error", "Please select an Excel file first")
                return
            
            # Load Excel data using service
            self.excel_data = self.excel_service.load_excel_file(self.config.excel_file_path)
            
            # Update UI with loaded data
            self.cargos_tab.update_data_preview(self.excel_data)
            
            success_msg = f"Excel file loaded successfully. {self.excel_data.total_rows} rows found."
            self.cargos_tab.log_message(success_msg)
            
            if self.excel_data.total_rows > self.config.preview_rows_limit:
                self.cargos_tab.log_message(
                    f"Note: Only first {self.config.preview_rows_limit} rows are shown in preview"
                )
                
        except Exception as e:
            error_msg = str(e)
            self.cargos_tab.log_message(error_msg, "ERROR")
            self.cargos_tab.show_error("Error", error_msg)
    
    def _handle_generate_files(self):
        """Handle file generation process."""
        try:
            if not self.excel_data or not self.excel_data.is_loaded:
                self.cargos_tab.show_error("Error", "Please load Excel data first")
                return
            
            # Validate configuration
            config_errors = self.config_service.validate_paths()
            if config_errors:
                error_msg = "\n".join(config_errors)
                self.cargos_tab.show_error("Configuration Error", error_msg)
                return
            
            # Generate files using service
            result = self.file_generation_service.generate_files(self.excel_data, self.config)
            
            if result.success:
                self.cargos_tab.log_message(result.message)
                if result.files_generated > 0:
                    self.cargos_tab.show_info(
                        "Success", 
                        f"Generated {result.files_generated} files successfully!"
                    )
            else:
                self.cargos_tab.log_message(result.message, "ERROR")
                self.cargos_tab.show_error("Generation Error", result.message)
                
        except Exception as e:
            error_msg = f"Unexpected error during file generation: {str(e)}"
            self.cargos_tab.log_message(error_msg, "ERROR")
            self.cargos_tab.show_error("Error", error_msg)
            self.logger.exception("Unexpected error in file generation")

def main():
    root = tk.Tk()
    app = FileGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
