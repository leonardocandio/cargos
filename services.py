"""
Service classes for Excel processing and file generation.
"""
import pandas as pd
import logging
from pathlib import Path
from typing import List, Optional
import traceback

from models import ExcelData, GenerationResult, AppConfig


class ExcelService:
    """Service for handling Excel file operations."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
    
    def load_excel_file(self, file_path: str) -> ExcelData:
        """
        Load Excel file and return ExcelData object.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            ExcelData object with loaded data or empty if failed
        """
        excel_data = ExcelData()
        
        try:
            if not file_path:
                raise ValueError("File path is empty")
            
            if not Path(file_path).exists():
                raise FileNotFoundError(f"File does not exist: {file_path}")
            
            self.logger.info(f"Loading Excel file: {file_path}")
            
            # Read Excel file
            data = pd.read_excel(file_path)
            
            excel_data.data = data
            excel_data.file_path = file_path
            excel_data.total_rows = len(data)
            excel_data.columns = list(data.columns)
            
            self.logger.info(f"Excel file loaded successfully. {excel_data.total_rows} rows found.")
            
            return excel_data
            
        except Exception as e:
            error_msg = f"Error loading Excel file: {str(e)}"
            self.logger.error(f"{error_msg}\n{traceback.format_exc()}")
            raise Exception(error_msg)
    
    def validate_excel_data(self, excel_data: ExcelData) -> bool:
        """
        Validate Excel data for processing.
        
        Args:
            excel_data: ExcelData object to validate
            
        Returns:
            True if data is valid for processing
        """
        if not excel_data.is_loaded:
            self.logger.warning("No Excel data loaded")
            return False
        
        if excel_data.total_rows == 0:
            self.logger.warning("Excel file is empty")
            return False
        
        # Add more validation rules here as needed
        return True


class FileGenerationService:
    """Service for generating files from Excel data using templates."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
    
    def get_template_files(self, templates_path: str) -> List[Path]:
        """
        Get list of Word template files.
        
        Args:
            templates_path: Path to templates directory
            
        Returns:
            List of template file paths
        """
        try:
            templates_dir = Path(templates_path)
            if not templates_dir.exists():
                self.logger.warning(f"Templates directory does not exist: {templates_path}")
                return []
            
            template_files = list(templates_dir.glob("*.docx"))
            self.logger.info(f"Found {len(template_files)} template files")
            
            return template_files
            
        except Exception as e:
            self.logger.error(f"Error getting template files: {str(e)}")
            return []
    
    def generate_files(self, excel_data: ExcelData, config: AppConfig) -> GenerationResult:
        """
        Generate files from Excel data using templates.
        
        Args:
            excel_data: Loaded Excel data
            config: Application configuration
            
        Returns:
            GenerationResult with success status and details
        """
        result = GenerationResult(success=False)
        
        try:
            if not excel_data.is_loaded:
                result.message = "No Excel data loaded"
                return result
            
            # Get template files
            template_files = self.get_template_files(config.templates_path)
            if not template_files:
                result.message = "No template files found"
                return result
            
            # Create destination directory
            dest_path = Path(config.destination_path)
            dest_path.mkdir(parents=True, exist_ok=True)
            
            self.logger.info("Starting file generation process...")
            
            # TODO: Implement actual file generation logic
            # This is where you'll add the specific logic for:
            # 1. Processing Excel data row by row
            # 2. Filling Word templates with data
            # 3. Converting to PDF
            # 4. Saving files with appropriate names
            
            # Placeholder implementation
            result.success = True
            result.files_generated = 0
            result.message = f"Ready to process {excel_data.total_rows} rows with {len(template_files)} templates"
            
            self.logger.info(result.message)
            
            return result
            
        except Exception as e:
            error_msg = f"Error during file generation: {str(e)}"
            self.logger.error(f"{error_msg}\n{traceback.format_exc()}")
            result.message = error_msg
            result.errors.append(error_msg)
            return result


class ConfigService:
    """Service for handling application configuration."""
    
    def __init__(self):
        self.config = AppConfig()
    
    def get_config(self) -> AppConfig:
        """Get current configuration."""
        return self.config
    
    def update_config(self, **kwargs) -> None:
        """Update configuration values."""
        for key, value in kwargs.items():
            if hasattr(self.config, key):
                setattr(self.config, key, value)
    
    def validate_paths(self) -> List[str]:
        """
        Validate configuration paths.
        
        Returns:
            List of validation errors
        """
        errors = []
        
        # Check templates path
        if not Path(self.config.templates_path).exists():
            errors.append(f"Templates path does not exist: {self.config.templates_path}")
        
        # Destination path will be created if it doesn't exist, so no validation needed
        
        return errors
