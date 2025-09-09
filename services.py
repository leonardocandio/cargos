"""
Service classes for Excel processing and file generation.
"""
import pandas as pd
import logging
from pathlib import Path
from typing import List, Optional
import traceback

from models import ExcelData, GenerationResult, AppConfig, ExcelValidationResult, WorksheetMetadata, WorksheetParsingResult
from config_manager import ConfigManager
from validators import TemplateValidator
from constants import (
    METADATA_ROW_FECHA_SOLICITUD, METADATA_COL_FECHA_SOLICITUD,
    METADATA_ROW_TIENDA, METADATA_COL_TIENDA,
    METADATA_ROW_ADMINISTRADOR, METADATA_COL_ADMINISTRADOR,
    DATA_START_ROW, IGNORE_COLUMN_INDEX, MAIN_DATA_END_COLUMN,
    UNIFORM_DATA_START_ROW, UNIFORM_DATA_START_COLUMN, UNIFORM_DATA_END_COLUMN,
    UNIFORM_COLUMN_NAMES
)


class ExcelService:
    """Service for handling Excel file operations."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
    
    def load_excel_file(self, file_path: str) -> ExcelData:
        """
        Load Excel file and parse all worksheets.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            ExcelData object with parsed worksheets or empty if failed
        """
        try:
            if not file_path:
                raise ValueError("File path is empty")
            
            if not Path(file_path).exists():
                raise FileNotFoundError(f"File does not exist: {file_path}")
            
            self.logger.info(f"Loading Excel file: {file_path}")
            
            # Read all sheets from Excel file
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            self.logger.info(f"Found {len(sheet_names)} worksheets: {sheet_names}")
            
            worksheets = []
            for sheet_name in sheet_names:
                worksheet_result = self._parse_worksheet(excel_file, sheet_name)
                worksheets.append(worksheet_result)
            
            excel_data = ExcelData(
                file_path=file_path,
                worksheets=worksheets
            )
            
            self.logger.info(f"Excel file loaded successfully. {excel_data.successful_worksheets}/{excel_data.total_worksheets} worksheets parsed, {excel_data.total_people_parsed} people total.")
            
            return excel_data
            
        except Exception as e:
            error_msg = f"Error loading Excel file: {str(e)}"
            self.logger.error(f"{error_msg}\n{traceback.format_exc()}")
            raise Exception(error_msg)
    
    def _parse_worksheet(self, excel_file: pd.ExcelFile, sheet_name: str) -> WorksheetParsingResult:
        """
        Parse a single worksheet and extract metadata and data.
        
        Args:
            excel_file: Pandas ExcelFile object
            sheet_name: Name of the sheet to parse
            
        Returns:
            WorksheetParsingResult with parsed data and metadata
        """
        metadata = WorksheetMetadata(sheet_name=sheet_name)
        result = WorksheetParsingResult(metadata=metadata)
        
        try:
            # Read the entire sheet first to extract metadata
            sheet_data = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            
            # Extract metadata from specific cells
            try:
                if len(sheet_data) > METADATA_ROW_FECHA_SOLICITUD and len(sheet_data.columns) > METADATA_COL_FECHA_SOLICITUD:
                    metadata.fecha_solicitud = str(sheet_data.iloc[METADATA_ROW_FECHA_SOLICITUD, METADATA_COL_FECHA_SOLICITUD]) if pd.notna(sheet_data.iloc[METADATA_ROW_FECHA_SOLICITUD, METADATA_COL_FECHA_SOLICITUD]) else ""
                
                if len(sheet_data) > METADATA_ROW_TIENDA and len(sheet_data.columns) > METADATA_COL_TIENDA:
                    metadata.tienda = str(sheet_data.iloc[METADATA_ROW_TIENDA, METADATA_COL_TIENDA]) if pd.notna(sheet_data.iloc[METADATA_ROW_TIENDA, METADATA_COL_TIENDA]) else ""
                
                if len(sheet_data) > METADATA_ROW_ADMINISTRADOR and len(sheet_data.columns) > METADATA_COL_ADMINISTRADOR:
                    metadata.administrador = str(sheet_data.iloc[METADATA_ROW_ADMINISTRADOR, METADATA_COL_ADMINISTRADOR]) if pd.notna(sheet_data.iloc[METADATA_ROW_ADMINISTRADOR, METADATA_COL_ADMINISTRADOR]) else ""
                    
            except Exception as meta_error:
                result.errors.append(f"Error extracting metadata: {str(meta_error)}")
                self.logger.warning(f"Metadata extraction error in sheet '{sheet_name}': {str(meta_error)}")
            
            # Extract main data (columns B through I) starting from configured row
            try:
                if len(sheet_data) > DATA_START_ROW:  # Ensure we have data rows
                    # Read main data: from data start row onwards, columns B through I
                    main_data_rows = sheet_data.iloc[DATA_START_ROW:, IGNORE_COLUMN_INDEX + 1:MAIN_DATA_END_COLUMN + 1]
                    
                    # Set the first data row as headers for main data
                    if len(main_data_rows) > 0:
                        headers = main_data_rows.iloc[0]  # First row becomes headers
                        main_data_rows = main_data_rows.iloc[1:]  # Remaining rows are data
                        main_data_rows.columns = headers
                        
                        # Remove completely empty rows from main data
                        main_data_rows = main_data_rows.dropna(how='all')
                        
                        # Check for missing DNI in rows with data
                        if 'DNI' in main_data_rows.columns or any('dni' in str(col).lower() for col in main_data_rows.columns):
                            # Find DNI column (case insensitive)
                            dni_col = None
                            for col in main_data_rows.columns:
                                if 'dni' in str(col).lower():
                                    dni_col = col
                                    break
                            
                            if dni_col is not None:
                                # Check for missing DNI in non-empty rows
                                non_empty_rows = main_data_rows.dropna(how='all')
                                missing_dni_count = non_empty_rows[dni_col].isna().sum()
                                
                                if missing_dni_count > 0:
                                    result.errors.append(f"{missing_dni_count} rows with data are missing DNI")
                        else:
                            result.warnings.append("No DNI column found - cannot validate DNI completeness")
                        
                        # Extract uniform data (columns J through R) starting from row 8
                        uniform_data = None
                        if len(sheet_data) > UNIFORM_DATA_START_ROW and len(sheet_data.columns) > UNIFORM_DATA_END_COLUMN:
                            uniform_data_rows = sheet_data.iloc[UNIFORM_DATA_START_ROW:, UNIFORM_DATA_START_COLUMN:UNIFORM_DATA_END_COLUMN + 1]
                            
                            # Set column names for uniform data
                            if len(uniform_data_rows.columns) == len(UNIFORM_COLUMN_NAMES):
                                uniform_data_rows.columns = UNIFORM_COLUMN_NAMES
                                
                                # Keep only rows that correspond to main data (same indexes after cleaning)
                                # Reset index to align with main data
                                uniform_data_rows = uniform_data_rows.reset_index(drop=True)
                                main_data_rows = main_data_rows.reset_index(drop=True)
                                
                                # Take only the rows that correspond to the cleaned main data
                                if len(uniform_data_rows) >= len(main_data_rows):
                                    uniform_data = uniform_data_rows.iloc[:len(main_data_rows)].copy()
                                else:
                                    uniform_data = uniform_data_rows.copy()
                                    result.warnings.append(f"Uniform data has fewer rows ({len(uniform_data_rows)}) than main data ({len(main_data_rows)})")
                            else:
                                result.warnings.append(f"Uniform data has {len(uniform_data_rows.columns)} columns, expected {len(UNIFORM_COLUMN_NAMES)}")
                        else:
                            result.warnings.append("Insufficient data for uniform columns (J-R)")
                        
                        result.data = main_data_rows
                        result.uniform_data = uniform_data
                        result.people_parsed = len(main_data_rows)  # After cleaning empty rows
                        result.total_lines = len(sheet_data)  # Total lines in sheet
                        
                        self.logger.info(f"Sheet '{sheet_name}': {result.people_parsed} people parsed from {result.total_lines} total lines")
                        if uniform_data is not None:
                            self.logger.info(f"Sheet '{sheet_name}': {len(uniform_data)} uniform data rows extracted")
                    else:
                        result.warnings.append("No data rows found after headers")
                else:
                    result.warnings.append(f"Sheet has insufficient rows (less than {DATA_START_ROW + 1})")
                    
            except Exception as data_error:
                result.errors.append(f"Error extracting data: {str(data_error)}")
                self.logger.error(f"Data extraction error in sheet '{sheet_name}': {str(data_error)}")
            
            # Validate metadata
            if not metadata.fecha_solicitud:
                result.warnings.append("Missing fecha_solicitud (C3)")
            if not metadata.tienda:
                result.warnings.append("Missing tienda (C4)")
            if not metadata.administrador:
                result.warnings.append("Missing administrador (C5)")
            
        except Exception as e:
            result.errors.append(f"Critical error parsing worksheet: {str(e)}")
            self.logger.error(f"Critical error parsing worksheet '{sheet_name}': {str(e)}\n{traceback.format_exc()}")
        
        return result
    
    def validate_excel_data(self, excel_data: ExcelData) -> ExcelValidationResult:
        """
        Validate Excel data for processing.
        
        Args:
            excel_data: ExcelData object to validate
            
        Returns:
            ExcelValidationResult with validation details
        """
        result = ExcelValidationResult(is_valid=False)
        
        try:
            if not excel_data.is_loaded:
                result.errors.append("No Excel worksheets could be loaded")
                result.message = "Excel file could not be loaded or has no valid worksheets"
                return result
            
            if excel_data.total_worksheets == 0:
                result.errors.append("Excel file contains no worksheets")
                result.message = "Excel file is empty or corrupted"
                return result
            
            # Validate each worksheet
            total_errors = 0
            total_warnings = 0
            
            for worksheet in excel_data.worksheets:
                if worksheet.errors:
                    result.errors.extend([f"Sheet '{worksheet.metadata.sheet_name}': {error}" for error in worksheet.errors])
                    total_errors += len(worksheet.errors)
                
                if worksheet.warnings:
                    result.warnings.extend([f"Sheet '{worksheet.metadata.sheet_name}': {warning}" for warning in worksheet.warnings])
                    total_warnings += len(worksheet.warnings)
            
            # Check if we have at least one successful worksheet
            if excel_data.successful_worksheets == 0:
                result.errors.append("No worksheets could be parsed successfully")
                result.message = f"All {excel_data.total_worksheets} worksheets failed to parse"
                return result
            
            # Validation passes if we have at least one successful worksheet
            result.is_valid = True
            
            # Create summary message
            success_msg = f"Excel file validated: {excel_data.successful_worksheets}/{excel_data.total_worksheets} worksheets parsed successfully"
            success_msg += f", {excel_data.total_people_parsed} people total"
            
            if total_errors > 0:
                success_msg += f", {total_errors} errors"
            if total_warnings > 0:
                success_msg += f", {total_warnings} warnings"
                
            result.message = success_msg
            
            return result
            
        except Exception as e:
            result.errors.append(f"Validation error: {str(e)}")
            result.message = "Unexpected error during validation"
            self.logger.error(f"Excel validation error: {str(e)}\n{traceback.format_exc()}")
            return result


class FileGenerationService:
    """Service for generating files from Excel data using templates."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
    
    
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
            
            # Validate template files
            template_errors = TemplateValidator.validate_template_files(config)
            if template_errors:
                result.message = "Template validation failed"
                result.errors.extend(template_errors)
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
            result.message = f"Ready to process {excel_data.total_people_parsed} people with CARGO and AUTORIZACION templates"
            
            self.logger.info(result.message)
            
            return result
            
        except Exception as e:
            error_msg = f"Error during file generation: {str(e)}"
            self.logger.error(f"{error_msg}\n{traceback.format_exc()}")
            result.message = error_msg
            result.errors.append(error_msg)
            return result


class ConfigService:
    """Service for handling application configuration with persistence."""
    
    def __init__(self):
        self.config_manager = ConfigManager()
        self.config = self.config_manager.load_config()
    
    def get_config(self) -> AppConfig:
        """Get current configuration."""
        return self.config
    
    def update_config(self, **kwargs) -> bool:
        """
        Update configuration values and save to file.
        
        Returns:
            True if saved successfully, False otherwise
        """
        return self.config_manager.update_and_save(self.config, **kwargs)
    
    def save_config(self) -> bool:
        """
        Save current configuration to file.
        
        Returns:
            True if saved successfully, False otherwise
        """
        return self.config_manager.save_config(self.config)
    
    def validate_paths(self) -> List[str]:
        """
        Validate configuration paths.
        
        Returns:
            List of validation errors
        """
        # Use centralized template validator
        return TemplateValidator.validate_template_files(self.config)
