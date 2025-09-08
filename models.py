"""
Data models and configuration classes for the File Generator application.
"""
from dataclasses import dataclass
from typing import Optional
import pandas as pd
from pathlib import Path


@dataclass
class AppConfig:
    """Application configuration settings."""
    templates_path: str = "templates/"
    destination_path: str = "output/"
    excel_file_path: str = ""
    log_file: str = "app.log"
    preview_rows_limit: int = 100


@dataclass
class ExcelData:
    """Container for Excel file data and metadata."""
    data: Optional[pd.DataFrame] = None
    file_path: str = ""
    total_rows: int = 0
    columns: list = None
    
    def __post_init__(self):
        if self.data is not None and self.columns is None:
            self.columns = list(self.data.columns)
            self.total_rows = len(self.data)
    
    @property
    def is_loaded(self) -> bool:
        """Check if data is loaded."""
        return self.data is not None and not self.data.empty
    
    @property
    def preview_data(self) -> pd.DataFrame:
        """Get preview data (limited rows)."""
        if not self.is_loaded:
            return pd.DataFrame()
        return self.data.head(100)


@dataclass
class GenerationResult:
    """Result of file generation process."""
    success: bool
    files_generated: int = 0
    errors: list = None
    message: str = ""
    
    def __post_init__(self):
        if self.errors is None:
            self.errors = []
