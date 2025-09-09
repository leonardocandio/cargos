"""
Data models and configuration classes for the File Generator application.
"""
from dataclasses import dataclass
from typing import Optional, TYPE_CHECKING
from constants import DEFAULT_OUTPUT_DIR, DEFAULT_LOG_FILE, DEFAULT_PREVIEW_ROWS

if TYPE_CHECKING:
    import pandas as pd


@dataclass
class AppConfig:
    """Application configuration settings."""
    destination_path: str = f"{DEFAULT_OUTPUT_DIR}/"
    excel_file_path: str = ""
    cargo_template_path: str = "templates/CARGO UNIFORMES.docx"
    autorizacion_template_path: str = "templates/50% - AUTORIZACIÓN DESCUENTO DE UNIFORMES (02).docx"
    log_file: str = DEFAULT_LOG_FILE
    preview_rows_limit: int = DEFAULT_PREVIEW_ROWS


@dataclass
class WorksheetMetadata:
    """Metadata extracted from a worksheet."""
    sheet_name: str
    fecha_solicitud: str = ""
    tienda: str = ""
    administrador: str = ""
    
    @property
    def identifier(self) -> str:
        """Get worksheet identifier."""
        return self.sheet_name


@dataclass
class WorksheetParsingResult:
    """Result of parsing a single worksheet."""
    metadata: WorksheetMetadata
    data: Optional['pd.DataFrame'] = None  # Main data (columns B through I)
    uniform_data: Optional['pd.DataFrame'] = None  # Uniform data (columns J through R)
    total_lines: int = 0
    people_parsed: int = 0
    errors: list = None
    warnings: list = None
    
    def __post_init__(self):
        if self.errors is None:
            self.errors = []
        if self.warnings is None:
            self.warnings = []
        # total_lines and people_parsed are set directly in services.py after processing


@dataclass
class ExcelData:
    """Container for Excel file data and metadata."""
    file_path: str = ""
    worksheets: list = None  # List[WorksheetParsingResult]
    total_worksheets: int = 0
    successful_worksheets: int = 0
    
    def __post_init__(self):
        if self.worksheets is None:
            self.worksheets = []
        self.total_worksheets = len(self.worksheets)
        self.successful_worksheets = len([w for w in self.worksheets if w.data is not None])
    
    @property
    def is_loaded(self) -> bool:
        """Check if data is loaded."""
        return len(self.worksheets) > 0 and self.successful_worksheets > 0
    
    @property
    def total_people_parsed(self) -> int:
        """Get total number of people parsed across all worksheets."""
        return sum(w.people_parsed for w in self.worksheets)
    
    @property
    def total_errors(self) -> int:
        """Get total number of errors across all worksheets."""
        return sum(len(w.errors) for w in self.worksheets)
    
    def get_worksheet_by_name(self, sheet_name: str) -> Optional['WorksheetParsingResult']:
        """Get worksheet by name."""
        for worksheet in self.worksheets:
            if worksheet.metadata.sheet_name == sheet_name:
                return worksheet
        return None


@dataclass
class ExcelValidationResult:
    """Result of Excel file validation."""
    is_valid: bool
    errors: list = None
    warnings: list = None
    message: str = ""
    
    def __post_init__(self):
        if self.errors is None:
            self.errors = []
        if self.warnings is None:
            self.warnings = []


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
