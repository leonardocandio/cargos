"""
Application constants and configuration values.
"""

# Excel parsing constants
METADATA_ROW_FECHA_SOLICITUD = 2  # C3 (0-indexed: row 2, col 2)
METADATA_COL_FECHA_SOLICITUD = 2
METADATA_ROW_TIENDA = 3  # C4
METADATA_COL_TIENDA = 2
METADATA_ROW_ADMINISTRADOR = 4  # C5
METADATA_COL_ADMINISTRADOR = 2

# Data extraction constants
DATA_START_ROW = 6  # Row 7 in Excel (0-indexed)
IGNORE_COLUMN_INDEX = 0  # Column A
MAIN_DATA_END_COLUMN = 8  # Column I (0-indexed, inclusive)

# Uniform data constants
UNIFORM_DATA_START_ROW = 7  # Row 8 in Excel (0-indexed)
UNIFORM_DATA_START_COLUMN = 9  # Column J (0-indexed)
UNIFORM_DATA_END_COLUMN = 17  # Column R (0-indexed)
UNIFORM_COLUMN_NAMES = [
    "camisa", "blusa", "mandilon", "andarin", 
    "deliverypolo", "deliverycasaca", "deliverygorro", "packergorra", "packerpolo"
]

# UI constants
DEFAULT_WINDOW_SIZE = "700x850"
DEFAULT_PREVIEW_ROWS = 100
DEFAULT_TREE_HEIGHT = 8

# Dialog constants
GENERATION_DIALOG_WIDTH = 600
GENERATION_DIALOG_HEIGHT = 600
GENERATION_DIALOG_CANVAS_HEIGHT = 400

# Tree column widths
TREE_COLUMN_WIDTH_PEOPLE = 120
TREE_COLUMN_WIDTH_STATUS = 100
TREE_COLUMN_WIDTH_UNIFORM = 100
TREE_COLUMN_WIDTH_DATA = 100

# Spanish month names for CARGO documents
SPANISH_MONTHS = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto", 
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}

# File paths
DEFAULT_TEMPLATES_DIR = "templates"
DEFAULT_OUTPUT_DIR = "output"
DEFAULT_LOG_FILE = "app.log"

IGNORE_QUANTITY_IN_PRICING = True