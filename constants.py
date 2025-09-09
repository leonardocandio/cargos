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
    "deliverypolo", "deliverycasaca", "packergorra", "packerpolo", "aaa"
]

# UI constants
DEFAULT_WINDOW_SIZE = "950x750"
DEFAULT_PREVIEW_ROWS = 100
DEFAULT_TREE_HEIGHT = 8

# File paths
DEFAULT_TEMPLATES_DIR = "templates"
DEFAULT_OUTPUT_DIR = "output"
DEFAULT_LOG_FILE = "app.log"
