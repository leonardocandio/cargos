# File Generator - Excel to PDF

A simple GUI application that generates PDF files from Excel data using Word templates.

## Features

- **Tabbed Interface**: Separate tabs for "Cargos" and "Stock" functionality
- **Excel File Loading**: Load and preview Excel files with data validation
- **Template Management**: Use Word document templates for file generation
- **Configurable Paths**: Set default paths for templates and output files
- **Logging**: Built-in logging system with error handling
- **Standalone Executable**: Package as a single executable file for multiple platforms

## Setup and Installation

### Requirements
- Python 3.7+
- Windows OS (for .exe packaging) or Unix-like systems (for binary packaging)

### Quick Start

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the Application**:
   ```bash
   python main.py
   ```

3. **Build Executable**:
   ```bash
   # Run the build script
   build.bat
   
   # Or manually with PyInstaller
   pyinstaller --onefile --windowed --name="FileGenerator" --add-data "templates;templates" main.py
   ```

## Usage

### Cargos Tab
1. **Select Excel File**: Browse and load your Excel file containing the data
2. **Set CARGO Template**: Select the CARGO UNIFORMES Word document template
3. **Set AUTORIZACION Template**: Select the AUTORIZACION DESCUENTO Word document template
4. **Set Destination Folder**: Choose where generated files will be saved (creates subfolders automatically)
5. **Load Excel**: Preview and validate the Excel data in the application
6. **Generate Files**: Process the data and create PDF files (logic TBD)

### Key Features
- **Persistent Configuration**: Template paths and settings are automatically saved
- **Excel Validation**: Comprehensive validation with detailed error reporting
- **UI-Only Interaction**: All errors and messages are shown in the GUI (no terminal windows)
- **Automatic Folder Creation**: Destination folder and subfolders are created automatically

### File Structure
```
cargos/
├── main.py                    # Application controller and entry point
├── models.py                  # Data models and configuration
├── services.py                # Business logic services
├── ui_components.py           # UI components and widgets
├── config_manager.py          # Configuration persistence
├── requirements.txt           # Python dependencies
├── build.bat                 # Build script for executable
├── config.json               # Persistent configuration (auto-generated)
├── templates/                # Word document templates
│   ├── CARGO UNIFORMES.docx
│   └── 50% - AUTORIZACIÓN DESCUENTO DE UNIFORMES (02).docx
├── sources/                  # Sample Excel files
│   └── REQUERIMIENTO DE UNIFORMES nuevo.xlsx
└── output/                   # Generated files destination
```

## Architecture

The application follows clean architecture principles with clear separation of concerns:

### **Core Components**
- **`main.py`**: Application controller and entry point
- **`models.py`**: Data models and configuration classes
- **`services.py`**: Business logic services (Excel processing, file generation)
- **`ui_components.py`**: UI components and widgets

### **Design Principles**
- **Separation of Concerns**: Each class has a single responsibility
- **Dependency Injection**: Services are injected into controllers
- **Type Safety**: Full type hints for better maintainability
- **Error Handling**: Comprehensive exception handling at all levels
- **Extensibility**: Easy to add new tabs and functionality

### **Class Structure**
```
FileGeneratorApp (Controller)
├── ConfigService (Configuration management)
├── ExcelService (Excel file operations)
├── FileGenerationService (File generation logic)
└── CargosTab (UI Components)
    ├── FileSelectionFrame
    ├── ControlFrame
    ├── DataPreviewFrame
    └── LogFrame
```

### Next Steps
- Implement Excel parsing logic based on your data structure
- Add Word document template processing
- Implement PDF generation functionality
- Add Stock tab functionality

## Building for Distribution

The application can be packaged as a standalone executable using PyInstaller:

```bash
pyinstaller --onefile --windowed --name="FileGenerator" --add-data "templates;templates" main.py
```

This creates a single executable file that includes all dependencies and can run on the target platform without Python installed.

### Cross-Platform Building

**Windows:**
```bash
build.bat
```

**Unix/Linux/macOS:**
```bash
./build.sh
```
