# DataValidationTool v3.1

## Project Overview

Windows desktop application for validating surveying/topographic CSV data against XML Field Code Libraries (FXL files). Built for professionals using Trimble surveying equipment. Validates geospatial coordinates, field codes, photos, and pipe/infrastructure data.

## Tech Stack

- **Language**: Python 3.13 (type hints, Protocol classes)
- **GUI**: tkinter + ttk (with optional tkinterdnd2 for drag-and-drop)
- **Excel**: openpyxl (read/write), win32com.client (COM automation + VBA injection)
- **Data**: pandas (CSV processing), xml.etree.ElementTree (FXL/JXL parsing)
- **Build**: PyInstaller (EXE), Inno Setup (installer)
- **Platform**: Windows-only (COM, Task Scheduler, Outlook integration)

## Architecture

**Single-file monolith** (`DataValidationTool-v3.1.py`, ~8,200 lines) — deliberate choice for portability as a standalone EXE. Contains:

- `DataValidationTool(TkBase)` — main class with 90+ methods
- `ToolTip` — tooltip helper class
- `_HeadlessCRDBRunner` — background CRDB monitoring
- Protocol classes: `RangeLike`, `WorksheetLike`, `ExcelApplicationLike`

## Development

```bash
# Activate virtual environment
venv\Scripts\activate

# Install dependencies
pip install -U pip wheel
pip install pyinstaller pandas openpyxl pywin32 tkinterdnd2 pypdf

# Run from source
python DataValidationTool-v3.1.py

# Build EXE (one-directory, preferred)
build\build_exe.bat          # or build\build_exe.ps1

# Build EXE (single file)
build\build_exe_onefile.bat  # or build\build_exe_onefile.ps1

# Background CRDB check (via Task Scheduler)
DataValidationTool.exe --crdb-check
```

## Key Files

| File | Purpose |
|------|---------|
| `DataValidationTool-v3.1.py` | Entire application source |
| `config.json` | User settings: FXL paths, numeric bounds, preferences |
| `validation_log.json` | Audit trail of validation sessions |
| `crdb_watchlist.json` | CRDB files tracked for background monitoring |
| `DataValidationTool.spec` | PyInstaller build configuration |
| `DataValidationTool_Setup.iss` | Inno Setup installer script |

## Code Conventions

- **Naming**: public methods `validate_in_excel()`, private `_load_fxl_path()`, temp `_tmpdir`
- **Types**: Python 3.13+ style (`dict[str, Any]`, `list[int]`, `X | None`)
- **Protocols**: Runtime-checkable protocols for typed COM objects
- **Error handling**: Silent failures for non-critical paths (config, logging); user-facing dialogs for validation/file errors
- **Config**: JSON-based persistence (`config.json`)
- **VBA**: Embedded as raw string literals, injected into Excel workbooks at runtime
- **Coordinates**: NEZ format at 10 decimal places

## Data Flow

1. User loads CSV (surveying data) + FXL (field code library)
2. Tool parses FXL for valid field codes and attributes
3. Row-by-row validation: field codes, numeric bounds, coordinate checks
4. Errors exported to color-coded XLSM report (via openpyxl + VBA injection)
5. Corrected data exported as clean CSV

## Build Requirements

- Windows 10/11, 64-bit Python matching Office bitness
- Excel installed (COM automation at runtime)
- Excel Trust Center: "Trust access to VBA project object model" enabled
- One-dir build preferred over one-file (AV compatibility)

## Important Notes

- FXL library path is a network UNC path (`S:\TOPOGRAPHIC DATA\...`)
- COM automation requires matching Python/Office bitness (both 32-bit or both 64-bit)
- The app gracefully degrades without tkinterdnd2 (no drag-and-drop, but functional)
- Avoid splitting the monolith — single-file design is intentional for EXE distribution
