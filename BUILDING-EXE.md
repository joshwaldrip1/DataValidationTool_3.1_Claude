Windows EXE Build (PyInstaller)

Prereqs
- Windows 10/11, 64-bit Python (same bitness as Office recommended)
- Excel installed (for COM at runtime)
- Optional: a virtualenv in `venv/`

1) Activate venv and install dependencies

```
venv\Scripts\activate
pip install -U pip wheel
pip install pyinstaller pandas openpyxl pywin32 tkinterdnd2 pypdf
```

2) Build the EXE

Run one of the scripts from the repo root:

- One-file (single EXE):
  - CMD: `build\build_exe_onefile.bat`
  - PowerShell: `build\build_exe_onefile.ps1`
  - Output: `dist\DataValidationTool.exe`

- One-dir (folder with EXE + files):
  - CMD: `build\build_exe.bat`
  - PowerShell: `build\build_exe.ps1`
  - Output: `dist\DataValidationTool\DataValidationTool.exe`

Notes
- The command adds hidden imports/collections for `win32com`, `openpyxl`, `pandas`, and `tkinterdnd2` so that PyInstaller bundles them correctly.
- One-file EXE starts slightly slower (it unpacks to a temp dir at runtime). If AV complains, prefer one-dir.
- If Excel macro injection fails, enable: Excel → Options → Trust Center → Trust Center Settings → Macro Settings → “Trust access to the VBA project object model”.
- If drag & drop isn’t available, install `tkinterdnd2` or continue without it — the app falls back gracefully.

Troubleshooting
- Missing DLLs or runtime errors: try `--onefile` off (use `--onedir` as provided).
- COM errors: Make sure you are running 64-bit Python if you have 64-bit Office.
- Antivirus false positives: prefer `--onedir` distribution.
