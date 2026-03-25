# Enhancements Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add 13 UX, validation, and audit enhancements to `DataValidationTool-v3.1.py`.

**Architecture:** All changes live in the single main file. Python changes are additive (new methods, expanded existing methods). VBA changes are string edits inside the embedded `vbcode` string. Config extensions read from/write to `config.json`.

**Tech Stack:** Python 3.13, tkinter + ttk, win32com, openpyxl, pandas, config.json, embedded VBA string.

---

## Chunk 1: Config, Persistence, Tooltips

### Task 1: Remember Last FXL Path

**Files:**
- Modify: `DataValidationTool-v3.1.py` — `_load_app_config`, `_load_fxl_path`, add `_save_app_config`
- Modify: `config.json` — no structural change needed; key written at runtime

- [ ] **Step 1: Add `_save_app_config` helper after `_load_app_config` (~line 246)**

```python
def _save_app_config(self, updates: dict) -> None:
    """Persist key/value updates to config.json next to the app/exe."""
    _base = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
    cfg_path = os.path.join(_base, "config.json")
    cfg: dict = {}
    try:
        if os.path.isfile(cfg_path):
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
    except Exception:
        pass
    cfg.update(updates)
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass
```

- [ ] **Step 2: In `_load_app_config`, after reading `single_excel_instance`, read `last_fxl_path`**

```python
raw_fxl: Any = cfg.get("last_fxl_path", None)
if isinstance(raw_fxl, str) and raw_fxl.strip() and os.path.isfile(raw_fxl.strip()):
    self._pending_fxl_path = raw_fxl.strip()
else:
    self._pending_fxl_path = None
```

- [ ] **Step 3: Add `self._pending_fxl_path: str | None = None` to `__init__` state block (~line 127)**

- [ ] **Step 4: At end of `_build_ui` (before return), schedule deferred FXL load**

```python
# Auto-load last FXL after UI is ready
if getattr(self, "_pending_fxl_path", None):
    self.after(200, lambda: self._load_fxl_path(self._pending_fxl_path, silent=True))
```

- [ ] **Step 5: In `_load_fxl_path`, after `self.fxl_path = path`, save to config**

```python
try:
    self._save_app_config({"last_fxl_path": path})
except Exception:
    pass
```

- [ ] **Step 6: Manual test** — Load an FXL, close the app, reopen. Status bar should show "Loaded FXL: …" with the previously used file.

---

### Task 2: Tooltips on Disabled Buttons

**Files:**
- Modify: `DataValidationTool-v3.1.py` — add `ToolTip` class near top of file (after imports), wire in `_build_ui`

- [ ] **Step 1: Add `ToolTip` class after the Protocol definitions (~line 90)**

```python
class ToolTip:
    """Simple hover tooltip for tkinter widgets."""
    def __init__(self, widget: tk.Widget, text: str) -> None:
        self._widget = widget
        self._text = text
        self._tip: tk.Toplevel | None = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _event: Any = None) -> None:
        if self._tip:
            return
        x = self._widget.winfo_rootx() + 20
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._tip = tk.Toplevel(self._widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(
            self._tip, text=self._text, justify="left",
            background="#ffffe0", relief="solid", borderwidth=1,
            wraplength=280, padx=4, pady=2,
        )
        lbl.pack()

    def _hide(self, _event: Any = None) -> None:
        if self._tip:
            self._tip.destroy()
            self._tip = None
```

- [ ] **Step 2: In `_build_ui`, after creating each button, attach tooltips**

```python
ToolTip(self.btn_validate_gui, "Load a CSV and FXL first — Excel will open and this button activates.")
ToolTip(self.btn_email, "Validate the sheet first, then use this to email the error report via Outlook.")
ToolTip(self.btn_export_outputs, "Exports corrected CSV and error report (.xlsm) to the original CSV folder.")
ToolTip(self.btn_missing_heats, "Load an MTR spreadsheet (drag & drop an Excel file with 'MTR' in the name) to enable.")
```

- [ ] **Step 3: Manual test** — Hover over each greyed-out button; tooltip should appear within 1 second.

---

### Task 3: Progress Bar During Validation

**Files:**
- Modify: `DataValidationTool-v3.1.py` — `_build_ui`, `validate_in_excel`, `_export_and_open_excel`

- [ ] **Step 1: Add `from tkinter import ttk` to imports line**

- [ ] **Step 2: In `_build_ui`, add a progress bar below the status label**

```python
self.progress = ttk.Progressbar(self, mode="indeterminate", length=200)
self.progress.pack(pady=2)
self.progress.pack_forget()  # hidden until needed
```

- [ ] **Step 3: Add helper methods `_progress_start` and `_progress_stop`**

```python
def _progress_start(self, msg: str = "Working…") -> None:
    self.status.config(text=msg)
    self.progress.pack(pady=2)
    self.progress.start(12)
    self.update_idletasks()

def _progress_stop(self) -> None:
    self.progress.stop()
    self.progress.pack_forget()
    self.update_idletasks()
```

- [ ] **Step 4: In `validate_in_excel`, wrap the `self._excel.Run(target)` call**

```python
self._progress_start("Running full validation…")
try:
    ...  # existing Run loop
finally:
    self._progress_stop()
    self.status.config(text="Validation complete.")
```

- [ ] **Step 5: In `_export_and_open_excel`, wrap the Excel open/VBA-inject section**

```python
self._progress_start(f"Opening {csv_name} in Excel…")
# ... existing code ...
# at the end of the try block, before setting self._excel:
self._progress_stop()
self.status.config(text=f"Ready — {csv_name} validated.")
```

In the except block: also call `self._progress_stop()`.

- [ ] **Step 6: Manual test** — Drop a large CSV; animated bar should appear while Excel is loading and disappear when done.

---

## Chunk 2: Batch Processing & Folder Mode

### Task 4: Batch Summary Window

**Files:**
- Modify: `DataValidationTool-v3.1.py` — `_on_drop`, add `_show_batch_summary`

- [ ] **Step 1: Add `_show_batch_summary` method**

```python
def _show_batch_summary(self, results: list[tuple[str, str]]) -> None:
    """Show a Toplevel table: [(filename, status_message), ...]"""
    if not results:
        return
    win = tk.Toplevel(self)
    win.title("Batch Processing Summary")
    win.resizable(True, True)
    tk.Label(win, text=f"Processed {len(results)} file(s):", font=("", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 2))
    frame = tk.Frame(win)
    frame.pack(fill="both", expand=True, padx=10, pady=4)
    sb = tk.Scrollbar(frame)
    sb.pack(side="right", fill="y")
    cols = ("File", "Result")
    tree = ttk.Treeview(frame, columns=cols, show="headings", height=min(len(results), 15), yscrollcommand=sb.set)
    tree.heading("File", text="File")
    tree.heading("Result", text="Result")
    tree.column("File", width=220, anchor="w")
    tree.column("Result", width=340, anchor="w")
    sb.config(command=tree.yview)
    for fname, status in results:
        tag = "ok" if status.lower().startswith("ok") else "err"
        tree.insert("", "end", values=(fname, status), tags=(tag,))
    tree.tag_configure("ok", foreground="#006600")
    tree.tag_configure("err", foreground="#cc0000")
    tree.pack(side="left", fill="both", expand=True)
    tk.Button(win, text="Close", command=win.destroy).pack(pady=8)
    win.grab_set()
```

- [ ] **Step 2: Modify the CSV-batch loop in `_on_drop` to collect results**

Replace the existing loop:
```python
batch_results: list[tuple[str, str]] = []
for csv_p in csvs:
    fname = os.path.basename(csv_p)
    try:
        self._process_pair(csv_p, drop_fxl)
        batch_results.append((fname, "OK — opened in Excel"))
    except Exception as e:
        self.status.config(text=f"Error processing {fname}: {e}")
        batch_results.append((fname, f"Error: {e}"))
if len(batch_results) > 1:
    self._show_batch_summary(batch_results)
```

- [ ] **Step 3: Manual test** — Drop 2+ CSVs with an FXL. After processing, a summary window should appear.

---

### Task 5: Process Folder Button

**Files:**
- Modify: `DataValidationTool-v3.1.py` — `_build_ui`, add `process_folder`

- [ ] **Step 1: Add "Process Folder…" button in `_build_ui` (second row)**

```python
btns2 = tk.Frame(self)
btns2.pack(pady=2)

self.btn_process_folder = tk.Button(
    btns2, text="Process Folder…", width=22,
    command=self.process_folder
)
self.btn_process_folder.grid(row=0, column=0, padx=8)
ToolTip(self.btn_process_folder, "Choose a folder — validates all CSV files found using the available FXL.")
```

- [ ] **Step 2: Add `process_folder` method**

```python
def process_folder(self) -> None:
    """Walk a folder, pair each CSV with the best FXL, process all."""
    folder = filedialog.askdirectory(title="Select folder to process")
    if not folder:
        return

    csvs = sorted(
        os.path.join(folder, f) for f in os.listdir(folder)
        if f.lower().endswith(".csv") and not f.lower().endswith("_corrected.csv")
    )
    if not csvs:
        messagebox.showinfo("No CSVs", "No CSV files found in the selected folder.")
        return

    # Find FXLs in folder
    fxls = sorted(os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith((".fxl", ".xml")))
    # Fallback to already-loaded FXL
    if not fxls and self._initial_fxl_path:
        fxls = [self._initial_fxl_path]
    if not fxls:
        p = filedialog.askopenfilename(
            title="Select FXL for this folder",
            initialdir=folder,
            filetypes=[("FXL files", "*.fxl"), ("XML files", "*.xml")]
        )
        if not p:
            return
        fxls = [p]

    # Use the first/only FXL; if multiple, prompt
    fxl_to_use = fxls[0]
    if len(fxls) > 1:
        from tkinter.simpledialog import askstring
        names = "\n".join(os.path.basename(f) for f in fxls)
        choice = messagebox.askyesno(
            "Multiple FXLs",
            f"Found {len(fxls)} FXL files:\n{names}\n\nUse '{os.path.basename(fxls[0])}'?",
        )
        if not choice:
            return

    confirm = messagebox.askyesno(
        "Process Folder",
        f"Process {len(csvs)} CSV file(s) in:\n{folder}\n\nUsing FXL: {os.path.basename(fxl_to_use)}\n\nContinue?"
    )
    if not confirm:
        return

    batch_results: list[tuple[str, str]] = []
    for csv_p in csvs:
        fname = os.path.basename(csv_p)
        self._progress_start(f"Processing {fname}…")
        try:
            self._process_pair(csv_p, fxl_to_use)
            batch_results.append((fname, "OK — opened in Excel"))
        except Exception as e:
            batch_results.append((fname, f"Error: {e}"))
        finally:
            self._progress_stop()

    self._show_batch_summary(batch_results)
```

- [ ] **Step 3: Manual test** — Click "Process Folder…", pick a folder with CSVs. Each CSV should open in Excel; summary window should appear at end.

---

## Chunk 3: File Loading Improvements

### Task 6: Auto-Detect FXL from Parent Folder

**Files:**
- Modify: `DataValidationTool-v3.1.py` — `_ensure_fxl_after_csv`

- [ ] **Step 1: In `_ensure_fxl_after_csv`, after step "2) Single candidate in folder", add step 2b for parent folder**

Insert between the "single candidate in folder" block and the "Ask the user" block:
```python
# 2b) Single candidate in parent folder
parent = os.path.dirname(folder)
if parent and parent != folder:
    parent_cands = [p for p in os.listdir(parent) if p.lower().endswith((".fxl", ".xml"))]
    if len(parent_cands) == 1:
        try:
            self._load_fxl_path(os.path.join(parent, parent_cands[0]))
            self.status.config(text=f"Auto-picked FXL from parent folder: {parent_cands[0]}")
            return True
        except Exception:
            pass
```

- [ ] **Step 2: Manual test** — Place a CSV in a subfolder and the FXL one level up. Drop the CSV; it should auto-pick the FXL from the parent folder without a dialog.

---

### Task 7: Fuzzy Heat Number Matching

**Files:**
- Modify: `DataValidationTool-v3.1.py` — `_collect_csv_heats`, `email_missing_heats`, `_load_mtr_path`

- [ ] **Step 1: In `_collect_csv_heats`, normalize heat values to uppercase stripped strings**

Change `hv = ("" if row[col] is None else str(row[col])).strip()` to:
```python
hv = ("" if row[col] is None else str(row[col])).strip().upper()
```

- [ ] **Step 2: In `_load_mtr_path`, after `.str.strip()` normalization, add uppercase normalization for HEAT column**

After `df_norm[c] = df_norm[c].str.strip()`:
```python
# Normalize HEAT column to uppercase for case-insensitive matching
if "HEAT" in df_norm.columns:
    df_norm["HEAT"] = df_norm["HEAT"].str.upper()
```

- [ ] **Step 3: In `email_missing_heats`, ensure mtr_heats is also uppercased** (already stripped via dataframe, but add `.upper()` for safety)

Change the `mtr_heats` comprehension:
```python
mtr_heats: set[str] = set(
    ("" if pd.isna(x) else str(x)).strip().upper() for x in self.mtr_df.get("HEAT", [])
)
```

- [ ] **Step 4: Manual test** — Create a CSV with heat "a1234" (lowercase) and an MTR with "A1234". Drop both; "Email Missing Heats" should report zero missing heats.

---

### Task 8: Validation Log File

**Files:**
- Modify: `DataValidationTool-v3.1.py` — add `_write_validation_log`, call from `_export_and_open_excel`

- [ ] **Step 1: Add `import datetime` to the imports block**

- [ ] **Step 2: Add `_write_validation_log` method**

```python
def _write_validation_log(self) -> None:
    """Append a JSON line to validation_log.json next to the CSV."""
    try:
        csv_dir = os.path.dirname(self.csv_path) if self.csv_path else None
        if not csv_dir:
            return
        log_path = os.path.join(csv_dir, "validation_log.json")
        entry = {
            "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
            "user": os.environ.get("USERNAME", os.environ.get("USER", "unknown")),
            "csv_file": os.path.basename(self.csv_path or ""),
            "fxl_file": os.path.basename(self.fxl_path or ""),
            "mtr_file": os.path.basename(self.mtr_path or "") if self.mtr_path else "",
        }
        # Append to array in file (or create new array)
        entries: list = []
        if os.path.isfile(log_path):
            try:
                with open(log_path, "r", encoding="utf-8") as f:
                    existing = json.load(f)
                if isinstance(existing, list):
                    entries = existing
            except Exception:
                pass
        entries.append(entry)
        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(entries, f, indent=2)
    except Exception:
        pass  # Log failures must never break the main flow
```

- [ ] **Step 3: In `_export_and_open_excel`, after `self._excel_opened = True`, call `self._write_validation_log()`**

- [ ] **Step 4: Manual test** — Process a CSV. Check that `validation_log.json` appears in the same folder with a timestamped entry.

---

## Chunk 4: UX & Help Features

### Task 9: In-App Help Panel

**Files:**
- Modify: `DataValidationTool-v3.1.py` — `_build_ui`, add `show_help`

- [ ] **Step 1: Add "?" help button in `btns2` frame (beside "Process Folder…")**

```python
self.btn_help = tk.Button(btns2, text="? Help", width=10, command=self.show_help)
self.btn_help.grid(row=0, column=1, padx=8)
```

- [ ] **Step 2: Add `show_help` method**

```python
def show_help(self) -> None:
    """Open a non-modal help window with color legend and usage notes."""
    win = tk.Toplevel(self)
    win.title("Data Validation Tool — Help")
    win.geometry("540x520")
    win.resizable(True, True)

    text = tk.Text(win, wrap="word", padx=10, pady=8, font=("Consolas", 9))
    sb = tk.Scrollbar(win, command=text.yview)
    text.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    text.pack(fill="both", expand=True)

    help_content = """DATA VALIDATION TOOL v3.1 — Quick Reference
═══════════════════════════════════════════

HOW TO USE
──────────
1. Drag & drop a CSV file (field survey data) onto the window.
2. Drag & drop an FXL file (Trimble feature library) — or the tool
   will auto-detect one in the same/parent folder.
3. Excel opens automatically and runs validation.
4. Review colored cells. Hover for error details (comment bubble).
5. Fix errors in Excel or export and send back to field staff.

OPTIONAL: Drag an MTR Excel file (filename must contain "MTR")
to enable material test record cross-checking.

BUTTONS
───────
Validate Entire Sheet   Re-run all checks after editing cells.
Email Report…           Attach the .xlsm report to an Outlook email.
Export CSV + Report     Save corrected CSV and .xlsm to the CSV folder.
Email Missing Heats     Email a list of heats not found in the MTR.
Process Folder…         Batch-validate all CSVs in a chosen folder.

COLOR LEGEND (Excel cells)
──────────────────────────
RED (strong)     Not ALL CAPS; invalid token (NA, UNK, -, _)
RED              Primary validation failure
ORANGE           Field Code not in FXL; MTR mismatch
YELLOW           Duplicate Point Number, Station, or coordinates
PURPLE           Value not in FXL allowed list
LIGHT GREEN      Unusual/NA value in a required list attribute
TEAL             Station format error (expected 0+00 or 0+00.00)
PINK             Joint length statistical outlier (>2 std devs)

RIGHT-CLICK MENU (in Excel)
───────────────────────────
Clear Validation Flag(s)       Remove color/comment from selection.
Ignore All Errors of This Type Remove all instances of that error.
Use MTR value for this cell    Auto-fill from MTR data.
Use MTR value for all Heats    Auto-fill all rows with same heat.

FXL FILE FORMAT
───────────────
Trimble FXL files are XML. The tool supports:
  PointFeatureDefinition, LineFeatureDefinition,
  PolygonFeatureDefinition, Feature, SurveyCode, etc.
Attribute types: List (dropdown), Text, Number, Photo.
Entry method "Required" = cell must not be blank.

MTR SPREADSHEET FORMAT
──────────────────────
Any Excel file with "MTR" in the filename. Required columns
(names are flexible — aliases are recognized):
  HEAT / HEAT NUMBER
  MANUFACTURER
  NOM DIAMETER / OUT DIAMETER
  WALL THICKNESS
  GRADE
  PIPE SPEC
  SEAM TYPE

VALIDATION LOG
──────────────
Each validation is logged to validation_log.json in the CSV
folder. Records: timestamp, user, CSV file, FXL file, MTR file.
"""
    text.insert("1.0", help_content)
    text.config(state="disabled")
    tk.Button(win, text="Close", command=win.destroy).pack(pady=6)
```

- [ ] **Step 3: Manual test** — Click "? Help". Window should open with scrollable text. Close button should work.

---

### Task 10: Column Mapping UI (Schema Override)

**Files:**
- Modify: `DataValidationTool-v3.1.py` — `_load_csv_path`, add `_ask_column_mapping`

- [ ] **Step 1: Add `_ask_column_mapping` method**

```python
def _ask_column_mapping(self, df: pd.DataFrame) -> tuple[bool, dict, list[int]] | None:
    """Show a dialog to manually map CSV columns to known roles.
    Returns (has_station, mapping, attr_indices) or None if cancelled."""
    ncols = df.shape[1]
    headers = [str(df.iloc[0, c]) if not df.empty else f"Col {c}" for c in range(ncols)]

    win = tk.Toplevel(self)
    win.title("Column Mapping")
    win.grab_set()
    win.resizable(False, False)

    tk.Label(win, text="Auto-detection failed. Map CSV columns to known fields:",
             wraplength=420, justify="left").pack(padx=12, pady=(10, 4))

    roles = ["Point Number", "Northing", "Easting", "Elevation", "Field Code", "Station (optional)"]
    choices: dict[str, tk.StringVar] = {}
    opts = ["(none)"] + [f"[{i}] {h[:40]}" for i, h in enumerate(headers)]

    frame = tk.Frame(win)
    frame.pack(padx=12, pady=4)
    for r, role in enumerate(roles):
        tk.Label(frame, text=role, width=22, anchor="e").grid(row=r, column=0, padx=4, pady=2)
        var = tk.StringVar(value="(none)")
        choices[role] = var
        om = tk.OptionMenu(frame, var, *opts)
        om.config(width=34)
        om.grid(row=r, column=1, padx=4, pady=2)

    result: list = [None]

    def on_ok():
        def idx_of(role: str) -> int | None:
            v = choices[role].get()
            if v == "(none)":
                return None
            try:
                return int(v.split("]")[0].lstrip("["))
            except Exception:
                return None

        pn = idx_of("Point Number")
        north = idx_of("Northing")
        east = idx_of("Easting")
        elev = idx_of("Elevation")
        fc = idx_of("Field Code")
        station = idx_of("Station (optional)")

        if any(v is None for v in [pn, north, east, elev, fc]):
            messagebox.showwarning("Incomplete", "Please map at least: Point Number, Northing, Easting, Elevation, Field Code.", parent=win)
            return

        used = {pn, north, east, elev, fc}
        if station is not None:
            used.add(station)
        attrs = [c for c in range(ncols) if c not in used]
        mapping = {"station": station, "pn": pn, "north": north, "east": east, "elev": elev, "fc": fc}
        result[0] = (station is not None, mapping, attrs)
        win.destroy()

    def on_cancel():
        win.destroy()

    btn_frame = tk.Frame(win)
    btn_frame.pack(pady=8)
    tk.Button(btn_frame, text="OK", width=10, command=on_ok).grid(row=0, column=0, padx=6)
    tk.Button(btn_frame, text="Cancel", width=10, command=on_cancel).grid(row=0, column=1, padx=6)
    win.wait_window()
    return result[0]
```

- [ ] **Step 2: In `_load_csv_path`, when `_detect_schema` raises, offer the mapping dialog instead of silently setting None**

Replace:
```python
try:
    self.has_station, self.mapping, self.attr_indices = self._detect_schema(self.df)
except Exception:
    self.has_station, self.mapping, self.attr_indices = None, None, []
```

With:
```python
try:
    self.has_station, self.mapping, self.attr_indices = self._detect_schema(self.df)
except Exception as schema_err:
    self.status.config(text=f"Auto-detection failed: {schema_err}. Manual mapping required.")
    mapped = self._ask_column_mapping(self.df)
    if mapped is not None:
        self.has_station, self.mapping, self.attr_indices = mapped
    else:
        self.has_station, self.mapping, self.attr_indices = None, None, []
```

- [ ] **Step 3: Manual test** — Use a CSV where column order is non-standard. When schema detection fails, the mapping dialog should appear. After mapping, validation should proceed normally.

---

## Chunk 5: Advanced Features

### Task 11: Numeric Bounds Validation

**Files:**
- Modify: `DataValidationTool-v3.1.py` — `_load_app_config`, `_export_and_open_excel` (BOUNDS sheet), VBA string (new `CheckBoundsForRow` sub + call in `ValidateRow`)

- [ ] **Step 1: Update `config.json` default with a `numeric_bounds` section**

Add these defaults to `config.json` (the file will be updated by the app at runtime):
```json
{
  "single_excel_instance": true,
  "numeric_bounds": {
    "ANGLE": [0, 360],
    "HORIZONTAL ANGLE": [0, 360],
    "VERTICAL ANGLE": [-90, 90],
    "BEND ANGLE": [0, 180],
    "JOINT LENGTH": [0, 80],
    "WALL THICKNESS": [0, 5],
    "DEPTH": [-30, 200]
  }
}
```

- [ ] **Step 2: In `_load_app_config`, read `numeric_bounds` into `self.numeric_bounds`**

```python
raw_bounds: Any = cfg.get("numeric_bounds", {})
if isinstance(raw_bounds, dict):
    self.numeric_bounds: dict[str, list[float]] = {
        str(k).strip().upper(): v
        for k, v in raw_bounds.items()
        if isinstance(v, list) and len(v) == 2
    }
else:
    self.numeric_bounds = {}
```

Also add `self.numeric_bounds: dict[str, list[float]] = {}` to `__init__` state block.

- [ ] **Step 3: In `_export_and_open_excel`, write a BOUNDS hidden sheet after writing FXL sheet**

```python
if self.numeric_bounds:
    ws_bounds = wb.create_sheet("BOUNDS")
    ws_bounds.append(["AttrName", "Min", "Max"])
    for attr_name, (mn, mx) in self.numeric_bounds.items():
        ws_bounds.append([attr_name, mn, mx])
    ws_bounds.sheet_state = "veryHidden"
```

- [ ] **Step 4: In the VBA string, add `ReadBoundsCache` helper after `GetFXLIndex`**

Find the `GetFXLIndex` function end and insert:
```vba
' ---- Numeric bounds cache (from BOUNDS sheet) ----
Private Function GetBoundsCache() As Object
    Static cache As Object
    Static loaded As Boolean
    If Not loaded Then
        loaded = True
        Set cache = CreateObject("Scripting.Dictionary")
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("BOUNDS")
        If Not ws Is Nothing Then
            Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            Dim r As Long
            For r = 2 To lr
                Dim nm As String: nm = Trim$(UCase$(ws.Cells(r, 1).Value))
                Dim mn As Double: mn = CDbl(ws.Cells(r, 2).Value)
                Dim mx As Double: mx = CDbl(ws.Cells(r, 3).Value)
                If nm <> "" And Err.Number = 0 Then
                    cache(nm) = Array(mn, mx)
                End If
                Err.Clear
            Next r
        End If
        On Error GoTo 0
    End If
    Set GetBoundsCache = cache
End Function

Private Sub CheckBoundsForRow(ByVal r As Long)
    Dim bc As Object: Set bc = GetBoundsCache()
    If bc Is Nothing Or bc.Count = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = Worksheets("Data")
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        Dim hdr As String: hdr = Trim$(UCase$(ws.Cells(1, c).Value))
        If bc.Exists(hdr) Then
            Dim cellVal As String: cellVal = Trim$(ws.Cells(r, c).Value)
            If cellVal <> "" Then
                Dim numVal As Double
                On Error Resume Next
                numVal = CDbl(cellVal)
                If Err.Number = 0 Then
                    Dim bounds As Variant: bounds = bc(hdr)
                    If numVal < bounds(0) Or numVal > bounds(1) Then
                        FlagOrange r, c, "Out of range (" & bounds(0) & "–" & bounds(1) & ")"
                    End If
                End If
                On Error GoTo 0
            End If
        End If
    Next c
End Sub
```

- [ ] **Step 5: In VBA `ValidateRow`, at the very end before `End Sub`, call `CheckBoundsForRow r`**

- [ ] **Step 6: Manual test** — Set a JOINT LENGTH value of "999" in a validated CSV. That cell should be flagged orange with "Out of range (0–80)".

---

### Task 12: Error Type Filter Button

**Files:**
- Modify: `DataValidationTool-v3.1.py` — `_build_ui`, VBA string

- [ ] **Step 1: Add "Filter: Errors Only" toggle button in `btns2`**

```python
self.btn_filter_errors = tk.Button(
    btns2, text="Show Errors Only", width=18,
    state="disabled", command=self.toggle_error_filter
)
self.btn_filter_errors.grid(row=0, column=2, padx=8)
self._error_filter_on = False
ToolTip(self.btn_filter_errors, "Hide rows with no validation errors to focus on problem rows.")
```

- [ ] **Step 2: Enable `btn_filter_errors` alongside the other buttons in `_export_and_open_excel`**

Add `self.btn_filter_errors.config(state="normal")` where other buttons are enabled (~line 3500).

Also add `self.btn_filter_errors.config(state="disabled")` where other buttons are disabled (on close).

- [ ] **Step 3: Add VBA `FilterErrorRows` and `ShowAllRows` subs after `NormalizeView`**

```vba
Public Sub FilterErrorRows()
    ' Hide all rows that have no colored cells (no errors)
    Dim ws As Worksheet: Set ws = Worksheets("Data")
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lc As Long: lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Application.ScreenUpdating = False
    Dim r As Long
    For r = 2 To lr
        Dim hasErr As Boolean: hasErr = False
        Dim c As Long
        For c = 1 To lc
            If ws.Cells(r, c).Interior.ColorIndex <> xlNone Then
                hasErr = True
                Exit For
            End If
        Next c
        ws.Rows(r).Hidden = Not hasErr
    Next r
    Application.ScreenUpdating = True
End Sub

Public Sub ShowAllRows()
    Dim ws As Worksheet: Set ws = Worksheets("Data")
    ws.Rows.Hidden = False
End Sub
```

- [ ] **Step 4: Add `toggle_error_filter` Python method**

```python
def toggle_error_filter(self) -> None:
    if not (self._excel and self._wb_com):
        return
    self._error_filter_on = not self._error_filter_on
    wbname = self._wb_com.Name
    macro = "FilterErrorRows" if self._error_filter_on else "ShowAllRows"
    for target in [f"'{wbname}'!ValidationModule.{macro}", f"ValidationModule.{macro}", macro]:
        try:
            self._excel.Run(target)
            break
        except Exception:
            pass
    label = "Show All Rows" if self._error_filter_on else "Show Errors Only"
    self.btn_filter_errors.config(text=label)
```

- [ ] **Step 5: Manual test** — After validation, click "Show Errors Only". Rows without any colored cells should disappear. Click again ("Show All Rows") to restore.

---

### Task 13: VBA Project Password Protection

**Files:**
- Modify: `DataValidationTool-v3.1.py` — VBA string (self-protection macro), `_export_and_open_excel` (call after injection)

Note: Setting VBA project passwords via COM is unreliable across Excel versions. The safest approach is a VBA macro that protects the project from within.

- [ ] **Step 1: Add `ProtectVBAProject` sub to VBA string**

```vba
Public Sub ProtectVBAProject()
    ' Protect the VBA project to prevent casual editing.
    ' The password is intentionally simple — this is a deterrent, not security.
    On Error Resume Next
    ThisWorkbook.VBProject.Protection.Password = "dvt31"
    ThisWorkbook.VBProject.Protection.EnforceProject = True
    On Error GoTo 0
End Sub
```

- [ ] **Step 2: In `_export_and_open_excel`, after running `ValidateSheetAll`, also run `ProtectVBAProject`**

```python
for target in [
    f"'{wb_com.Name}'!ValidationModule.ProtectVBAProject",
    "ValidationModule.ProtectVBAProject",
]:
    try:
        excel.Run(target)
        break
    except Exception:
        pass
```

- [ ] **Step 3: Manual test** — After validation opens in Excel, go to Developer → Visual Basic. The VBA project should be marked as protected.

---

## Execution Notes

- All changes are to the single file `DataValidationTool-v3.1.py` (plus minor `config.json` edit).
- VBA changes are string edits inside the `vbcode = r"""..."""` block (~line 1285 to ~3460).
- No test framework exists; all tests are manual run steps.
- Tasks 1–8 are safe/additive. Tasks 9–13 involve more invasive changes; test each individually.
- Commit after each Task completes.
