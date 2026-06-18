# Python Scripts

A collection of Python scripts for data processing, automation, and academic projects.

Each script is independent and documented internally within the code.

---

## 📂 Scripts

### 🌦️ descarga_estaciones_smn.py

Downloads monthly climatological normals from CONAGUA's SMN for a list of weather station IDs provided in a CSV file. Parses the raw `.txt` response files and generates one formatted Excel workbook per station, with one sheet per climate variable.

- Reads station IDs from a CSV
- Downloads data from the SMN CONAGUA server
- Parses variable blocks and statistical rows (min, max, mean, std)
- Generates styled Excel files (green headers, aligned columns, fixed widths)
- Includes retry logic for network timeouts

**Requirements:** `requests`, `openpyxl`

---

### 🌦️ descarga_estaciones_smn_sin_formato.py

Same logic as `descarga_estaciones_smn.py` but outputs plain Excel files without any styling (no colors, no alignment, no column widths). Lighter and faster alternative when formatting is not needed.

**Requirements:** `requests`, `openpyxl`

---

### 🌿 Michoacan_agricultural_filter.py

Filters and aggregates agricultural production data from SIAP open datasets for Michoacán (temporal modality only). Processes all CSV files in a folder, normalizes column names across dataset versions (2015–2020 format differences), aggregates by crop/municipality/cycle, computes yield (`Rendimiento = Volumen / Cosechada`), and exports a clean CSV per input file.

**Requirements:** `pandas`

---

### 🧾 word_tables_to_excel.py

Converts tables from Word documents (`.docx`) into Excel files (`.xlsx`), preserving structure, merges, and basic formatting.

- Handles merged cells (horizontal and vertical)
- Preserves text formatting and colors
- Automatically detects header rows
- Filters out irrelevant tables (text-only layouts)
- Generates one Excel file per Word document
- Supports batch processing from an input folder to an output folder

**Requirements:** `python-docx`, `openpyxl`

---

### 👤 registro_INFOTEC.py

OOP-based terminal registration form that collects and validates user input (name, age, email). Implements a retry counter (max 4 attempts per field) with specific error messages per field and graceful exit on interruption or too many failed attempts. Designed as a class to support unit testing without modifying core logic.

**Requirements:** `sys` (standard library)

---

### 🧪 test_registro_INFOTEC.py

Unit tests for `registro_INFOTEC.py` using `unittest` and `unittest.mock`. Simulates incorrect inputs to verify error messages, retry countdown, and final accepted values for each field (age, name, email).

**Requirements:** `unittest`, `unittest.mock` (standard library)

---

## 📌 Notes

- Scripts are fully independent from each other
- Each file includes detailed inline documentation
- `registro_INFOTEC.py` and `test_registro_INFOTEC.py` must be in the same directory to run the tests
