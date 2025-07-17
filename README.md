# Manufacturing CLI Tool

A simple command-line tool to assist with common manufacturing data tasks such as parsing test files, downloading images, and converting file formats.

---
### Run via Terminal
```bash
python cli_tool.py
```

### Or

- Double-click on **`cli_tool.py`**
- Choose a script from the menu
- Drag & drop or paste the required file path
- Done!

---

## Scripts Menu

<img src="web%20images/menu_d.png" alt="CLI Menu" width="300"/>

---

## Available Scripts

### 1️⃣ Download Images from URL

- Downloads images listed in an Excel file (with URLs)
- Output folder: `downloaded_images/`

### 2️⃣ Analyze JSON/ZIP

- Parses PoE EELoad or Adapter Edac test data  
- Accepts `.json`, `.txt`, or zipped multiple files (`.zip`)
- Output folder: `extracted/`

### 3️⃣ Convert CSV to Excel

- Converts PoE EELoad `.csv` files into Excel format with custom headers
- Output folder: `extracted/`

### 4️⃣ Split CSV Tests

- Splits `.csv` files containing over 5000 tests  
- Useful for uploading data to Factory Web
- Output folder: `extracted/`  

### 5️⃣ File Parser

- Parses raw PoE Network or Adapter Edac test files from `.zip` archives
- Output folder: `extracted/`

---

## Requirements

- Windows
- Python 3.7+
- Packages (auto-installed)

## Internal Use Only

This tool is intended for internal factory and engineering team workflows only.
