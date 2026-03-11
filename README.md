# Skoda DSR Generator

A specialized Nagarkot tool for extracting container-wise shipment data from Bill of Lading (BL) PDFs and Invoices for Skoda, Audi, and Volkswagen shipments.

## Tech Stack
- **Python 3.10+**
- **Tkinter** (GUI)
- **PyMuPDF** (PDF Parsing)
- **openpyxl** (Excel Generation)
- **Pillow** (Logo Support)

---

## Features
- **Multi-Format Parsing**: Supports Maersk (Non-Negotiable Waybill) and Hapag-Lloyd (Sea Waybill) formats.
- **Container-Wise Extraction**: Automatically splits multiple containers within a single BL into separate DSR rows.
- **Invoice Recognition**: Detects invoice numbers from filenames or within the BL text.
- **Master DSR Integration**: Append data directly to an existing Master DSR file or create a new one.
- **Nagarkot UI Standards**: Clean, branded interface with date selectors and live previews.

---

## Installation

### 1. Clone or Copy
Download the project files to your local directory.

### 2. Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment to ensure dependency isolation.

**1. Create virtual environment:**
```bash
python -m venv venv
```

**2. Activate (REQUIRED):**

Windows:
```powershell
venv\Scripts\activate
```

Mac/Linux:
```bash
source venv/bin/activate
```

**3. Install dependencies:**
```bash
pip install -r requirements.txt
```

**4. Run application:**
```bash
python skoda_dsr_generator.py
```

---

## Usage Guide

1.  **Select PDFs**: Choose BL PDFs and Invoice PDFs (can select multiple folders).
2.  **Manual Fields**: Enter the User Name, Pre-alert Date, and Vessel ETA.
3.  **Master DSR (Optional)**: If you have a master file, browse and select it to append data.
4.  **Generate**: Click "Generate DSR" to produce the Excel output.

---

## Notes
- **Always use virtual environment** for development and execution.
- Ensure `Nagarkot Logo.png` is in the same directory as the script for proper branding.
- Do not commit `venv/` or sensitive Excel files to version control.
