# Skoda DSR Generator - User Guide

The **Skoda DSR Generator** is a standalone desktop application designed to automate the extraction of shipment data from PDFs and synchronize it with Shakti (Zoho Creator) while maintaining a master Daily Status Report (DSR).

---

## 🛠 Features

### 1. Invoice, MBL, HBL Extractor (Main Workflow)
Extracts container-wise data from multiple PDF sources simultaneously.
*   **Supported BL Formats:** Maersk, Hapag-Lloyd, Evergreen.
*   **Automatic Mapping:** Links Invoice numbers to specific containers by scanning PDF content.
*   **Global Field Sync:** Automatically applies User, Month, and ETA to all parsed records.
*   **Shakti Integration:** Pushes cleaned data directly to the Shakti (Zoho) server.

### 2. Shakti Export to DSR Converter
Converts bulk Excel exports from Shakti into user-specific DSR files.
*   **Automatic Splitting:** Generates separate files for **Ashish (CSN)**, **Ranjit (PUNE)**, and **CLC**.
*   **Status Sorting:** Splits records into **'Live shipments'** and **'Cleared shipments'** based on Dispatch Date.

---

## 🚀 How to Use

### Phase 1: Data Extraction
1.  **Select Files:** Click **Select Invoice**, **Select MBL**, and **Select HBL** to load your PDFs. (You can select multiple invoices/HBLs).
2.  **Verify Settings:** 
    *   Select the correct **User** and **Month**.
    *   Pick the **Pre-alert Receive Date** and **Vessel ETA** using the calendar 📅 buttons.
3.  **Extract:** Click the green **Extract** button. The data will appear in the queue below.
4.  **Review:** Click **1. Review & Confirm Data**. 
    *   **Double-click** any cell to edit details manually.
    *   Use **+ Add Row** to insert manual data or **✖ Remove Row** to delete duplicates.
5.  **Submit:** Click **2. Push to Shakti & Export**. This will:
    *   Upload data to the Shakti server.
    *   Append the records to your local `SKODA_MASTER_DSR.xlsx`.

### Phase 2: Generating User-wise DSRs
1.  Navigate to the **Shakti export file to DSR's** tab.
2.  Click **Browse Shakti Excel** and select your export file.
3.  Click **CONVERT & GENERATE DSR's**.
4.  Check your application folder for the new `.xlsx` files named by User and Date.

---

## ⚠️ Important Notes
*   **Master File:** Ensure `SKODA_MASTER_DSR.xlsx` is in the same folder as the EXE. If deleted, the tool will create a new one.
*   **Duplicate Check:** The tool automatically warns you if it finds a container number and invoice combo already present in the Master DSR.
*   **Internet Connection:** An active internet connection is required to push records to Shakti.

---
© Nagarkot Forwarders Pvt Ltd
