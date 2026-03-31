"""
BL (Bill of Lading) Parser — Maersk & Hapag-Lloyd formats.

Extracts container-wise shipment data from BL PDFs for DSR generation.
Each BL may contain multiple containers; returns one record per container.
"""

from __future__ import annotations

import re
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF

logger = logging.getLogger(__name__)


@dataclass
class ContainerRecord:
    """One DSR row worth of data extracted from a BL."""

    shipping_line: str = ""          # Col D
    port_of_loading: str = ""        # Col E
    vessel_name: str = ""            # Col F
    bl_date: str = ""                # Col G
    container_no: str = ""           # Col J
    container_size: str = ""         # Col K  (20/40)
    container_type: str = ""         # Col L  (HQ/DV/SD)
    bl_no: str = ""                  # Col N
    supplier_name: str = ""          # Col O
    invoice_nos: str = ""            # Col P
    inco_terms: str = ""             # Col Q
    num_packages: str = ""           # Col R
    gross_weight: str = ""           # Col S
    port_of_discharge: str = ""      # Extra context
    consignee: str = ""              # Extra context
    hs_codes: list[str] = field(default_factory=list)
    user: str = ""                   # Global manual input
    user_month: str = ""             # Global manual input
    pre_alert_date: str = ""         # Global manual input
    vessel_eta: str = ""             # Global manual input
    bl_type: str = ""                # Global manual input
    bl_mode: str = ""                # Global manual input
    hbl_no: str = ""                 # New field for combined scenario
    mbl_no: str = ""                 # New field for combined scenario
    raw_mbl_no: str = ""             # Unsliced original MBL number


def parse_bl(pdf_path: str | Path) -> list[ContainerRecord]:
    """
    Auto-detect BL format (Maersk / Hapag-Lloyd) and parse.

    Returns a list of ContainerRecord — one per container found in the BL.
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"BL file not found: {pdf_path}")

    doc = fitz.open(str(pdf_path))
    all_pages_text: list[str] = []
    for page in doc:
        all_pages_text.append(page.get_text())
    doc.close()
    full_text = "\n".join(all_pages_text)

    # Auto-detect format
    text_upper = full_text.upper()
    if "HAPAG" in text_upper or "HLCU" in text_upper:
        logger.info("Detected Hapag-Lloyd BL format")
        return _parse_hapag(full_text, pdf_path.stem)
    elif "MAERSK" in text_upper or "MAEU" in text_upper:
        logger.info("Detected Maersk BL format")
        return _parse_maersk(full_text, pdf_path.stem)
    elif "EVERGREEN" in text_upper or "EGLV" in text_upper:
        logger.info("Detected Evergreen BL format")
        return _parse_evergreen(full_text, pdf_path.stem)
    else:
        logger.warning("Unknown BL format, attempting Maersk parser as fallback")
        return _parse_maersk(full_text, pdf_path.stem)


# ─── Maersk Parser ──────────────────────────────────────────────────────────

def _parse_maersk(text: str, filename: str) -> list[ContainerRecord]:
    """Parse Maersk Non-Negotiable Waybill."""

    base = ContainerRecord(shipping_line="MAERSK")

    # BL Number — from filename or text
    # Look for common patterns: DESTRA..., EGLV..., HLCU..., MAEU...
    bl_patterns = [
        r"(?:B/L|WAYBILL|SWB)\s+(?:NOS?|NO\.?|NUMBER)?[:\s]*([A-Z]{4,6}[0-9]{6,})", # Prefixed (DESTRA, HLCU)
        r"(?:B/L|WAYBILL|SWB)\s+(?:NOS?|NO\.?|NUMBER)?[:\s]*([0-9]{9,12})",     # Numeric (Maersk)
        r"\b(DESTRA[0-9]{8,12})\b",                                             # Specific DESTRA
        r"B/L[:\s]*([A-Z0-9]{8,})",                                              # Generic Captive
    ]
    
    ignore_words = {"ATTACHMENT", "NON-NEGOTIABLE", "ORIGINAL", "COPY", "DRAFT"}
    
    best_candidate = ""
    for p in bl_patterns:
        matches = re.finditer(p, text, re.IGNORECASE)
        for m in matches:
            val = m.group(1).upper().strip()
            if val in ignore_words or len(val) < 8:
                continue
            # If it's a specific prefix we want (like DESTRA), take it immediately
            if val.startswith("DESTRA"):
                best_candidate = val
                break
            if not best_candidate:
                best_candidate = val
        if best_candidate and best_candidate.startswith("DESTRA"):
            break
            
    if best_candidate:
        base.bl_no = best_candidate
            
    if not base.bl_no:
        # Fallback to filename
        if re.match(r"(?:MAEU)?(\d{9})", filename.upper()):
            match = re.match(r"(?:MAEU)?(\d{9})", filename.upper())
            base.bl_no = match.group(1) if match else filename
        else:
            # Try to find any alphanumeric ID in filename
            fid_match = re.search(r"([A-Z0-9]{8,})", filename.upper())
            base.bl_no = fid_match.group(1) if fid_match else filename

    if base.bl_no.isdigit() and len(base.bl_no) >= 9:
        base.raw_mbl_no = f"MAEU{base.bl_no}"
    else:
        base.raw_mbl_no = base.bl_no


    # Vessel
    vessel_match = re.search(r"Vessel\s*\n\s*(.+?)(?:\n|Voyage)", text, re.DOTALL)
    if vessel_match:
        base.vessel_name = vessel_match.group(1).strip()

    # Port of Loading
    pol_match = re.search(r"Port of Loading\s*\n\s*(.+)", text)
    if pol_match:
        base.port_of_loading = pol_match.group(1).strip()

    # Port of Discharge
    pod_match = re.search(r"Port of Discharge\s*\n\s*(.+)", text)
    if pod_match:
        base.port_of_discharge = pod_match.group(1).strip()

    # Shipped on Board Date
    date_match = re.search(r"Shipped on Board.*?\n\s*(\d{4}-\d{2}-\d{2})", text)
    if date_match:
        base.bl_date = date_match.group(1)

    # Shipper (Supplier) — look for known shipper names
    if re.search(r"VOLKSWAGEN\s+KONZERNLOGISTIK", text, re.IGNORECASE):
        base.supplier_name = "VOLKSWAGEN KONZERNLOGISTIK GMBH & CO.OHG AS AGENT OF SKODA AUTO A.S."
    elif re.search(r"SKODA\s+AUTO\s+A\.?S\.?", text, re.IGNORECASE):
        base.supplier_name = "SKODA AUTO A.S."

    # Consignee
    consignee_match = re.search(
        r"Consignee.*?\)\s*\n(.+?)(?:Port of Discharge)",
        text, re.DOTALL
    )
    if consignee_match:
        lines = [l.strip() for l in consignee_match.group(1).strip().split("\n") if l.strip()]
        base.consignee = lines[0] if lines else ""

    # Freight terms → INCO hint
    if "FREIGHT COLLECT" in text.upper():
        base.inco_terms = "FCA"  # Collect typically = FCA
    elif "FREIGHT PREPAID" in text.upper():
        base.inco_terms = "CIF"

    # ── Extract container blocks ─────────────────────────────────────────
    # Pattern: Container line like "TCLU9480892  40 DRY 9'6  52 PACKAGES  16484.633 KGS"
    # Also matches PALLETS (used in Skoda AS / After Sales BLs)
    # And handles optional seal numbers between container number and size
    container_pattern = re.compile(
        r"([A-Z]{4}\d{7})\s+"           # Container No
        r"(?:[\w-]+\s+)?"                # Optional Seal/Extra Info (e.g. ML-DE2359372)
        r"(\d{2})\s+"                    # Size (20/40)
        r"(DRY|HIGH\s*CUBE|REEFER)\s+"   # Type keyword
        r"(?:\d+'\d+(?:''|\"|')?\s+)?"   # Optional height (e.g. 9'6, 8'6, 9'6")
        r"(\d+)\s+(?:PACKAGES?|PALLETS?)\s+"  # Package/Pallet count
        r"([\d.,]+)\s+KGS",             # Weight
        re.IGNORECASE
    )

    # Find all invoice numbers mentioned in the BL
    invoice_matches = re.findall(r"INVOICE\s+NOS?\.?:?\s*\n?\s*(\d{5,})", text, re.IGNORECASE)
    all_invoices = "/".join(dict.fromkeys(invoice_matches))

    # Maersk BLs often have a duplicate COPY section (pages 4-6 repeat pages 1-3).
    # We will deduplicate exact container lines.
    seen_container_lines = set()

    # Find all HS codes
    hs_matches = re.findall(r"(?:HS\s+Codes?:?\s*\n?)((?:\s*\d{6}\s*\n?)+)", text, re.IGNORECASE)
    all_hs = []
    for block in hs_matches:
        codes = re.findall(r"(\d{6})", block)
        all_hs.extend(codes)
    all_hs = list(dict.fromkeys(all_hs))

    # Find unique containers with their details
    seen_containers: dict[str, ContainerRecord] = {}

    for match in container_pattern.finditer(text):
        full_line = match.group(0).strip()
        if full_line in seen_container_lines:
            continue
        seen_container_lines.add(full_line)

        cno = match.group(1)
        size = match.group(2)
        type_raw = match.group(3).upper()
        pkgs = match.group(4)
        weight = match.group(5).replace(",", "")

        # Determine container type
        if "HIGH" in type_raw or "CUBE" in type_raw:
            ctype = "HQ"
        elif "REEFER" in type_raw:
            ctype = "RF"
        else:
            ctype = "SD"  # Standard container (changed from DV) # Standard/Dry container

        if cno not in seen_containers:
            rec = ContainerRecord(
                shipping_line=base.shipping_line,
                port_of_loading=base.port_of_loading,
                vessel_name=base.vessel_name,
                bl_date=base.bl_date,
                container_no=cno,
                container_size=size,
                container_type=ctype,
                bl_no=base.bl_no,
                raw_mbl_no=base.raw_mbl_no,
                supplier_name=base.supplier_name,
                inco_terms=base.inco_terms,
                num_packages="0",
                gross_weight="0",
                port_of_discharge=base.port_of_discharge,
                consignee=base.consignee,
                invoice_nos=all_invoices,
                hs_codes=all_hs,
            )
            seen_containers[cno] = rec

        # Accumulate packages and weight across repeated container lines
        rec = seen_containers[cno]
        rec.num_packages = str(int(rec.num_packages) + int(pkgs))
        current_wt = float(rec.gross_weight)
        rec.gross_weight = str(round(current_wt + float(weight), 3))

    records = list(seen_containers.values())

    # If no containers found via regex, create a single record with what we have
    if not records:
        logger.warning("No container details parsed via regex, creating single record")
        base.invoice_nos = all_invoices
        base.hs_codes = all_hs
        records = [base]

    return records


# ─── Hapag-Lloyd Parser ─────────────────────────────────────────────────────

def _parse_hapag(text: str, filename: str) -> list[ContainerRecord]:
    """Parse Hapag-Lloyd Sea Waybill."""

    base = ContainerRecord(shipping_line="HAPAG-LLOYD")

    # SWB (Sea Waybill) Number
    swb_match = re.search(r"SWB[- ]No\.?\s*[:\s]*(HLCU\w+)", text)
    if swb_match:
        base.bl_no = swb_match.group(1)
    elif filename.upper().startswith("HLCU"):
        base.bl_no = filename

    base.raw_mbl_no = base.bl_no


    # Vessel and Voyage
    vessel_match = re.search(
        r"(?:SAVANNAH|SALERNO|SANTOS|SOFIA|STOCKHOLM|SINGAPORE|STRAIT|SAJIR|"
        r"[A-Z][A-Z\s]+EXPRESS|[A-Z][A-Z\s]+HIGHWAY|"
        r"[A-Z]{3,}(?:\s+[A-Z]+)?)\s+(\d{2,4}[A-Z]?)",
        text
    )
    if vessel_match:
        full = vessel_match.group(0)
        # Split vessel name from voyage number
        parts = full.rsplit(None, 1)
        base.vessel_name = parts[0].strip() if len(parts) > 1 else full.strip()

    # Better vessel extraction from SHIPPED ON BOARD line
    sob_match = re.search(r"VESSEL\s+NAME:\s*(.+?)\s*VOYAGE", text, re.IGNORECASE)
    if sob_match:
        base.vessel_name = sob_match.group(1).strip()

    # Port of Loading from SHIPPED ON BOARD block
    pol_matches = re.findall(r"PORT OF LOADING:\s*([A-Z][^\n\r]+)", text, re.IGNORECASE)
    for p in pol_matches:
        pol_val = p.strip()
        if pol_val.upper() not in ("PORT", "OF", "PORT OF DISCHARGE:"):
            base.port_of_loading = pol_val.title()
            break
            
    if not base.port_of_loading:
        # Fallback: look for known port names anywhere on an indented line
        for port_name in ["HAMBURG", "BREMERHAVEN"]:
            if re.search(rf"^\s+{port_name}", text, re.MULTILINE | re.IGNORECASE):
                base.port_of_loading = port_name.title()
                break

    # Shipped on Board Date
    date_match = re.search(
        r"SHIPPED ON BOARD.*?DATE\s*:\s*(\d{1,2})[.\-](\w{3})[.\-](\d{4})",
        text, re.IGNORECASE
    )
    if date_match:
        day = date_match.group(1).zfill(2)
        month_str = date_match.group(2).upper()[:3]
        year = date_match.group(3)
        months = {
            "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
            "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
            "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12",
        }
        mm = months.get(month_str, "01")
        base.bl_date = f"{year}-{mm}-{day}"

    # Port of Discharge
    pod_match = re.search(r"NHAVA SHEVA|JAWAHARLAL NEHRU|MUNDRA|CHENNAI", text, re.IGNORECASE)
    if pod_match:
        base.port_of_discharge = pod_match.group(0).title()

    # Shipper
    if re.search(r"VOLKSWAGEN\s+KONZERNLOGISTIK", text, re.IGNORECASE):
        base.supplier_name = "VOLKSWAGEN KONZERNLOGISTIK GMBH & CO.OHG AS AGENT OF SKODA AUTO A.S."
    elif re.search(r"SKODA\s+AUTO\s+A\.?S\.?", text, re.IGNORECASE):
        base.supplier_name = "SKODA AUTO A.S."

    # Consignee
    consignee_match = re.search(r"SKODA AUTO VOLKSWAGEN INDIA\s+PRIVATE", text, re.IGNORECASE)
    if consignee_match:
        base.consignee = "SKODA AUTO VOLKSWAGEN INDIA PVT LTD"

    # Freight / INCO
    if "FREIGHT COLLECT" in text.upper():
        base.inco_terms = "FCA"
    elif "FREIGHT PREPAID" in text.upper():
        base.inco_terms = "CIF"

    # Check for CIF in the text directly
    cif_match = re.search(r"(?:Terms of delivery|Lieferbedingungen)\s*\n?\s*(CIF|FCA|FOB|CFR)", text, re.IGNORECASE)
    if cif_match:
        base.inco_terms = cif_match.group(1).upper()

    # Invoice numbers — Hapag has "INVOICE NOS.:" and then 8-digit numbers starting with 4 or 5.
    inv_idx = text.upper().find("INVOICE NOS")
    if inv_idx > 0:
        # Search the text after INVOICE NOS for any 8-digit numbers starting with 4 or 5
        search_area = text[inv_idx:inv_idx+500]
        all_inv: list[str] = re.findall(r"\b([45]\d{7})\b", search_area)
    else:
        all_inv: list[str] = []
    
    all_invoices = "/".join(dict.fromkeys(all_inv))

    # HS codes
    hs_matches = re.findall(r"(?:HS\s+Codes?:?\s*\n?)((?:\s*\d{6}\s*\n?)+)", text, re.IGNORECASE)
    all_hs = []
    for block in hs_matches:
        codes = re.findall(r"(\d{6})", block)
        all_hs.extend(codes)
    all_hs = list(dict.fromkeys(all_hs))

    # ── Container extraction ─────────────────────────────────────────────
    # Hapag format: "UACU  5689190     29 PACKAGES    11542.425   67.850"
    # Some BLs omit the measurement column, so it must be optional.
    container_pattern = re.compile(
        r"([A-Z]{4})\s+(\d{7})\s+"          # Container prefix + number (sometimes spaced)
        r"(\d+)\s+PACKAGES?\s+"             # Packages
        r"([\d.]+)"                         # Weight (measurement is optional after this)
        r"(?:\s+([\d.]+))?",                # Measurement (optional)
        re.IGNORECASE
    )

    # Also try: "40'X9'6" HIGH CUBE" pattern for size/type
    size_match = re.search(r"(\d{2})'[Xx](?:8'6\"|9'6\")\s*(HIGH\s*CUBE|STANDARD)?", text)
    default_size = "40"
    default_type = "SD" # Standard dryer container (changed from DV)
    if size_match:
        default_size = size_match.group(1)
        if size_match.group(2) and "HIGH" in size_match.group(2).upper():
            default_type = "HQ"

    seen_containers: dict[str, ContainerRecord] = {}

    for match in container_pattern.finditer(text):
        prefix = match.group(1)
        number = match.group(2)
        cno = f"{prefix}{number}"
        pkgs = match.group(3)
        weight = match.group(4)

        if cno not in seen_containers:
            rec = ContainerRecord(
                shipping_line=base.shipping_line,
                port_of_loading=base.port_of_loading,
                vessel_name=base.vessel_name,
                bl_date=base.bl_date,
                container_no=cno,
                container_size=default_size,
                container_type=default_type,
                bl_no=base.bl_no,
                raw_mbl_no=base.raw_mbl_no,
                supplier_name=base.supplier_name,
                inco_terms=base.inco_terms,
                num_packages=pkgs,
                gross_weight=weight,
                port_of_discharge=base.port_of_discharge,
                consignee=base.consignee,
                invoice_nos=all_invoices,
                hs_codes=all_hs,
            )
            seen_containers[cno] = rec

    records = list(seen_containers.values())

    # When a single container has multiple cargo blocks (e.g. DG + non-DG split),
    # the BL provides a summary line like "=====\n 44 PACKAGES   6079.495   69.755".
    # Use that summary for total packages and weight if we found exactly 1 container.
    if len(records) == 1:
        summary_match = re.search(
            r"=+\s*\n\s*(\d+)\s+PACKAGES?\s+([\d.]+)",
            text, re.IGNORECASE
        )
        if summary_match:
            records[0].num_packages = summary_match.group(1)
            records[0].gross_weight = summary_match.group(2)

    if not records:
        logger.warning("No containers parsed from Hapag BL, creating single record")
        base.invoice_nos = all_invoices
        base.hs_codes = all_hs
        records = [base]

    return records


# ─── Evergreen Parser ───────────────────────────────────────────────────────

def _parse_evergreen(text: str, filename: str) -> list[ContainerRecord]:
    """Parse Evergreen Bill of Lading."""
    base = ContainerRecord(shipping_line="EVERGREEN")

    # BL Number
    bl_match = re.search(r"B/L\s+NO\.\s*(EGLV\d+)", text, re.IGNORECASE)
    if bl_match:
        base.bl_no = bl_match.group(1).upper()
    elif filename.upper().startswith("EGLV"):
        base.bl_no = filename.upper()
    
    # Try looking for EGLV followed by numbers anywhere
    if not base.bl_no:
        bl_alt = re.search(r"(EGLV\d{5,})", text, re.IGNORECASE)
        if bl_alt:
            base.bl_no = bl_alt.group(1).upper()
            
    base.raw_mbl_no = base.bl_no


    # Vessel Name
    # Evergreen often has "M.V. EVER ..." or "EVER ..."
    vessel_match = re.search(r"(EVER\s+[A-Z]+)\s+(\d{3,4}[A-Z]?)", text)
    if vessel_match:
        base.vessel_name = f"{vessel_match.group(1)} {vessel_match.group(2)}".strip()
    else:
        # Try finding "M.V. " line
        mv_match = re.search(r"M\.V\.\s+([A-Z\s]+)", text)
        if mv_match:
            base.vessel_name = mv_match.group(1).strip()

    # Port of Loading
    pol_match = re.search(r"PORT OF LOADING\s*\n\s*(.+)", text, re.IGNORECASE)
    if not pol_match:
         # Often appears twice, once in a header and once in shipment info
         pol_match = re.search(r"PENANG|NHAVA SHEVA|PORT KELANG", text, re.IGNORECASE)
    
    if pol_match:
        # If it was a group(1) match, use it, else use the raw string
        pol_val = pol_match.group(1).strip() if len(pol_match.groups()) > 0 else pol_match.group(0).strip()
        base.port_of_loading = pol_val.title()

    # Date
    # Evergreen date format: MAR.13,2026
    date_match = re.search(r"([A-Z]{3})\.?(\d{1,2}),\s*(\d{4})", text, re.IGNORECASE)
    if date_match:
        mon_str = date_match.group(1).upper()[:3]
        day = date_match.group(2).zfill(2)
        year = date_match.group(3)
        months = {
            "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
            "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
            "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12",
        }
        mm = months.get(mon_str, "01")
        base.bl_date = f"{year}-{mm}-{day}"

    # Shipper / Supplier
    if "PREMIUM SOUND SOLUTIONS" in text.upper():
        base.supplier_name = "PREMIUM SOUND SOLUTIONS SDN BHD"
    elif "VOLKSWAGEN" in text.upper():
        base.supplier_name = "VOLKSWAGEN AG"
    elif "DR.ING.H.C.F.PORSCHE" in text.upper():
        base.supplier_name = "DR.ING.H.C.F.PORSCHE AG"

    # Container Parsing
    # Pattern: EGSU1926250/40H/EMCSGN8724/42 PALLETS
    container_pattern = re.compile(
        r"([A-Z]{4}\d{7})\s*/\s*"         # Container No
        r"(\d{2})([A-Z]?)",               # Size (20/40) + Type (H/D etc)
        re.IGNORECASE
    )

    seen_containers: dict[str, ContainerRecord] = {}

    for match in container_pattern.finditer(text):
        cno = match.group(1).upper()
        size = match.group(2)
        type_code = match.group(3).upper()
        
        ctype = "HQ" if type_code == "H" else "SD"
        
        if cno not in seen_containers:
            rec = ContainerRecord(
                shipping_line="EVERGREEN",
                bl_no=base.bl_no,
                vessel_name=base.vessel_name,
                port_of_loading=base.port_of_loading,
                bl_date=base.bl_date,
                supplier_name=base.supplier_name,
                raw_mbl_no=base.raw_mbl_no,
                container_no=cno,
                container_size=size,
                container_type=ctype,
            )
            
            # Find weight/packages for this container
            # Evergreen layout is tricky; weight typically appears once for the whole BL if single container
            weight_match = re.search(r"([\d,.]+)\s*KGS", text)
            if weight_match:
                rec.gross_weight = weight_match.group(1).replace(",", "")
            
            pkg_match = re.search(r"(\d+)\s+(?:PALLETS|PACKAGES|BOXES)", text, re.IGNORECASE)
            if pkg_match:
                rec.num_packages = pkg_match.group(1)
                
            seen_containers[cno] = rec

    records = list(seen_containers.values())
    if not records:
        records = [base]
    
    return records

