"""
Skoda DSR Generator
Extracts shipment data from BL PDFs and Invoices, generating DSR Excel (container-wise).

Features:
- Maersk & Hapag-Lloyd BL formats.
- Invoices extracted from PDF filenames.
- tkcalendar for date selection.
- GUI layout matches Nagarkot VW/Audi tool standards.
"""

from __future__ import annotations

import re
import logging
import os
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

try:
    from tkcalendar import Calendar
except ImportError:
    # This will still show red if the IDE doesn't see the package, 
    # but the runtime check is here.
    Calendar = None 

import fitz  # PyMuPDF
import openpyxl
import openpyxl.utils
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image, ImageTk

from bl_parser import ContainerRecord, parse_bl
from zoho_api import ZohoCreatorAPI

# ─── Constants ───────────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).resolve().parent
LOGO_PATH = SCRIPT_DIR / "Nagarkot Logo.png"
MASTER_FILE_PATH = SCRIPT_DIR / "SKODA_MASTER_DSR.xlsx"

# DSR column headers in order (A → BF = 58 columns)
DSR_HEADERS: list[str] = [
    "User",                           # A
    "Pre-alert Receive date",         # B
    "Month",                          # C
    "FF/ Shipping Line",              # D
    "Port of Loading",                # E
    "Vessel Name",                    # F
    "BL Date",                        # G
    "Vessel ETA",                     # H
    "CHA Job No.",                    # I
    "Container No.",                  # J
    "Size (20'40' LCL)",              # K
    "Container Type (HQ,DV,SD)",      # L
    "Current Status",                 # M
    "BL No.",                         # N
    "Supplier Name",                  # O
    "Invoice No.",                    # P
    "INCO",                           # Q
    "No.of Pkg.",                     # R
    "GrossWt",                        # S
    "CFS Name",                       # T
    "IGM No.",                        # U
    "IGM No. Date",                   # V
    "IGM Inward Date",                # W
    "B/E No",                         # X
    "B/E Date",                       # Y
    "AO Ass",                         # Z
    "AC Assess",                      # AA
    "RMS/ Examine",                   # AB
    "Duty Request recd from CHA",     # AC
    "Duty Paid date",                 # AD
    "Assessable Value",               # AE
    "Debit Duty (RODTEP)",            # AF
    "Total Duty",                     # AG
    "DUTY%",                          # AH
    "STAMP DUTY",                     # AI
    "Interest (IfAny)",               # AJ
    "Penalty (Ifany)",                # AK
    "Reason for Interest/Penalty",    # AL
    "OOC Date",                       # AM
    "Dispatch date to plant/WH",      # AN
    "Remarks (Daywise Cronology)",    # AO
    "Clearnace TAT",                  # AP
    "Reason for Clearance TAT delay", # AQ
    "E-Waybill No.",                  # AR
    "Detention/Demurrage (IfAny)",    # AS
    "Total BCD Value",                # AT
    "Total SWS Value",                # AU
    "Total IGST Value",               # AV
    "CHA JOB NO",                     # AW
    "Transporter",                    # AX
    "STAMP DUTY PAID DT",             # AY
    "Under Protest",                  # AZ
    "BOE filing TAT",                 # BA
    "Reason for BOE filing TAT delay",# BB
    "Conatainer arrival date in CFS", # BC
    "OOC COPY RECD YES/NO",           # BD
    "Remarks",                        # BE
    "SIMS Registration date",         # BF
]

MONTH_MAP = {
    "01": "JAN", "02": "FEB", "03": "MAR", "04": "APR",
    "05": "MAY", "06": "JUN", "07": "JUL", "08": "AUG",
    "09": "SEP", "10": "OCT", "11": "NOV", "12": "DEC",
}

# ─── Theme ───────────────────────────────────────────────────────────────────

NAGARKOT_BLUE = "#1B3A5C"
BTN_BLUE = "#0056b3"
WHITE = "#FFFFFF"
BG_LIGHT = "#F8F9FA"

class BetterDateEntry(tk.Frame):
    """Custom DatePicker without the Windows scaling double-calendar bug in DateEntry."""
    def __init__(self, master, width=15, bg="white", **kwargs):
        super().__init__(master, bg=bg)
        self.entry_var = tk.StringVar()
        self.entry = ttk.Entry(self, textvariable=self.entry_var, width=width, font=("Segoe UI", 9))
        self.entry.pack(side="left")
        
        self.btn = tk.Button(self, text="📅", command=self._popup, relief="flat", bg="white", cursor="hand2")
        self.btn.pack(side="left", padx=2)
        
    def _popup(self):
        top = tk.Toplevel(self)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        top.geometry(f"+{x}+{y}")
        
        frame = tk.Frame(top, highlightbackground="#0056b3", highlightthickness=2)
        frame.pack()
        
        if not Calendar:
            messagebox.showerror("Missing Dependency", "Please run: pip install tkcalendar")
            top.destroy()
            return

        cal = Calendar(frame, selectmode="day", date_pattern="yyyy-mm-dd", showweeknumbers=False, 
                       font=("Segoe UI", 9), background="#0056b3", foreground="white",
                       headersbackground="#F8F9FA", headersforeground="black")
        cal.pack()
        
        def on_select(e):
            self.entry_var.set(cal.get_date())
            top.destroy()
            
        cal.bind("<<CalendarSelected>>", on_select)
        top.bind("<FocusOut>", lambda e: top.destroy())
        cal.focus_set()

    def get(self):
        return self.entry_var.get()
        
    def delete(self, first, last=None):
        self.entry.delete(first, last)

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(500, self.showtip)

    def unschedule(self):
        id_ = self.id
        self.id = None
        if id_:
            self.widget.after_cancel(id_)

    def showtip(self, event=None):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                      background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                      font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

class DataPreviewWindow(tk.Toplevel):
    def __init__(self, parent, records: list[ContainerRecord], on_confirm):
        super().__init__(parent)
        self.title("Review & Edit Extracted Data")
        self.geometry("1400x650")
        self.configure(bg=WHITE)
        self.records = records
        self.on_confirm = on_confirm
        
        self.transient(parent)
        self.grab_set()
        
        header = tk.Frame(self, bg=WHITE, height=60)
        header.pack(fill="x", side="top")
        tk.Label(header, text="Review & Edit Data", font=("Segoe UI", 16, "bold"), bg=WHITE, fg=BTN_BLUE).pack(pady=5)
        tk.Label(header, text="Double-click any cell to edit details per row.", font=("Segoe UI", 10), bg=WHITE, fg="#6c757d").pack(pady=2)

        frame = tk.Frame(self, bg=WHITE)
        frame.pack(fill="both", expand=True, padx=20, pady=5)
        
        self.col_map = [
            ("user", "User", 90),
            ("user_month", "Month", 60),
            ("pre_alert_date", "Pre-Alert", 90),
            ("vessel_eta", "Vessel ETA", 90),
            ("bl_type", "BL Type", 80),
            ("bl_mode", "Mode", 80),
            ("bl_no", "BL No", 100),
            ("container_no", "Container No", 100),
            ("invoice_nos", "Invoice Nos", 180),
            ("supplier_name", "Supplier", 180),
            ("inco_terms", "INCO", 70),
            ("num_packages", "Packages", 80),
            ("gross_weight", "Gross Wt", 80),
            ("container_size", "Size", 50),
            ("container_type", "Type", 50),
            ("vessel_name", "Vessel Name", 120),
            ("port_of_loading", "POL", 100),
            ("shipping_line", "Line", 100),
            ("bl_date", "BL Date", 90)
        ]
        
        cols = [c[1] for c in self.col_map]
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)
        
        for attr, heading, w in self.col_map:
            self.tree.heading(heading, text=heading)
            self.tree.column(heading, width=w, stretch=False, anchor="w")
            
        for rec in self.records:
            vals = [getattr(rec, attr) for attr, _, _ in self.col_map]
            self.tree.insert("", "end", values=vals)
            
        self.tree.bind("<Double-1>", self._on_double_click)
        
        footer = tk.Frame(self, bg=WHITE, height=50)
        footer.pack(fill="x", side="bottom", padx=20, pady=10)
        
        tk.Button(
            footer, text="Cancel", font=("Segoe UI", 10), width=15, 
            command=self.destroy
        ).pack(side="left")
        
        tk.Button(
            footer, text="Confirm & Submit", font=("Segoe UI", 10, "bold"),
            bg=BTN_BLUE, fg=WHITE, width=20, cursor="hand2",
            command=self._do_confirm
        ).pack(side="right")
        
    def _on_double_click(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
            
        x, y, width, height = self.tree.bbox(row_id, col_id)
        col_idx = int(col_id[1:]) - 1
        
        value = self.tree.item(row_id, "values")[col_idx]
        
        entry = ttk.Entry(self.tree, font=("Segoe UI", 9))
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, value)
        entry.focus()
        
        def commit(e=None):
            new_val = entry.get()
            vals = list(self.tree.item(row_id, "values"))
            vals[col_idx] = new_val
            self.tree.item(row_id, values=vals)
            
            rec_idx = self.tree.index(row_id)
            attr_name = self.col_map[col_idx][0]
            setattr(self.records[rec_idx], attr_name, new_val)
            entry.destroy()
            
        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        
    def _do_confirm(self):
        self.destroy()
        self.on_confirm()

class DSRGeneratorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Skoda DSR Generator — Nagarkot")
        self.root.state("zoomed")
        self.root.configure(bg=WHITE)
        self.root.minsize(1000, 700)

        # Style configurations
        style = ttk.Style()
        style.theme_use("clam")
        
        # LabelFrame Style
        style.configure("TLabelFrame", background=WHITE)
        style.configure("TLabelFrame.Label", font=("Segoe UI", 10, "bold"), foreground=BTN_BLUE, background=WHITE)
        
        # Treeview Style
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#E9ECEF", foreground="#495057")
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=25)
        
        # State
        self.files_by_dir: dict[Path, list[Path]] = {}
        self.parsed_records: list[ContainerRecord] = []
        self.confirmed_records: list[ContainerRecord] = []
        self.master_dsr_path: Path = MASTER_FILE_PATH
        self._load_logo()
        self._build_header()
        self._build_file_selection()
        self._build_manual_settings()
        self._build_footer()
        self._build_preview()

    def _load_logo(self) -> None:
        self.logo_img = None
        if LOGO_PATH.exists():
            try:
                img = Image.open(LOGO_PATH)
                img = img.resize((150, 22), Image.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
            except Exception:
                pass

    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=WHITE, height=80)
        header.pack(fill="x", side="top", pady=10)
        
        # Logo on left
        if self.logo_img:
            lbl_logo = tk.Label(header, image=self.logo_img, bg=WHITE)
            lbl_logo.pack(side="left", padx=20)
        
        # Titles strictly centered
        title_frame = tk.Frame(header, bg=WHITE)
        title_frame.pack(side="top", expand=True)

        tk.Label(
            title_frame, text="Skoda DSR Generator",
            font=("Segoe UI", 20, "bold"), fg=BTN_BLUE, bg=WHITE
        ).pack()

        tk.Label(
            title_frame, text="Extract container-wise data from Bills of Lading and Invoices",
            font=("Segoe UI", 11), fg="#6c757d", bg=WHITE
        ).pack(pady=(2, 0))

    def _build_file_selection(self) -> None:
        frame = ttk.LabelFrame(self.root, text="File Selection", padding=(10, 8))
        frame.pack(fill="x", padx=20, pady=5)
        
        btn_frame = tk.Frame(frame, bg=WHITE)
        btn_frame.pack(fill="x", side="left")

        # Standard buttons
        btn_sel_pdfs = tk.Button(btn_frame, text="Select PDFs", width=15, command=self._on_select_pdfs)
        btn_sel_pdfs.pack(side="left", padx=(0, 10))
        
        btn_sel_folder = tk.Button(btn_frame, text="Select Folder", width=15, command=self._on_select_folder)
        btn_sel_folder.pack(side="left", padx=(0, 10))
        
        btn_clear = tk.Button(btn_frame, text="Clear List", width=15, command=self._on_clear_list)
        btn_clear.pack(side="left", padx=(0, 20))

        self.lbl_file_status = tk.Label(btn_frame, text="No files selected", font=("Segoe UI", 9), fg="#495057", bg=WHITE)
        self.lbl_file_status.pack(side="left")

    def _build_manual_settings(self) -> None:
        frame = ttk.LabelFrame(self.root, text="Manual / Zoho Fields", padding=(10, 8))
        frame.pack(fill="x", padx=20, pady=5)
        
        inner = tk.Frame(frame, bg=WHITE)
        inner.pack(fill="x")
        
        # User
        tk.Label(inner, text="User:", font=("Segoe UI", 9, "bold"), bg=WHITE).grid(row=0, column=0, sticky="w", padx=5)
        self.var_user = tk.StringVar(value="Ashish (CSN)")
        self.cb_user = ttk.Combobox(inner, textvariable=self.var_user, values=["Ashish (CSN)", "Ranjit (PUNE)"], width=27)
        self.cb_user.grid(row=0, column=1, sticky="w", padx=5)
        
        # Pre-alert Date (custom pop-up)
        tk.Label(inner, text="Pre-alert Receive Date:", font=("Segoe UI", 9, "bold"), bg=WHITE).grid(row=0, column=2, sticky="w", padx=(20, 5))
        self.cal_pre_alert = BetterDateEntry(inner, width=15, bg=WHITE)
        self.cal_pre_alert.grid(row=0, column=3, sticky="w", padx=5)
        self.cal_pre_alert.delete(0, "end")
        
        # Vessel ETA (custom pop-up)
        tk.Label(inner, text="Vessel ETA:", font=("Segoe UI", 9, "bold"), bg=WHITE).grid(row=0, column=4, sticky="w", padx=(20, 5))
        self.cal_vessel_eta = BetterDateEntry(inner, width=15, bg=WHITE)
        self.cal_vessel_eta.grid(row=0, column=5, sticky="w", padx=5)
        self.cal_vessel_eta.delete(0, "end")
        
        # Month Dropdown
        tk.Label(inner, text="Month:", font=("Segoe UI", 9, "bold"), bg=WHITE).grid(row=0, column=6, sticky="w", padx=(20, 5))
        self.var_month = tk.StringVar()
        self.cb_month = ttk.Combobox(inner, textvariable=self.var_month, values=list(MONTH_MAP.values()), width=6, state="readonly")
        self.cb_month.grid(row=0, column=7, sticky="w", padx=5)
        
        # BL Type Radio buttons
        tk.Label(inner, text="BL Type:", font=("Segoe UI", 9, "bold"), bg=WHITE).grid(row=0, column=8, sticky="w", padx=(20, 5))
        self.var_bl_type = tk.StringVar(value="MAWB_MBL")
        rb_frame = tk.Frame(inner, bg=WHITE)
        rb_frame.grid(row=0, column=9, sticky="w")
        tk.Radiobutton(rb_frame, text="MBL", variable=self.var_bl_type, value="MAWB_MBL", bg=WHITE, font=("Segoe UI", 9)).pack(side="left")
        tk.Radiobutton(rb_frame, text="HBL", variable=self.var_bl_type, value="HAWB_HBL", bg=WHITE, font=("Segoe UI", 9)).pack(side="left")
        
        # Mode Dropdown
        tk.Label(inner, text="Mode:", font=("Segoe UI", 9, "bold"), bg=WHITE).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.var_mode = tk.StringVar(value="Sea (FCL)")
        self.cb_mode = ttk.Combobox(inner, textvariable=self.var_mode, values=["Air", "Sea (FCL)", "Sea (LCL)", "Sea (BB)"], width=15, state="readonly")
        self.cb_mode.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        # Branch (Locked to MUMBAI)
        tk.Label(inner, text="Branch:", font=("Segoe UI", 9, "bold"), bg=WHITE).grid(row=1, column=2, sticky="w", padx=(20, 5), pady=5)
        self.var_branch = tk.StringVar(value="MUMBAI")
        self.entry_branch = ttk.Entry(inner, textvariable=self.var_branch, width=15, state="readonly")
        self.entry_branch.grid(row=1, column=3, sticky="w", padx=5, pady=5)


    def _build_preview(self) -> None:
        frame = ttk.LabelFrame(self.root, text="Data Preview / Processing Queue", padding=(10, 8))
        frame.pack(fill="both", expand=True, padx=20, pady=5)

        columns = ("Directory", "Files", "Status", "Parsed Container(s)", "Invoice Nos", "BL No")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings")
        
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)

        col_widths = {
            "Directory": 200, "Files": 100, "Status": 100, 
            "Parsed Container(s)": 150, "Invoice Nos": 200, "BL No": 120
        }
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths.get(col, 100), anchor="w")

    def _build_footer(self) -> None:
        footer = tk.Frame(self.root, bg=WHITE, height=60)
        footer.pack(fill="x", side="bottom", padx=20, pady=10)
        
        tk.Label(footer, text="© Nagarkot Forwarders Pvt Ltd", font=("Segoe UI", 8), fg="#6c757d", bg=WHITE).pack(side="left")
        
        btn_frame = tk.Frame(footer, bg=WHITE)
        btn_frame.pack(side="right")

        self.btn_review = tk.Button(
            btn_frame, text="1. Review & Confirm Data", font=("Segoe UI", 10, "bold"),
            bg="#f39c12", fg=WHITE, activebackground="#e67e22", activeforeground=WHITE,
            width=25, height=2, borderwidth=0, cursor="hand2",
            command=self._on_review
        )
        self.btn_review.pack(side="left", padx=10)

        self.btn_push = tk.Button(
            btn_frame, text="2. Push to Zoho & Export", font=("Segoe UI", 10, "bold"),
            bg=BTN_BLUE, fg=WHITE, activebackground="#004494", activeforeground=WHITE,
            width=25, height=2, borderwidth=0, cursor="arrow",
            command=self._on_push_and_export, state="disabled"
        )
        self.btn_push.pack(side="left")

    # ═══ Actions ═════════════════════════════════════════════════════════

    def _on_select_pdfs(self) -> None:
        files = filedialog.askopenfilenames(
            title="Select BL & Invoice PDFs",
            filetypes=[("PDF Files", "*.pdf")],
        )
        if files:
            for f in files:
                p = Path(f)
                d = p.parent
                if d not in self.files_by_dir:
                    self.files_by_dir[d] = []
                if p not in self.files_by_dir[d]:
                    self.files_by_dir[d].append(p)
            self._parse_and_refresh()

    def _on_select_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder containing shipment PDFs")
        if folder:
            p = Path(folder)
            pdfs = list(p.glob("*.pdf"))
            if pdfs:
                if p not in self.files_by_dir:
                    self.files_by_dir[p] = []
                for f in pdfs:
                    if f not in self.files_by_dir[p]:
                        self.files_by_dir[p].append(f)
                self._parse_and_refresh()
            else:
                messagebox.showinfo("No PDFs", "No PDF files found in the selected folder.")

    def _on_clear_list(self) -> None:
        self.files_by_dir.clear()
        self.parsed_records.clear()
        self.confirmed_records.clear()
        self.tree.delete(*self.tree.get_children())
        self.lbl_file_status.config(text="No files selected")
        self.btn_push.config(state="disabled", cursor="arrow")


    def _parse_and_refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        self.parsed_records.clear()
        
        total_files = sum(len(flist) for flist in self.files_by_dir.values())
        self.lbl_file_status.config(text=f"{total_files} file(s) across {len(self.files_by_dir)} folder(s)")

        for directory, files in self.files_by_dir.items():
            # Identify BL vs Invoice
            bl_files = []
            invoice_pdf_files = []
            
            for f in files:
                stem_upper = f.stem.upper()
                if (stem_upper.startswith("MAEU") or stem_upper.startswith("HLCU") 
                    or stem_upper == "BL" or stem_upper.startswith("MEAU")
                    or stem_upper.startswith("SWB")):
                    bl_files.append(f)
                else:
                    invoice_pdf_files.append(f)
            
            # Parse BLs
            dir_records = []
            for bl_file in bl_files:
                try:
                    recs = parse_bl(bl_file)
                    dir_records.extend(recs)
                except Exception as e:
                    logger.error(f"Error parsing {bl_file.name}: {e}")
                    status = f"Error parsing BL"
            
            container_to_invoices = {rec.container_no.upper(): set() for rec in dir_records}
            container_to_supplier = {}
            unmapped_invoices = set()
            unmapped_supplier = None  # Supplier detected from invoices that have no container numbers
            all_invoice_stems = []
            
            for inv_f in invoice_pdf_files:
                matches = re.findall(r"\b([45]\d{7})\b", inv_f.stem)
                # Also try 8-digit numbers starting with 7 or 8 (Skoda AS format)
                if not matches:
                    matches = re.findall(r"\b(\d{8})\b", inv_f.stem)
                inv_no = matches[0] if matches else inv_f.stem
                all_invoice_stems.append(inv_no)
                
                try:
                    doc = fitz.open(str(inv_f))
                    text_all = "".join(page.get_text().upper() for page in doc)
                    doc.close()
                    
                    found_containers = set(re.findall(r"\b([A-Z]{4}\d{7})\b", text_all))
                    
                    # Detect Supplier from Invoice Content
                    detected_supplier = None
                    if "AUDI HUNGARIA" in text_all:
                        detected_supplier = "AUDI HUNGARIA ZRT."
                    elif "VOLKSWAGEN AG" in text_all:
                        detected_supplier = "VOLKSWAGEN AG"
                    elif "AUDI AG" in text_all:
                        detected_supplier = "AUDI AG"
                    elif "SKODA AUTO" in text_all or "CELKOV" in text_all or "IN SEA" in inv_f.stem.upper():
                        detected_supplier = "Skoda Auto A.S."

                    mapped = False
                    for c_no in found_containers:
                        if c_no in container_to_invoices:
                            container_to_invoices[c_no].add(inv_no)
                            if detected_supplier:
                                container_to_supplier[c_no] = detected_supplier
                            mapped = True
                    
                    if not mapped:
                        # Invoice has no matching container (e.g. Skoda AS invoices)
                        unmapped_invoices.add(inv_no)
                        if detected_supplier:
                            unmapped_supplier = detected_supplier
                        
                except Exception as e:
                    logger.error(f"Error parsing invoice {inv_f.name}: {e}")
                    unmapped_invoices.add(inv_no)

            all_invoice_stems = list(dict.fromkeys(all_invoice_stems))
            invoices_str = "/".join(all_invoice_stems)

            status = "Parsed"
            containers_str = "-"
            bl_no = "-"

            if not bl_files:
                status = "Error: No BL found"
                self.tree.insert("", "end", values=(directory.name, f"{len(files)} files", status, containers_str, invoices_str, bl_no))
                continue
                
            # Apply invoice mappings to the parsed base records
            if invoice_pdf_files:
                for rec in dir_records:
                    cno = rec.container_no.upper()
                    mapped_set = container_to_invoices.get(cno, set())
                    
                    # Update Supplier from container-specific invoice mapping first
                    if cno in container_to_supplier:
                        rec.supplier_name = container_to_supplier[cno]
                    elif unmapped_supplier:
                        # Fallback: Use supplier from unmapped invoices (e.g. Skoda AS with no containers)
                        rec.supplier_name = unmapped_supplier
                    
                    if mapped_set:
                        # Exclusively use invoices mapped specifically to this container
                        all_for_container = sorted(list(mapped_set))
                        # Also append unmapped invoices to each container
                        if unmapped_invoices:
                            all_for_container = sorted(list(mapped_set | unmapped_invoices))
                        rec.invoice_nos = "/".join(all_for_container)
                    else:
                        # No container-specific mapping: assign all unmapped or all invoices
                        if unmapped_invoices:
                            rec.invoice_nos = "/".join(sorted(list(unmapped_invoices)))
                        else:
                            rec.invoice_nos = invoices_str

            self.parsed_records.extend(dir_records)

            if dir_records and not self.var_month.get():
                m_date = dir_records[0].bl_date
                if m_date and len(m_date) >= 7:
                    auto_month = MONTH_MAP.get(m_date[5:7], "")
                    if auto_month:
                        self.var_month.set(auto_month)

            if dir_records:
                containers_str = ", ".join(r.container_no for r in dir_records)
                bl_nos = list(dict.fromkeys(r.bl_no for r in dir_records))
                bl_no = ", ".join(bl_nos)
            
            self.tree.insert("", "end", values=(directory.name, f"{len(files)} PDFs", status, containers_str, invoices_str, bl_no))

    def _on_review(self) -> None:
        if not self.parsed_records:
            messagebox.showwarning("No Data", "No valid parsed container records to review.")
            return
            
        # Before spawning preview, explicitly assign global manual settings to ALL the records
        global_user = self.var_user.get().strip()
        global_month = self.var_month.get().strip()
        global_pre = self.cal_pre_alert.get()
        global_eta = self.cal_vessel_eta.get()
        global_bl_type = self.var_bl_type.get()
        global_bl_mode = self.var_mode.get()
        
        for r in self.parsed_records:
            r.user = global_user
            r.user_month = global_month
            r.pre_alert_date = global_pre
            r.vessel_eta = global_eta
            r.bl_type = global_bl_type
            r.bl_mode = global_bl_mode

        # Pop up the Data Review modal. 
        DataPreviewWindow(self.root, self.parsed_records, self._on_confirmation_complete)

    def _on_confirmation_complete(self) -> None:
        """Called after user finishes editing/confirming in the Review window."""
        self.confirmed_records = list(self.parsed_records)
        self.btn_push.config(state="normal", cursor="hand2")
        messagebox.showinfo("Ready", "Data confirmed! You can now click '2. Push to Zoho & Export'.")

    def _get_existing_invoices(self) -> set[tuple[str, str]]:
        """Reads the local Master DSR and returns a set of (Container No, Invoice No) tuples."""
        existing_keys = set()
        if not MASTER_FILE_PATH.exists():
            return existing_keys
            
        try:
            wb = openpyxl.load_workbook(MASTER_FILE_PATH, read_only=True)
            # Find common sheet names
            sheet_name = next((n for n in ["Live shipments", "DSR"] if n in wb.sheetnames), wb.sheetnames[0])
            ws = wb[sheet_name]
            
            # DSR_HEADERS: J=9 (Container No), P=15 (Invoice No) (0-indexed)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if len(row) > 15:
                    c_no = str(row[9]).strip().upper() if row[9] else ""
                    inv_no_str = str(row[15]).strip().upper() if row[15] else ""
                    if c_no and inv_no_str:
                        # Split multiple invoices to check individually
                        for single_inv in inv_no_str.split("/"):
                            existing_keys.add((c_no, single_inv.strip()))
            wb.close()
        except Exception as e:
            logger.warning(f"Could not read master file for duplicate check: {e}")
        return existing_keys

    def _on_push_and_export(self) -> None:
        if not self.confirmed_records:
            messagebox.showwarning("Error", "Please review and confirm data first.")
            return

        # 1. Duplicate Check using local Master file
        existing_keys = self._get_existing_invoices()
        records_to_process = []
        duplicate_count = 0
        
        for rec in self.confirmed_records:
            c_no = (rec.container_no or "").strip().upper()
            inv_str = (rec.invoice_nos or "").strip().upper()
            
            # Check if this container+invoice combo already exists
            is_new = False
            for single_inv in inv_str.split("/"):
                inv_key = single_inv.strip()
                if not inv_key or (c_no, inv_key) not in existing_keys:
                    is_new = True
                    break
            
            if not is_new and inv_str: # If ALL invoices for this container are already there
                duplicate_count += 1
            else:
                records_to_process.append(rec)

        if not records_to_process:
            messagebox.showinfo("Duplicates Found", f"All {duplicate_count} records were already found in the Master DSR.\nNo new data to push.")
            return

        if duplicate_count > 0:
            if not messagebox.askyesno("Duplicates Found", f"{duplicate_count} records appear to be already in the Master DSR.\n\nContinue pushing {len(records_to_process)} records?"):
                return

        # 2. Push to Zoho
        try:
            zoho = ZohoCreatorAPI()
            success, msg = zoho.push_records(records_to_process)
            
            if not success:
                messagebox.showerror("Zoho Push Failed", f"Excel export aborted because Zoho push failed:\n\n{msg}")
                return
                
            zoho_msg = f"Zoho API: {msg}"
            
        except Exception as ze:
            messagebox.showerror("Zoho API Error", f"Excel export aborted due to API error:\n\n{ze}")
            return

        # 3. If Zoho is successful, Update Master Excel
        try:
            if MASTER_FILE_PATH.exists():
                self._append_to_master(MASTER_FILE_PATH, records_to_process)
                excel_msg = "Data successfully appended to Master DSR!"
            else:
                self._create_new_dsr(MASTER_FILE_PATH, records_to_process)
                excel_msg = "Master DSR created and data saved successfully!"
            
            messagebox.showinfo("Success", f"{excel_msg}\n\n{zoho_msg}")
            
            # Reset workflow
            self.confirmed_records.clear()
            self.btn_push.config(state="disabled", cursor="arrow")

        except Exception as exc:
            logger.exception("Failed to update Master Excel")
            messagebox.showerror("Excel Error", f"An error occurred after Zoho push while updating the Master file:\n{exc}")

    # ── Excel Export Logic ───────────────────────────────────────────────

    def _record_to_row(self, rec: ContainerRecord) -> list:
        row = [""] * len(DSR_HEADERS)
        
        def _parse_dt(dt_str):
            if not dt_str: return None
            try: return datetime.strptime(dt_str, "%Y-%m-%d")
            except: 
                try: return datetime.strptime(dt_str, "%d-%b-%Y")
                except: return dt_str

        row[0] = rec.user
        row[1] = _parse_dt(rec.pre_alert_date)
        row[2] = rec.user_month
        row[3] = rec.shipping_line
        row[4] = rec.port_of_loading
        row[5] = rec.vessel_name

        bl_dt = None
        try: bl_dt = datetime.strptime(rec.bl_date, "%Y-%m-%d")
        except: pass
        row[6] = bl_dt if bl_dt else rec.bl_date

        row[7] = _parse_dt(rec.vessel_eta)
        row[9] = rec.container_no
        row[10] = rec.container_size
        row[11] = rec.container_type
        row[13] = rec.bl_no
        row[14] = rec.supplier_name
        row[15] = rec.invoice_nos
        row[16] = rec.inco_terms

        try: row[17] = int(rec.num_packages)
        except: row[17] = rec.num_packages

        try: row[18] = float(rec.gross_weight)
        except: row[18] = rec.gross_weight

        return row

    def _apply_dsr_styling(self, ws: Worksheet) -> None:
        """Applies NAGARKOT styling to the DSR sheet."""
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="1B3A5C")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        data_font = Font(name="Segoe UI", size=9)
        data_align = Alignment(vertical="center", wrap_text=False)

        # Style header
        for col_idx, _ in enumerate(DSR_HEADERS, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        ws.row_dimensions[1].height = 40
        ws.freeze_panes = "A2"

        # Style data rows
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = data_font
                cell.alignment = data_align
                cell.border = thin_border
                # Date formats for columns B, G, H (1-indexed: 2, 7, 8)
                if cell.column in (2, 7, 8):
                    cell.number_format = "YYYY-MM-DD"

        # Auto-width
        for col_idx in range(1, len(DSR_HEADERS) + 1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 15

    def _append_to_master(self, path: Path, records: list[ContainerRecord]) -> None:
        wb = openpyxl.load_workbook(path)
        sheet_name = next((n for n in ["Live shipments", "DSR"] if n in wb.sheetnames), wb.sheetnames[0])
        ws = wb[sheet_name]
        
        for rec in records:
            ws.append(self._record_to_row(rec))
        
        self._apply_dsr_styling(ws)
        wb.save(path)
        wb.close()

    def _create_new_dsr(self, save_path: Path, records: list[ContainerRecord]) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Live shipments"
        ws.append(DSR_HEADERS)
        
        for rec in records:
            ws.append(self._record_to_row(rec))
            
        self._apply_dsr_styling(ws)
        # Create second sheet empty as per original code pattern
        cleared = wb.create_sheet("Cleared shipments")
        # Apply header to second sheet too
        cleared.append(DSR_HEADERS)
        self._apply_dsr_styling(cleared)

        wb.save(save_path)
        wb.close()


def main() -> None:
    root = tk.Tk()
    app = DSRGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
