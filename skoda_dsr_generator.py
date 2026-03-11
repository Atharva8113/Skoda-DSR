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
    messagebox.showerror("Missing Dependency", "Please run: pip install tkcalendar")
    sys.exit(1)

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image, ImageTk

from bl_parser import ContainerRecord, parse_bl
from zoho_api import ZohoCreatorAPI

# ─── Constants ───────────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).resolve().parent
LOGO_PATH = SCRIPT_DIR / "Nagarkot Logo.png"

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
        self.master_dsr_path: Optional[Path] = None

        self._load_logo()
        self._build_header()
        self._build_file_selection()
        self._build_manual_settings()
        self._build_output_settings()
        self._build_preview()
        self._build_footer()

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
        self.var_user = tk.StringVar(value="Ashish(CSN)")
        self.cb_user = ttk.Combobox(inner, textvariable=self.var_user, values=["Ashish(CSN)", "Ranjit(PUNE)"], width=27)
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

    def _build_output_settings(self) -> None:
        frame = ttk.LabelFrame(self.root, text="Output Settings", padding=(10, 8))
        frame.pack(fill="x", padx=20, pady=5)
        
        inner = tk.Frame(frame, bg=WHITE)
        inner.pack(fill="x")
        
        tk.Label(inner, text="Master DSR (optional):", font=("Segoe UI", 9), bg=WHITE).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        
        self.var_master_dsr = tk.StringVar()
        entry_master = tk.Entry(inner, textvariable=self.var_master_dsr, width=100)
        entry_master.grid(row=0, column=1, padx=5, pady=2)
        entry_master.config(state="readonly")
        
        tk.Button(inner, text="Browse...", width=10, command=self._on_browse_master).grid(row=0, column=2, padx=5, pady=2)
        tk.Button(inner, text="Clear", width=8, command=self._on_clear_master).grid(row=0, column=3, padx=5, pady=2)
        
        tk.Label(inner, text="Output Folder:", font=("Segoe UI", 9), bg=WHITE).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.var_out_folder = tk.StringVar(value="")
        tk.Entry(inner, textvariable=self.var_out_folder, width=100).grid(row=1, column=1, padx=5, pady=2)
        tk.Button(inner, text="Browse...", width=10, command=self._on_browse_out_folder).grid(row=1, column=2, padx=5, pady=2)
        
        tk.Label(inner, text="Output Filename:", font=("Segoe UI", 9), bg=WHITE).grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.var_out_filename = tk.StringVar(value=f"SKODA_DSR_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        tk.Entry(inner, textvariable=self.var_out_filename, width=100).grid(row=2, column=1, padx=5, pady=2)
        tk.Label(inner, text="(.xlsx added automatically)", font=("Segoe UI", 8), fg="#6c757d", bg=WHITE).grid(row=2, column=2, sticky="w", padx=5, pady=2)
        
        self.var_push_zoho = tk.BooleanVar(value=True)
        tk.Checkbutton(inner, text="Push Data to Zoho Creator API", font=("Segoe UI", 9, "bold"), bg=WHITE, variable=self.var_push_zoho).grid(row=3, column=1, sticky="w", padx=5, pady=10)

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
        footer = tk.Frame(self.root, bg=WHITE, height=50)
        footer.pack(fill="x", side="bottom", padx=20, pady=10)
        
        tk.Label(footer, text="© Nagarkot Forwarders Pvt Ltd", font=("Segoe UI", 8), fg="#6c757d", bg=WHITE).pack(side="left")
        
        tk.Button(
            footer, text="Generate DSR", font=("Segoe UI", 10, "bold"),
            bg=BTN_BLUE, fg=WHITE, activebackground="#004494", activeforeground=WHITE,
            width=25, height=2, borderwidth=0, cursor="hand2",
            command=self._on_generate
        ).pack(side="right")

    # ═══ Actions ═════════════════════════════════════════════════════════

    def _on_select_pdfs(self) -> None:
        files = filedialog.askopenfilenames(
            title="Select BL & Invoice PDFs",
            filetypes=[("PDF Files", "*.pdf")],
        )
        if files:
            if not self.var_out_folder.get():
                self.var_out_folder.set(str(Path(files[0]).parent))
            
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
            if not self.var_out_folder.get():
                self.var_out_folder.set(str(p))
            
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
        self.tree.delete(*self.tree.get_children())
        self.lbl_file_status.config(text="No files selected")
        self.var_out_folder.set("")

    def _on_browse_master(self) -> None:
        f = filedialog.askopenfilename(
            title="Select Master DSR",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if f:
            self.master_dsr_path = Path(f)
            self.var_master_dsr.set(str(self.master_dsr_path))

    def _on_clear_master(self) -> None:
        self.master_dsr_path = None
        self.var_master_dsr.set("")

    def _on_browse_out_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.var_out_folder.set(folder)

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
                if stem_upper.startswith("MAEU") or stem_upper.startswith("HLCU") or stem_upper == "BL" or stem_upper.startswith("MEAU"):
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
            
            import fitz
            import re
            
            container_to_invoices = {rec.container_no.upper(): set() for rec in dir_records}
            unmapped_invoices = set()
            all_invoice_stems = []
            
            for inv_f in invoice_pdf_files:
                matches = re.findall(r"\b([45]\d{7})\b", inv_f.stem)
                inv_no = matches[0] if matches else inv_f.stem
                all_invoice_stems.append(inv_no)
                
                try:
                    doc = fitz.open(str(inv_f))
                    text_all = "".join(page.get_text().upper() for page in doc)
                    doc.close()
                    
                    found_containers = set(re.findall(r"\b([A-Z]{4}\d{7})\b", text_all))
                    mapped = False
                    for c_no in found_containers:
                        if c_no in container_to_invoices:
                            container_to_invoices[c_no].add(inv_no)
                            mapped = True
                    
                    if not mapped:
                        unmapped_invoices.add(inv_no)
                        
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
                    
                    if mapped_set:
                        # Exclusively use invoices mapped specifically to this container
                        rec.invoice_nos = "/".join(sorted(list(mapped_set)))
                    else:
                        # Fallback: Assign unmapped invoices or all invoices if we couldn't confidently tie them
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

    def _on_generate(self) -> None:
        if not self.parsed_records:
            messagebox.showwarning("No Data", "No valid parsed container records to generate DSR.")
            return
            
        # Before spawning preview, explicitly assign global manual settings to ALL the records
        global_user = self.var_user.get().strip()
        global_month = self.var_month.get().strip()
        global_pre = self.cal_pre_alert.get()
        global_eta = self.cal_vessel_eta.get()
        
        for r in self.parsed_records:
            r.user = global_user
            r.user_month = global_month
            r.pre_alert_date = global_pre
            r.vessel_eta = global_eta

        # Pop up the Data Preview modal. Only if they confirm do we proceed.
        DataPreviewWindow(self.root, self.parsed_records, self._process_generation)

    def _process_generation(self) -> None:
        out_name = self.var_out_filename.get().strip()
        if not out_name.endswith(".xlsx"):
            out_name += ".xlsx"
        
        out_folder = self.var_out_folder.get().strip()
        if not out_folder:
            out_folder = str(SCRIPT_DIR)
        
        save_path = Path(out_folder) / out_name

        try:
            if self.master_dsr_path and self.master_dsr_path.exists():
                # Append directly to the master file in its original location
                self._append_to_master(self.master_dsr_path)
                excel_msg = f"DSR data appended successfully to Master file!\n\nAppended to: {self.master_dsr_path.name}"
            else:
                self._create_new_dsr(save_path)
                excel_msg = f"DSR generated successfully!\n\nSaved to: {save_path.name}"
            
            # API Push
            zoho_msg = ""
            if self.var_push_zoho.get():
                try:
                    zoho = ZohoCreatorAPI()
                    success, msg = zoho.push_records(self.parsed_records)
                    if success:
                        zoho_msg = f"\n\nZoho API: {msg}"
                    else:
                        zoho_msg = f"\n\nZoho API Warning: {msg}"
                except Exception as ze:
                    zoho_msg = f"\n\nZoho API Error: {ze}"
            
            messagebox.showinfo("Success", excel_msg + zoho_msg)

        except Exception as exc:
            logger.exception("Failed to generate DSR")
            err_msg = str(exc)
            if "old .xls file format" in err_msg or "InvalidFileException" in str(type(exc)):
                messagebox.showerror("Export Error", "The Master DSR file must be a valid .xlsx file.\nOlder .xls formats are not supported.\n\nPlease open the master file in Excel and 'Save As' -> .xlsx first.")
            else:
                messagebox.showerror("Export Error", f"An error occurred while generating DSR:\n{exc}")

    # ── Excel Export Logic ───────────────────────────────────────────────

    def _record_to_row(self, rec: ContainerRecord) -> list:
        row = [""] * len(DSR_HEADERS)
        
        # Format the user-editable dates on-the-fly for Excel (using datetime objects looks better in Excel)
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

        # Parse BL Date object
        bl_dt = None
        try:
            bl_dt = datetime.strptime(rec.bl_date, "%Y-%m-%d")
        except:
            pass
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

    def _style_header_row(self, ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="1B3A5C")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

        for col_idx, header in enumerate(DSR_HEADERS, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        ws.row_dimensions[1].height = 40
        ws.freeze_panes = "A2"

    def _create_new_dsr(self, save_path: Path) -> None:
        wb = openpyxl.Workbook()
        ws_live = wb.active
        ws_live.title = "Live shipments"
        self._style_header_row(ws_live)

        data_font = Font(name="Segoe UI", size=9)
        data_align = Alignment(vertical="center", wrap_text=False)
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

        for row_idx, rec in enumerate(self.parsed_records, 2):
            row_data = self._record_to_row(rec)
            for col_idx, value in enumerate(row_data, 1):
                cell = ws_live.cell(row=row_idx, column=col_idx, value=value)
                cell.font = data_font
                cell.alignment = data_align
                cell.border = thin_border
                if col_idx in (2, 7, 8):
                    cell.number_format = "YYYY-MM-DD"

        for col_idx in range(1, len(DSR_HEADERS) + 1):
            ws_live.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 15

        ws_cleared = wb.create_sheet("Cleared shipments")
        self._style_header_row(ws_cleared)
        for col_idx in range(1, len(DSR_HEADERS) + 1):
            ws_cleared.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 15

        wb.save(save_path)
        wb.close()

    def _append_to_master(self, master_path: Path) -> None:
        wb = openpyxl.load_workbook(master_path)

        sheet_name = next((name for name in wb.sheetnames if "live" in name.lower()), wb.sheetnames[0])
        ws = wb[sheet_name]
        next_row = ws.max_row + 1

        data_font = Font(name="Segoe UI", size=9)
        data_align = Alignment(vertical="center", wrap_text=False)
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

        for rec in self.parsed_records:
            row_data = self._record_to_row(rec)
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=next_row, column=col_idx, value=value)
                cell.font = data_font
                cell.alignment = data_align
                cell.border = thin_border
                if col_idx in (2, 7, 8):
                    cell.number_format = "YYYY-MM-DD"
            next_row += 1

        wb.save(master_path)
        wb.close()


def main() -> None:
    root = tk.Tk()
    app = DSRGeneratorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
