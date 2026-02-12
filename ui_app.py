import os
import datetime
import sys
from pathlib import Path
def _set_playwright_paths():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS).resolve() 
    elif getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        base = Path(__file__).resolve().parent
    candidates = [
        base / "playwright" / "driver" / "package",
        base / "_internal" / "playwright" / "driver" / "package",
    ]
    pw_pkg = next((p for p in candidates if p.exists()), None)
    if not pw_pkg:
        print("Playwright package not found. Tried:")
        for c in candidates:
            print("  -", c)
        return
    pw_browsers = pw_pkg / ".local-browsers"
    os.environ["PLAYWRIGHT_DRIVER_SEARCH_PATH"] = str(pw_pkg)
    if pw_browsers.exists():
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(pw_browsers)
_set_playwright_paths()
print("PW_DRIVER_SEARCH_PATH =", os.environ.get("PLAYWRIGHT_DRIVER_SEARCH_PATH"))
print("PW_BROWSERS_PATH      =", os.environ.get("PLAYWRIGHT_BROWSERS_PATH"))
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from scrape_and_generate import scrape_complaint, fill_docx
import re
DEFAULT_CONFIG_PATH = Path("config.yaml")
BG_LIGHT = "#f3f4f6"
CARD_BG = "#ffffff"
TEXT_DARK = "#111827"
TEXT_MUTED = "#6b7280"
ACCENT = "#2563eb"
def extract_country_from_address(address: str) -> str:
    if not address:
        return "USA"
    address_upper = address.upper()
    countries = {
        "USA": ["USA", "U.S.A", "U.S.A.", "UNITED STATES"],
        "Canada": ["CANADA"],
        "UK": ["UK", "U.K.", "UNITED KINGDOM", "ENGLAND", "SCOTLAND", "WALES"],
        "Germany": ["GERMANY", "DEUTSCHLAND"],
        "France": ["FRANCE"],
        "Italy": ["ITALY", "ITALIA"],
        "Spain": ["SPAIN", "ESPAÑA"],
        "Mexico": ["MEXICO", "MÉXICO"],
        "Australia": ["AUSTRALIA"],
        "Japan": ["JAPAN"],
        "China": ["CHINA", "PRC"],
    }
    for country, indicators in countries.items():
        for indicator in indicators:
            if indicator in address_upper:
                return country
    return "USA"
def parse_ir_address_block(ir_block: str) -> dict:
    if not ir_block:
        return {
            "ir_name": "",
            "facility_name": "",
            "facility_address": "",
            "country": "USA"
        }
    lines = [line.strip() for line in ir_block.strip().split('\n') if line.strip()]
    if len(lines) == 0:
        return {
            "ir_name": "",
            "facility_name": "",
            "facility_address": "",
            "country": "USA"
        }
    elif len(lines) == 1:
        return {
            "ir_name": lines[0],
            "facility_name": "",
            "facility_address": "",
            "country": "USA"
        }
    elif len(lines) == 2:
        line2_lower = lines[1].lower()
        has_address_markers = any(marker in line2_lower for marker in [
            'street', 'st ', 'avenue', 'ave ', 'road', 'rd ', 'drive', 
            'lane', 'blvd', 'plaza', '-', 'se-', 'ne-', 'sw-', 'nw-'
        ])
        if has_address_markers or len(lines[1]) > 50:
            return {
                "ir_name": lines[0],
                "facility_name": "",
                "facility_address": lines[1],
                "country": extract_country_from_address(lines[1])
            }
        else:
            return {
                "ir_name": lines[0],
                "facility_name": lines[1],
                "facility_address": "",
                "country": "USA"
            }
    else:
        ir_name = lines[0]
        facility_name = lines[1]
        facility_address = "\n".join(lines[2:])
        country = extract_country_from_address(facility_address)
        return {
            "ir_name": ir_name,
            "facility_name": facility_name,
            "facility_address": facility_address,
            "country": country
        }
def build_ir_address_block(ir_name: str, facility_name: str, facility_address: str, country: str) -> str:
    parts = []
    if ir_name.strip():
        parts.append(ir_name.strip())
    if facility_name.strip():
        parts.append(facility_name.strip())
    if facility_address.strip():
        parts.append(facility_address.strip())
    if country.strip():
        parts.append(country.strip())
    return "\n".join(parts)
def _normalize_newlines(text: str) -> str:
    if not text:
        return ""
    return text.replace("\r\n", "\n").replace("\r", "\n")
def _remove_exact_unwanted_line(text: str) -> str:
    text = _normalize_newlines(text)
    text = re.sub(
        r'^[ \t]*The product sample was not returned to the Visual inspection:[ \t]*\n?',
        '',
        text,
        flags=re.MULTILINE
    )
    return text
def _remove_picture_evaluation_bullets_only(text: str) -> str:
    text = _normalize_newlines(text)

    lines = text.split("\n")
    out = []
    i = 0
    n = len(lines)
    bullet_re = re.compile(r'^\s*(?:•|-|\*|\u2022)\s+')
    section_header_re = re.compile(r'^\s*[A-Za-z][^:\n]{0,80}:\s*$')
    while i < n:
        line = lines[i]
        out.append(line)
        if re.match(r'^\s*Picture Evaluation:\s*$', line, flags=re.IGNORECASE):
            j = i + 1
            while j < n and lines[j].strip() == "":
                j += 1
            while j < n and bullet_re.match(lines[j]):
                j += 1
            while j < n and lines[j].strip() == "":
                if j + 1 < n and section_header_re.match(lines[j + 1]):
                    break
                j += 1

            i = j
            continue
        i += 1
    cleaned = "\n".join(out)
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned).strip()
    return cleaned
def sanitize_analysis_text(text: str) -> str:
    text = _remove_exact_unwanted_line(text)
    text = _remove_picture_evaluation_bullets_only(text)
    return text
class CustomerLetterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Customer Letter Generator")
        self.minsize(900, 600)
        self.configure(bg=BG_LIGHT)
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Main.TFrame", background=BG_LIGHT)
        style.configure(
            "Card.TFrame",
            background=CARD_BG,
            relief="groove",
            borderwidth=1,
        )
        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 18, "bold"),
            foreground=TEXT_DARK,
            background=BG_LIGHT,
        )
        style.configure(
            "Step.TLabel",
            font=("Segoe UI", 10, "bold"),
            foreground=TEXT_MUTED,
            background=BG_LIGHT,
        )
        style.configure(
            "CardTitle.TLabel",
            font=("Segoe UI", 14, "bold"),
            foreground=TEXT_DARK,
            background=CARD_BG,
        )
        style.configure(
            "CardText.TLabel",
            font=("Segoe UI", 10),
            foreground=TEXT_MUTED,
            background=CARD_BG,
            wraplength=700,
        )
        style.configure(
            "Accent.TButton",
            font=("Segoe UI", 10, "bold"),
            foreground="#ffffff",
            background=ACCENT,
            padding=6,
            borderwidth=0,
        )
        style.map("Accent.TButton", background=[("active", "#1d4ed8")])
        style.configure(
            "Ghost.TButton",
            font=("Segoe UI", 10),
            foreground=TEXT_DARK,
            background="#e5e7eb",
            padding=6,
            borderwidth=0,
        )
        style.map("Ghost.TButton", background=[("active", "#d1d5db")])
        self.values = {}
        self.products = []
        self.cfg = None
        self.template_path = None
        self.out_dir = None
        self.last_saved_path = None
        self.current_analysis_idx = 0        # 0-based
        self.current_investigation_idx = 0   # 0-based
        self.complaint_var = tk.StringVar()
        self.status_var = tk.StringVar()
        header = ttk.Frame(self, style="Main.TFrame", padding=(20, 15, 20, 5))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        self.title_label = ttk.Label(
            header,
            text="Customer Letter Generator",
            style="Title.TLabel",
        )
        self.title_label.grid(row=0, column=0, sticky="w")
        self.step_label = ttk.Label(
            header,
            text="Step 1 of 4 · Enter GCH PE Number",
            style="Step.TLabel",
        )
        self.step_label.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.main_frame = ttk.Frame(self, style="Main.TFrame", padding=(20, 10, 20, 20))
        self.main_frame.grid(row=1, column=0, sticky="nsew")
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)
        self.step1_frame = ttk.Frame(self.main_frame, style="Card.TFrame", padding=20)
        self.step2_frame = ttk.Frame(self.main_frame, style="Card.TFrame", padding=20)
        self.step3_frame = ttk.Frame(self.main_frame, style="Card.TFrame", padding=20)  # Analyses per product
        self.step4_frame = ttk.Frame(self.main_frame, style="Card.TFrame", padding=20)  # Investigations per product + save
        self._build_step1()
        self._build_step2()
        self._build_step3_analysis()
        self._build_step4_investigation_per_product()
        self._show_step(self.step1_frame, "Step 1 of 4 · Enter GCH PE Number")
    def _collect_debug_info(self):
        from datetime import datetime
        import json
        info_lines = [
            "=== Customer Letter Generator Debug Info ===",
            f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"PE Number: {self.complaint_var.get()}",
            "",
        ]
        info_lines.append("=== Scraping Log Output ===")
        info_lines.append("")
        try:
            log_files = sorted(Path(".").glob("customer_letter_debug_*.log"), 
                            key=lambda p: p.stat().st_mtime, 
                            reverse=True)
            if log_files:
                recent_log = log_files[0]
                info_lines.append(f"Log file: {recent_log.name}")
                info_lines.append("")
                with open(recent_log, 'r', encoding='utf-8') as f:
                    log_content = f.read()
                    info_lines.append(log_content)
            else:
                info_lines.append("No log file found - data may not have been scraped yet")
        except Exception as e:
            info_lines.append(f"Error reading log: {e}")
        info_lines.append("")
        info_lines.append("=== Values Dictionary ===")
        info_lines.append(json.dumps(self.values, indent=2))
        info_lines.append("")
        info_lines.append("=== Products List ===")
        info_lines.append(json.dumps(self.products, indent=2))
        return "\n".join(info_lines)
    def _copy_debug_to_clipboard(self):
        try:
            debug_info = self._collect_debug_info()
            self.clipboard_clear()
            self.clipboard_append(debug_info)
            self.update()  # Ensure clipboard is updated
            messagebox.showinfo(
                "Debug Info Copied", 
                "Debug information has been copied to clipboard.\n\n"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy debug info: {e}")
    def _build_step1(self):
        f = self.step1_frame
        ttk.Label(f, text="Enter GCH PE Number", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            f,
            text="We'll pull initial reporter data, event description, product info, analyses, and investigations from GCH.",
            style="CardText.TLabel",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(5, 15))
        ttk.Label(f, text="PE number:", style="CardText.TLabel").grid(
            row=2, column=0, sticky="w", pady=(0, 5)
        )
        entry = ttk.Entry(f, textvariable=self.complaint_var, width=30)
        entry.grid(row=2, column=1, sticky="w", padx=(5, 0), pady=(0, 5))
        go_btn = ttk.Button(f, text="Go", style="Accent.TButton", command=self.on_go_clicked)
        go_btn.grid(row=2, column=2, padx=(10, 0), pady=(0, 5))
        status = ttk.Label(f, textvariable=self.status_var, style="CardText.TLabel")
        status.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
        f.grid_columnconfigure(1, weight=1)
    def _bring_to_front(self):
        try:
            self.lift()
            self.focus_force()
            self.attributes('-topmost', True)
            self.after(100, lambda: self.attributes('-topmost', False))
        except Exception:
            pass
    def on_go_clicked(self):
        complaint_id = self.complaint_var.get().strip()
        if not complaint_id:
            messagebox.showerror("Missing complaint number", "Please enter a complaint number first.")
            return
        self.status_var.set("Contacting GCH and collecting data…")
        self.update_idletasks()
        try:
            values, products, cfg, template_path, out_dir = scrape_complaint(
                complaint_id,
                str(DEFAULT_CONFIG_PATH),
            )
        except Exception as e:
            self.status_var.set("")
            messagebox.showerror("Error", f"Failed to scrape data from GCH:\n{e}")
            return
        self.values = values or {}
        self.products = products or []
        self.cfg = cfg
        self.template_path = template_path
        self.out_dir = out_dir
        self.last_saved_path = None
        ir_block = self.values.get("ir_with_address", "") or ""
        ir_name_from_values = self.values.get("ir_name", "") or ""
        parsed = parse_ir_address_block(ir_block)
        if ir_name_from_values:
            if parsed["ir_name"] and parsed["ir_name"] != ir_name_from_values:
                if not parsed["facility_name"]:
                    parsed["facility_name"] = parsed["ir_name"]
            parsed["ir_name"] = ir_name_from_values
        self.ir_name_entry.delete(0, tk.END)
        self.ir_name_entry.insert(0, parsed["ir_name"])
        self.facility_name_entry.delete(0, tk.END)
        self.facility_name_entry.insert(0, parsed["facility_name"])
        self.facility_address_text.delete("1.0", "end")
        self.facility_address_text.insert("1.0", parsed["facility_address"])
        self.country_entry.delete(0, tk.END)
        self.country_entry.insert(0, parsed["country"])
        self.status_var.set("")
        self._bring_to_front()
        self._show_step(self.step2_frame, "Step 2 of 4 · Edit Initial Reporter / Address")
    def _build_step2(self):
        f = self.step2_frame
        ttk.Label(f, text="Initial Reporter & Facility Address", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(
            f,
            text="Review and adjust the recipient information. This will appear in the letter heading.",
            style="CardText.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 15))
        ttk.Label(f, text="Initial Reporter Name:", style="CardText.TLabel").grid(
            row=2, column=0, sticky="w", pady=(0, 5)
        )
        self.ir_name_entry = ttk.Entry(f, width=60)
        self.ir_name_entry.grid(row=2, column=1, sticky="ew", pady=(0, 5), padx=(10, 0))
        ttk.Label(f, text="Facility Name:", style="CardText.TLabel").grid(
            row=3, column=0, sticky="w", pady=(0, 5)
        )
        self.facility_name_entry = ttk.Entry(f, width=60)
        self.facility_name_entry.grid(row=3, column=1, sticky="ew", pady=(0, 5), padx=(10, 0))
        ttk.Label(f, text="Facility Address:", style="CardText.TLabel").grid(
            row=4, column=0, sticky="nw", pady=(0, 5)
        )
        self.facility_address_text = tk.Text(
            f,
            width=60,
            height=3,
            wrap="word",
            bg="#ffffff",
            fg=TEXT_DARK,
            insertbackground=TEXT_DARK,
            relief="solid",
            borderwidth=1,
        )
        self.facility_address_text.grid(row=4, column=1, sticky="ew", pady=(0, 5), padx=(10, 0))
        ttk.Label(f, text="Country:", style="CardText.TLabel").grid(
            row=5, column=0, sticky="w", pady=(0, 10)
        )
        self.country_entry = ttk.Entry(f, width=60)
        self.country_entry.grid(row=5, column=1, sticky="ew", pady=(0, 10), padx=(10, 0))
        debug_btn = ttk.Button(
            f, 
            text="Copy Debug Info", 
            style="Ghost.TButton", 
            command=self._copy_debug_to_clipboard
        )
        debug_btn.grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 10))
        btn_frame = ttk.Frame(f, style="Card.TFrame")
        btn_frame.grid(row=7, column=0, columnspan=2, sticky="e")
        back_btn = ttk.Button(
            btn_frame,
            text="← Back",
            style="Ghost.TButton",
            command=lambda: self._show_step(
                self.step1_frame, "Step 1 of 4 · Enter GCH PE Number"
            ),
        )
        back_btn.grid(row=0, column=0, padx=5)
        next_btn = ttk.Button(
            btn_frame,
            text="Next · Analyses →",
            style="Accent.TButton",
            command=self.on_ir_next,
        )
        next_btn.grid(row=0, column=1, padx=5)
        f.grid_columnconfigure(1, weight=1)
    def on_ir_next(self):
        ir_name = self.ir_name_entry.get().strip()
        facility_name = self.facility_name_entry.get().strip()
        facility_address = self.facility_address_text.get("1.0", "end-1c").strip()
        country = self.country_entry.get().strip()
        self.values["ir_with_address"] = build_ir_address_block(
            ir_name, facility_name, facility_address, country
        )
        if not self.products:
            self._show_step(
                self.step4_frame,
                "Step 3 of 4 · Edit Investigations (per product) & Save",
            )
            return
        self.current_analysis_idx = 0
        self._load_current_analysis()
        self._show_step(
            self.step3_frame, "Step 3 of 4 · Edit Analyses (per product)"
        )
    def _build_step3_analysis(self):
        f = self.step3_frame
        self.analysis_header_label = ttk.Label(f, text="Product Analysis", style="CardTitle.TLabel")
        self.analysis_header_label.grid(row=0, column=0, sticky="w")
        self.analysis_product_label = ttk.Label(
            f,
            text="",
            style="CardText.TLabel",
        )
        self.analysis_product_label.grid(row=1, column=0, sticky="w", pady=(4, 10))
        self.analysis_text_widget = tk.Text(
            f,
            width=100,
            height=12,
            wrap="word",
            bg="#ffffff",
            fg=TEXT_DARK,
            insertbackground=TEXT_DARK,
            relief="solid",
            borderwidth=1,
        )
        self.analysis_text_widget.grid(row=2, column=0, sticky="nsew", pady=(5, 10))
        btn_frame = ttk.Frame(f, style="Card.TFrame")
        btn_frame.grid(row=3, column=0, sticky="e")
        self.analysis_back_btn = ttk.Button(
            btn_frame, text="← Back", style="Ghost.TButton", command=self.on_analysis_back
        )
        self.analysis_back_btn.grid(row=0, column=0, padx=5)
        self.analysis_next_btn = ttk.Button(
            btn_frame,
            text="Next Analysis →",
            style="Accent.TButton",
            command=self.on_analysis_next,
        )
        self.analysis_next_btn.grid(row=0, column=1, padx=5)
        f.grid_rowconfigure(2, weight=1)
        f.grid_columnconfigure(0, weight=1)
    def _load_current_analysis(self):
        if not self.products:
            self.analysis_product_label.config(text="No products found.")
            self.analysis_text_widget.delete("1.0", "end")
            return
        idx = self.current_analysis_idx  # 0-based
        product_num = idx + 1
        prod = self.products[idx] if idx < len(self.products) else {}
        pid = (prod.get("id") or self.values.get(f"product_id_{product_num}", "") or "").strip()
        pdesc = (prod.get("desc") or self.values.get(f"product_desc_{product_num}", "") or "").strip()
        header_text = f"Analysis for Product {product_num} of {len(self.products)}"
        self.analysis_header_label.config(text=header_text)
        product_line = f"Product: {pid or '(no ID)'}"
        if pdesc:
            product_line += f" — {pdesc}"
        self.analysis_product_label.config(text=product_line)
        analysis_key = f"analysis_{product_num}"
        analysis_text = self.values.get(analysis_key, "") or ""
        analysis_text = sanitize_analysis_text(analysis_text)
        self.values[analysis_key] = analysis_text
        self.analysis_text_widget.delete("1.0", "end")
        self.analysis_text_widget.insert("1.0", analysis_text)
        if product_num == len(self.products):
            self.analysis_next_btn.config(text="Next · Investigations →")
        else:
            self.analysis_next_btn.config(text="Next Analysis →")
    def _save_current_analysis(self):
        if not self.products:
            return
        idx = self.current_analysis_idx
        product_num = idx + 1
        analysis_key = f"analysis_{product_num}"
        text = self.analysis_text_widget.get("1.0", "end-1c")
        text = sanitize_analysis_text(text)
        self.values[analysis_key] = text
    def on_analysis_next(self):
        if self.products:
            self._save_current_analysis()
        if not self.products:
            self._show_step(
                self.step4_frame,
                "Step 3 of 4 · Edit Investigations (per product) & Save",
            )
            return
        if self.current_analysis_idx >= len(self.products) - 1:
            self.current_investigation_idx = 0
            self._load_current_investigation()
            self._show_step(
                self.step4_frame,
                "Step 4 of 4 · Edit Investigations (per product) & Save",
            )
        else:
            self.current_analysis_idx += 1
            self._load_current_analysis()
            self._show_step(
                self.step3_frame,
                "Step 3 of 4 · Edit Analyses (per product)",
            )
    def on_analysis_back(self):
        if not self.products:
            self._show_step(self.step2_frame, "Step 2 of 4 · Edit Initial Reporter / Address")
            return
        self._save_current_analysis()
        if self.current_analysis_idx == 0:
            self._show_step(self.step2_frame, "Step 2 of 4 · Edit Initial Reporter / Address")
        else:
            self.current_analysis_idx -= 1
            self._load_current_analysis()
            self._show_step(
                self.step3_frame,
                "Step 3 of 4 · Edit Analyses (per product)",
            )
    def _build_step4_investigation_per_product(self):
        f = self.step4_frame
        self.inv_pp_header_label = ttk.Label(
            f, text="Product Investigation", style="CardTitle.TLabel"
        )
        self.inv_pp_header_label.grid(row=0, column=0, sticky="w")
        self.inv_pp_product_label = ttk.Label(
            f,
            text="",
            style="CardText.TLabel",
        )
        self.inv_pp_product_label.grid(row=1, column=0, sticky="w", pady=(4, 10))
        self.inv_pp_text_widget = tk.Text(
            f,
            width=100,
            height=12,
            wrap="word",
            bg="#ffffff",
            fg=TEXT_DARK,
            insertbackground=TEXT_DARK,
            relief="solid",
            borderwidth=1,
        )
        self.inv_pp_text_widget.grid(row=2, column=0, sticky="nsew", pady=(5, 10))
        self.saved_link_label = ttk.Label(
            f,
            text="",
            style="CardText.TLabel",
            foreground="#1d4ed8",
        )
        self.saved_link_label.grid(row=3, column=0, sticky="w", pady=(0, 5))
        self.saved_link_label.bind("<Button-1>", self._on_saved_link_click)
        self.saved_link_label.configure(cursor="hand2")
        btn_frame = ttk.Frame(f, style="Card.TFrame")
        btn_frame.grid(row=4, column=0, sticky="e")
        self.inv_pp_open_btn = ttk.Button(
            btn_frame,
            text="Open Letter",
            style="Ghost.TButton",
            command=self.on_open_letter,
            state="disabled",
        )
        self.inv_pp_open_btn.grid(row=0, column=0, padx=5)
        self.inv_pp_restart_btn = ttk.Button(
            btn_frame,
            text="Start Over",
            style="Ghost.TButton",
            command=self.on_start_over,
            state="disabled",
        )
        self.inv_pp_restart_btn.grid(row=0, column=1, padx=5)
        self.inv_pp_back_btn = ttk.Button(
            btn_frame,
            text="← Back",
            style="Ghost.TButton",
            command=self.on_inv_pp_back,
        )
        self.inv_pp_back_btn.grid(row=0, column=2, padx=5)
        self.inv_pp_next_btn = ttk.Button(
            btn_frame,
            text="Next Investigation →",
            style="Accent.TButton",
            command=self.on_inv_pp_next,
        )
        self.inv_pp_next_btn.grid(row=0, column=3, padx=5)
        f.grid_rowconfigure(2, weight=1)
        f.grid_columnconfigure(0, weight=1)
    def _load_current_investigation(self):
        if not self.products:
            self.inv_pp_product_label.config(text="No products found.")
            self.inv_pp_text_widget.delete("1.0", "end")
            return
        idx = self.current_investigation_idx
        product_num = idx + 1
        prod = self.products[idx] if idx < len(self.products) else {}
        pid = (prod.get("id") or self.values.get(f"product_id_{product_num}", "") or "").strip()
        pdesc = (prod.get("desc") or self.values.get(f"product_desc_{product_num}", "") or "").strip()
        header_text = f"Investigation for Product {product_num} of {len(self.products)}"
        self.inv_pp_header_label.config(text=header_text)
        product_line = f"Product: {pid or '(no ID)'}"
        if pdesc:
            product_line += f" — {pdesc}"
        self.inv_pp_product_label.config(text=product_line)
        inv_key = f"investigation_{product_num}"
        inv_text = self.values.get(inv_key, "") or ""
        self.inv_pp_text_widget.config(state="normal")
        self.inv_pp_text_widget.delete("1.0", "end")
        self.inv_pp_text_widget.insert("1.0", inv_text)
        if product_num == len(self.products):
            self.inv_pp_next_btn.config(text="Save Letter…")
        else:
            self.inv_pp_next_btn.config(text="Next Investigation →")
    def _save_current_investigation(self):
        if not self.products:
            return
        idx = self.current_investigation_idx
        product_num = idx + 1
        inv_key = f"investigation_{product_num}"
        text = self.inv_pp_text_widget.get("1.0", "end-1c")
        self.values[inv_key] = text
    def on_inv_pp_next(self):
        if self.products:
            self._save_current_investigation()
        if not self.products:
            self.on_save_clicked()
            return
        if self.current_investigation_idx >= len(self.products) - 1:
            self.on_save_clicked()
        else:
            self.current_investigation_idx += 1
            self._load_current_investigation()
            self._show_step(
                self.step4_frame,
                "Step 4 of 4 · Edit Investigations (per product) & Save",
            )
    def on_inv_pp_back(self):
        if not self.products:
            self._show_step(
                self.step2_frame,
                "Step 2 of 4 · Edit Initial Reporter / Address",
            )
            return
        self._save_current_investigation()
        if self.current_investigation_idx == 0:
            self.current_analysis_idx = max(0, len(self.products) - 1)
            self._load_current_analysis()
            self._show_step(
                self.step3_frame,
                "Step 3 of 4 · Edit Analyses (per product)",
            )
        else:
            self.current_investigation_idx -= 1
            self._load_current_investigation()
            self._show_step(
                self.step4_frame,
                "Step 4 of 4 · Edit Investigations (per product) & Save",
            )
    def _update_combined_analysis_results(self):
        if not self.products:
            return
        blocks = []
        for i in range(1, len(self.products) + 1):
            txt = (self.values.get(f"analysis_{i}") or "").strip()
            if txt:
                blocks.append(txt)
        if blocks:
            self.values["analysis_results"] = "\n\n".join(blocks)
    def _update_combined_investigation_results(self):
        if not self.products:
            return
        blocks = []
        for i in range(1, len(self.products) + 1):
            txt = (self.values.get(f"investigation_{i}") or "").strip()
            if txt:
                blocks.append(txt)
        if blocks:
            self.values["investigation_summary"] = "\n\n\n".join(blocks)
    def on_save_clicked(self):
        self._update_combined_analysis_results()
        self._update_combined_investigation_results()
        if not self.template_path:
            messagebox.showerror("Error", "Template path is not set.")
            return
        pattern = (self.cfg or {}).get("output_name_pattern", "Customer_Letter_{complaint_id}.docx")
        try:
            default_name = pattern.format(**self.values)
        except Exception:
            default_name = f"Customer_Letter_{self.values.get('complaint_id', 'letter')}.docx"
        initial_dir = str(self.out_dir) if self.out_dir else "."
        save_path = filedialog.asksaveasfilename(
            title="Save generated letter",
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
            initialdir=initial_dir,
            initialfile=default_name,
        )
        if not save_path:
            return
        try:
            fill_docx(str(self.template_path), save_path, self.values, self.products)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate document:\n{e}")
            return
        self.last_saved_path = save_path
        msg = f"Letter saved to: {save_path}"
        self.saved_link_label.config(text=msg)
        self.inv_pp_open_btn.config(state="normal")
        self.inv_pp_restart_btn.config(state="normal")
        self.inv_pp_text_widget.config(state="disabled")
        self.inv_pp_next_btn.config(state="disabled")
        messagebox.showinfo("Done", msg)
    def _on_saved_link_click(self, event):
        if self.last_saved_path:
            self._open_file(self.last_saved_path)
    def on_open_letter(self):
        if self.last_saved_path:
            self._open_file(self.last_saved_path)
    def on_start_over(self):
        self.values = {}
        self.products = []
        self.cfg = None
        self.template_path = None
        self.out_dir = None
        self.last_saved_path = None
        self.current_analysis_idx = 0
        self.current_investigation_idx = 0
        self.complaint_var.set("")
        self.status_var.set("")
        self.ir_name_entry.delete(0, tk.END)
        self.facility_name_entry.delete(0, tk.END)
        self.facility_address_text.delete("1.0", "end")
        self.country_entry.delete(0, tk.END)
        self.analysis_text_widget.config(state="normal")
        self.analysis_text_widget.delete("1.0", "end")
        self.inv_pp_text_widget.config(state="normal")
        self.inv_pp_text_widget.delete("1.0", "end")
        self.saved_link_label.config(text="")
        self.inv_pp_open_btn.config(state="disabled")
        self.inv_pp_restart_btn.config(state="disabled")
        self.inv_pp_next_btn.config(state="normal")
        self._show_step(self.step1_frame, "Step 1 of 4 · Enter GCH PE Number")
    def _open_file(self, path: str):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")
    def _show_step(self, frame_to_show: ttk.Frame, step_text: str):
        for f in (self.step1_frame, self.step2_frame, self.step3_frame, self.step4_frame):
            f.grid_forget()
        frame_to_show.grid(row=0, column=0, sticky="nsew")
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.step_label.config(text=step_text)
if __name__ == "__main__":
    app = CustomerLetterApp()
    app.mainloop()