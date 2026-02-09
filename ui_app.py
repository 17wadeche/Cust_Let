import os
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
DEFAULT_CONFIG_PATH = Path("config.yaml")
BG_LIGHT = "#f3f4f6"
CARD_BG = "#ffffff"
TEXT_DARK = "#111827"
TEXT_MUTED = "#6b7280"
ACCENT = "#2563eb"
class PartnerSelectionDialog(tk.Toplevel):
    def __init__(self, master, partners):
        super().__init__(master)
        self.title("Select Letter Recipients")
        self.resizable(False, False)
        self.configure(bg=BG_LIGHT)
        self.partners = partners
        self.result_name_index = None
        self.result_address_index = None
        self.name_var = tk.StringVar()
        self.address_var = tk.StringVar()
        self.transient(master)
        self.grab_set()
        ttk.Label(
            self,
            text="Multiple partners found. Please select which to use for the letter.",
            style="CardText.TLabel",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(20, 10))
        ttk.Label(self, text="Select Name:", style="CardText.TLabel").grid(
            row=1, column=0, sticky="w", padx=20, pady=(10, 5)
        )
        self.name_combo = ttk.Combobox(
            self, textvariable=self.name_var, state="readonly", width=95
        )
        self.name_combo.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        ttk.Label(self, text="Select Address:", style="CardText.TLabel").grid(
            row=3, column=0, sticky="w", padx=20, pady=(10, 5)
        )
        self.address_combo = ttk.Combobox(
            self, textvariable=self.address_var, state="readonly", width=95
        )
        self.address_combo.grid(row=4, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.name_options = []
        self.address_options = []
        for c in partners:
            pf = c.get("partner_function", "")
            name = c.get("name", "")
            addr = c.get("address", "").replace("\n", " / ")
            name_display = f"{pf}: {name}" if name else f"{pf}: (no name)"
            addr_display = f"{pf}: {addr}" if addr else f"{pf}: (no address)"
            self.name_options.append(name_display)
            self.address_options.append(addr_display)
        self.name_combo["values"] = self.name_options
        self.address_combo["values"] = self.address_options
        default_name_idx = self._find_partner_index(["Initial Reporter", "Initial Contact"])
        default_addr_idx = self._find_partner_index(["Facility", "Health Care Facility", "Healthcare Facility"])
        if partners:
            self.name_combo.current(default_name_idx if default_name_idx is not None else 0)
            self.address_combo.current(default_addr_idx if default_addr_idx is not None else 0)
        btn_frame = ttk.Frame(self, style="Main.TFrame")
        btn_frame.grid(row=5, column=0, sticky="e", padx=20, pady=(0, 20))
        ttk.Button(btn_frame, text="Use Selected", style="Accent.TButton", command=self._on_ok).grid(
            row=0, column=0, padx=5
        )
        ttk.Button(btn_frame, text="Cancel", style="Ghost.TButton", command=self._on_cancel).grid(
            row=0, column=1, padx=5
        )
        self.bind("<Return>", lambda e: self._on_ok())
        self.bind("<Escape>", lambda e: self._on_cancel())
        self.columnconfigure(0, weight=1)
        self.update_idletasks()
        if master is not None:
            x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
            y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{x}+{y}")
        self.lift()
        self.focus_force()
        self.attributes('-topmost', True)
        self.after(100, lambda: self.attributes('-topmost', False))
    def _find_partner_index(self, keywords):
        for i, p in enumerate(self.partners):
            pf = (p.get("partner_function", "") or "").lower()
            for kw in keywords:
                if kw.lower() in pf:
                    return i
        return None
    def _on_ok(self):
        try:
            name_idx = self.name_combo.current()
            addr_idx = self.address_combo.current()
            self.result_name_index = name_idx if name_idx >= 0 else None
            self.result_address_index = addr_idx if addr_idx >= 0 else None
        except Exception:
            self.result_name_index = None
            self.result_address_index = None
        self.destroy()
    def _on_cancel(self):
        self.result_name_index = None
        self.result_address_index = None
        self.destroy()
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
    def _choose_external_contact_if_needed(self):
        partners = self.values.get("_external_contacts") or []
        if len(partners) <= 1:
            if len(partners) == 1:
                p = partners[0]
                name = (p.get("name") or "").strip()
                addr = (p.get("address") or "").strip().replace(" / ", "\n")
                if name:
                    self.values["ir_name"] = name
                if addr:
                    self.values["ir_with_address"] = f"{name}\n{addr}".strip()
            return
        dlg = PartnerSelectionDialog(self, partners)
        self.wait_window(dlg)
        if dlg.result_name_index is None and dlg.result_address_index is None:
            return
        selected_name = ""
        selected_addr = ""
        if dlg.result_name_index is not None and 0 <= dlg.result_name_index < len(partners):
            name_partner = partners[dlg.result_name_index]
            selected_name = (name_partner.get("name") or "").strip()
            if selected_name:
                self.values["ir_name"] = selected_name
        if dlg.result_address_index is not None and 0 <= dlg.result_address_index < len(partners):
            addr_partner = partners[dlg.result_address_index]
            selected_addr = (addr_partner.get("address") or "").strip().replace(" / ", "\n")
        if not selected_name:
            selected_name = (self.values.get("ir_name") or "").strip()
        if not selected_addr:
            facility_block = (self.values.get("ir_with_address") or "").strip()
            lines = [ln.strip() for ln in facility_block.splitlines() if ln.strip()]
            if len(lines) >= 2:
                selected_addr = "\n".join(lines[1:]).strip()
            elif lines:
                selected_addr = lines[0]
        if selected_name and selected_addr:
            self.values["ir_with_address"] = f"{selected_name}\n{selected_addr}"
        elif selected_name:
            self.values["ir_with_address"] = selected_name
        elif selected_addr:
            self.values["ir_with_address"] = selected_addr
    def _build_step1(self):
        f = self.step1_frame
        ttk.Label(f, text="Enter GCH PE Number", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            f,
            text="We’ll pull initial reporter data, event description, product info, analyses, and investigations from GCH.",
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
        self._choose_external_contact_if_needed()
        ir_text = self.values.get("ir_with_address", "") or ""
        self.ir_text_widget.delete("1.0", "end")
        self.ir_text_widget.insert("1.0", ir_text)
        self.status_var.set("")
        self._bring_to_front()
        self._show_step(self.step2_frame, "Step 2 of 4 · Edit Initial Reporter / Address")
    def _build_step2(self):
        f = self.step2_frame
        ttk.Label(f, text="Initial Reporter & Facility Address", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            f,
            text="Review and adjust the initial reporter / facility block before it goes into the template.",
            style="CardText.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(5, 10))
        self.ir_text_widget = tk.Text(
            f,
            width=100,
            height=10,
            wrap="word",
            bg="#ffffff",
            fg=TEXT_DARK,
            insertbackground=TEXT_DARK,
            relief="solid",
            borderwidth=1,
        )
        self.ir_text_widget.grid(row=2, column=0, sticky="nsew", pady=(5, 10))
        btn_frame = ttk.Frame(f, style="Card.TFrame")
        btn_frame.grid(row=3, column=0, sticky="e")
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
        f.grid_rowconfigure(2, weight=1)
        f.grid_columnconfigure(0, weight=1)
    def on_ir_next(self):
        self.values["ir_with_address"] = self.ir_text_widget.get("1.0", "end-1c")
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
        self.ir_text_widget.config(state="normal")
        self.ir_text_widget.delete("1.0", "end")
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
